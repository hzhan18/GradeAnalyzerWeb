# data_processing.py

import os
import pandas as pd
from flask import session
from report_generation import generate_word_report
from plotting import plot_distribution

def extract_text_between(text, start_str, end_str):
    try:
        start = text.index(start_str) + len(start_str)
        end = text.index(end_str, start)
        return text[start:end].strip()
    except ValueError:
        return ""

def extract_text_after(text, start_str):
    try:
        start = text.index(start_str) + len(start_str)
        return text[start:].strip()
    except ValueError:
        return ""

def extract_text_in_parentheses(text):
    try:
        start = text.index('(') + 1
        end = text.index(')', start)
        return text[start:end].strip()
    except ValueError:
        return ""

def detect_format(df_preview):
    for idx, row in df_preview.iterrows():
        if "序号" in row.values:
            return "format_1", idx
        elif "编号" in row.values:
            return "format_2", idx
    return "unknown", None

def calculate_statistics(df, column):
    total_count = df.shape[0]
    stats = {
        '总人数': total_count,
        '最高分': df[column].max(),
        '最低分': df[column].min(),
        '平均分': "{:.2f}".format(df[column].mean()),
    }
    score_ranges = [(i, i+10) for i in range(0, 90, 10)]
    score_ranges.append((90, 100))
    distribution_text = {}

    for lower, upper in score_ranges:
        if upper == 100:
            count = df[column].between(lower, upper, inclusive='both').sum()
        else:
            count = df[column].between(lower, upper, inclusive='left').sum()

        percentage = (count / total_count * 100) if total_count > 0 else 0
        distribution_text[f'{lower}-{upper}分'] = {'人数': count, '占比': round(percentage, 2)}

    plot_ranges = [(0, 60), (60, 70), (70, 80), (80, 90), (90, 100)]
    distribution_plot = {}
    for lower, upper in plot_ranges:
        if upper == 100:
            count = df[column].between(lower, upper, inclusive='both').sum()
        else:
            count = df[column].between(lower, upper, inclusive='left').sum()
        percentage = (count / total_count * 100) if total_count > 0 else 0
        distribution_plot[f'{lower}-{upper}分'] = {'人数': count, '占比': round(percentage, 2)}
    
    return stats, distribution_text, distribution_plot

def run_report_generation(file_path, class_name1, class_name2, session_id):
    if not os.path.exists(file_path):
        return {"status": "error", "message": "文件不存在，请检查文件名和路径。"}

    excel_dir = os.path.dirname(file_path)
    output_dir = os.path.splitext(os.path.basename(file_path))[0].strip()
    output_path = os.path.join(excel_dir, output_dir)

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # Initialize session progress tracking
    session[session_id] = {"progress": 0, "status": "正在读取Excel文件..."}
    
    xl = pd.ExcelFile(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0].strip()

    df_preview = xl.parse(xl.sheet_names[0], nrows=6)

    # Detect file format
    file_format, start_row = detect_format(df_preview)

    # Initialize variables for semester info and course name
    semester_info = ""
    course_name = ""

    if file_format == "format_1":
        sheet = xl.parse(xl.sheet_names[0], header=None)
        cell_A1 = sheet.iloc[0, 0]
        semester_info = extract_text_between(cell_A1, "厦门理工学院", "成绩登记表")
        cell_A3 = sheet.iloc[2, 0]
        course_name = extract_text_after(cell_A3, "课程名称：")
        df = xl.parse(xl.sheet_names[0], skiprows=start_row + 1)
        df.columns = ['序号', '姓名', '学号', 'Unnamed: 3', '平时', '实验', '期末', '总评', '备注']
        id_column = '序号'
    elif file_format == "format_2":
        sheet = xl.parse(xl.sheet_names[0], header=None)
        cell_A2 = sheet.iloc[1, 0]
        semester_info = extract_text_in_parentheses(cell_A2)
        cell_E3 = sheet.iloc[2, 4]
        course_name = cell_E3.strip()
        df = xl.parse(xl.sheet_names[0], skiprows=start_row + 1)
        df.columns = ['编号', '学号', '姓名', 'Unnamed: 3', 'Unnamed: 4', '平时', '实验', '期末', '总评', '备注']
        id_column = '编号'
    else:
        return {"status": "error", "message": "无法识别的文件格式。"}

    # Clean data
    df_cleaned = df[pd.to_numeric(df[id_column], errors='coerce').notnull()].copy()
    score_columns = ['平时', '实验', '期末', '总评']
    for column in score_columns:
        df_cleaned[column] = pd.to_numeric(df_cleaned[column], errors='coerce').fillna(0)

    total_students = df_cleaned.shape[0]

    # Update session progress
    session[session_id]["progress"] = 20
    session[session_id]["status"] = "正在计算统计数据..."

    all_scores_data = []
    for i, score_type in enumerate(score_columns):
        stats, distribution_text, distribution_plot = calculate_statistics(df_cleaned, score_type)
        plot_title = f"{score_type}成绩分布"
        plot_file_name = os.path.join(output_path, f"{base_name}_{plot_title}.png")
        plot_distribution(distribution_plot, plot_title, plot_file_name)
        all_scores_data.append({
            'score_type': score_type,
            'stats': stats,
            'distribution_text': distribution_text,
            'distribution_plot': distribution_plot,
            'plot_file_name': plot_file_name
        })
        # Update session progress
        progress_increment = int(60 / len(score_columns))
        session[session_id]["progress"] += progress_increment
        session[session_id]["status"] = f"正在处理 {score_type} 成绩数据..."

    # Generate Word report
    report_path = os.path.join(output_path, f"{base_name}_成绩信息汇总.docx")
    generate_word_report(
        base_name,
        report_path,
        all_scores_data,
        output_path,
        semester_info,
        course_name,
        total_students,
        class_name1,
        class_name2
    )

    session[session_id]["progress"] = 100
    session[session_id]["status"] = "完成"

    return {
        "status": "success",
        "report_path": report_path,
        "output_path": output_path,
        "session_id": session_id
    }
