import os
import pandas as pd
from report_generation import generate_word_report
from plotting import plot_distribution
from flask import session

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

# Assuming session_id is passed into this function for tracking progress
def run_report_generation(file_path, class_name1, class_name2, session_id):
    if not os.path.exists(file_path):
        return {"status": "error", "message": "File does not exist. Please check the file name and path."}

    excel_dir = os.path.dirname(file_path)
    output_dir = os.path.splitext(os.path.basename(file_path))[0].strip()
    output_path = os.path.join(excel_dir, output_dir)

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # Initialize session progress tracking
    session[session_id] = {"progress": 0, "status": "Initializing report generation"}

    # Load the file and process
    xl = pd.ExcelFile(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0].strip()

    # Example of updating session progress within this function
    session[session_id]["progress"] = 20
    session[session_id]["status"] = "Reading Excel file"

    # (Add the rest of the report generation logic here)
    # After each major step, update the session with progress, e.g.:
    # session[session_id]["progress"] += 20
    # session[session_id]["status"] = "Processing XYZ"

    # Finalize
    session[session_id]["progress"] = 100
    session[session_id]["status"] = "Completed"

    # Example return
    report_path = os.path.join(output_path, f"{base_name}_Summary.docx")
    print("Report path generated:", report_path)
    return {"status": "success", "report_path": report_path, "session_id": session_id}
    

def calculate_statistics(df, column):
    total_count = df.shape[0]
    stats = {
        '最高分': df[column].max(),
        '最低分': df[column].min(),
        '平均分': round(df[column].mean(), 2),
        '总人数': total_count
    }
    
    # Prepare distribution for plotting and table generation
    distribution_plot = {}
    distribution_text = {}
    for lower, upper in [(0, 60), (60, 70), (70, 80), (80, 90), (90, 100)]:
        count = df[column].between(lower, upper, inclusive='both' if upper == 100 else 'left').sum()
        percentage = (count / total_count) * 100 if total_count > 0 else 0
        range_key = f'{lower}-{upper}'
        distribution_plot[range_key] = {'人数': count}  # Ensuring distribution_plot has the '人数' key
        distribution_text[range_key] = {'人数': count, '占比': round(percentage, 2)}

    return stats, distribution_plot, distribution_text


    all_scores_data = []
    for i, score_type in enumerate(score_columns):
        stats, distribution_plot, distribution_text = calculate_statistics(df_cleaned, score_type)
        plot_title = f"{score_type}成绩分布"
        plot_file_name = os.path.join(output_path, f"{base_name}_{plot_title}.png")
        plot_distribution(distribution_plot, plot_title, plot_file_name)
        all_scores_data.append({
            'score_type': score_type,
            'stats': stats,
            'distribution_text': distribution_text,  # Include distribution_text here
            'plot_file_name': plot_file_name
        })
        # Update session progress
        session[session_id]["progress"] = (i + 1) * 20
        session[session_id]["status"] = f"Processing {score_type} data"

    report_path = os.path.join(output_path, f"{base_name}_Summary.docx")
    generate_word_report(
        base_name, report_path, all_scores_data, output_path,
        semester_info, course_name, total_students, class_name1, class_name2
    )

    session[session_id]["progress"] = 100
    session[session_id]["status"] = "Completed"
    
    return {"status": "success", "report_path": report_path, "output_path": output_path, "session_id": session_id}
