# ai_interface.py

import random
from http import HTTPStatus
import dashscope
import logging

dashscope.api_key = "sk-95ddb5205f614d35aeb98c0ca4f2f8f2"

def call_with_messages(content, report_style='formal'):
    # 根据 report_style 为内容添加不同风格的提示
    style_prompt = {
        "formal": (
            "请使用正式且专业的语言风格撰写成绩数据分析。"
            "逻辑清晰。避免仅展示分数段数据。"
            "请减少转接词的使用（例如首先、其次、最后、综上所述、总的来说、此外、值得注意的是）。"
            "确保数据分析专业、严谨。"
        ),
        "concise": (
            "请使用简明扼要的语言风格撰写成绩数据分析。"
            "内容应简洁、清晰，重点突出，避免冗余信息。"
            "除非学生成绩数据特别需要强调，否则请避免展示过多的分数段数据。"
            "请减少转接词的使用（例如首先、其次、最后、综上所述、总的来说、此外、值得注意的是）。"
            "确保内容简洁、清晰且易于理解，突出主要观点和关键数据。"
        ),
        "detailed": (
            "请较为详尽的撰写成绩数据分析。"
            "探讨各分数段的分布情况，并结合具体数据进行说明。"
            "请减少转接词的使用（例如首先、其次、最后、综上所述、总的来说、此外、值得注意的是）。"
            "提供较为充分的数据支持和分析，必要时引用具体案例。"
        ),
    }
    
    # 获取用户选择的风格提示，默认为正式风格
    style_instruction = style_prompt.get(report_style, style_prompt["formal"])
    
    
    # 将风格提示添加到消息内容
    messages = [
        {'role': 'user', 'content': f"{style_instruction} {content}"}
    ]
    
    # 添加日志，输出当前的 prompt 和 style
    logging.info(f"使用的报告风格: {report_style}")
    logging.info(f"生成的 prompt: {messages[0]['content']}")
    
    try:

        response = dashscope.Generation.call(
            model="qwen-turbo",
            messages=messages,
            seed=random.randint(1, 10000),
            result_format='message',
            temperature=0.5,  # 控制生成内容的随机性
            top_p=1
        )
        if response.status_code == HTTPStatus.OK:
            return response['output']['choices'][0]['message']['content']
        
        error_message = (
            f'请求失败: {response.request_id}, 状态码: {response.status_code}, '
            f'错误代码: {response.code}, 错误信息: {response.message}'
        )
        print(error_message)
        return "抱歉，无法完成分析。请稍后再试。"
    except Exception as e:
        print(f"发生错误: {str(e)}")
        return "抱歉，处理您的请求时出错。请稍后再试。"
