# ai_interface.py

import random
from http import HTTPStatus
import dashscope

dashscope.api_key = "sk-95ddb5205f614d35aeb98c0ca4f2f8f2"

def call_with_messages(content, report_style='formal'):
    # 根据 report_style 为内容添加不同风格的提示
    style_prompt = {
        "formal": "请使用正式且专业的语言撰写报告内容。",
        "concise": "请简洁明了地撰写报告内容，避免冗长。",
        "detailed": "请提供详细且深入的分析报告内容。",
    }
    
    # 获取用户选择的风格提示，默认为正式风格
    style_instruction = style_prompt.get(report_style, "请使用正式且专业的语言撰写报告内容。")
    
    # 将风格提示添加到消息内容
    messages = [
        {'role': 'user', 'content': f"{style_instruction} {content}"}
    ]
    
    try:
        response = dashscope.Generation.call(
            model="qwen-turbo",
            messages=messages,
            seed=random.randint(1, 10000),
            result_format='message'
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
