from openai import OpenAI
import pandas as pd
import os
import json
from datetime import datetime
from pathlib import Path
from flask import Flask, request, jsonify
from typing import List, Dict, Union
app = Flask(__name__)

client = OpenAI(
# deepseek
base_url="https://uni-api.cstcloud.cn/v1",
api_key="e6f0f5790dd3c0050e0bba50d85204a5da3f0566df60014b21776f0eb680167f"
)

# 配置参数
PRIORITY_CONTACTS = ["合作期刊编辑", "实验室合作伙伴", "重要客户"]  # 重要联系人列表
OUTPUT_FORMATS = ["text", "json", "html", "markdown"]  # 支持的输出格式

def ask(prompt: str, triples_excel: List[Dict], output_format: str = "text") -> Union[str, Dict]:
    # 预处理邮件数据，标记优先级
    processed_emails = []
    for email in triples_excel:
        # 自动判断优先级
        priority = "normal"
        if "紧急" in email.get("邮件标题", "") or "紧急" in email.get("正文", ""):
            priority = "high"
        elif any(contact in email.get("发件人", "") for contact in PRIORITY_CONTACTS):
            priority = "high"
        
        # 添加优先级标记
        processed_email = email.copy()
        processed_email["优先级"] = priority
        processed_emails.append(processed_email)
    
    # 系统提示词
    system_prompt = (
        "你是一个专业的邮箱处理助手，请按照以下要求处理邮件：\n"
        "1. 自动标记来自重要联系人({})或标有'紧急'的邮件为高优先级\n"
        "2. 根据内容相关性进一步分类\n"
        "3. 提供清晰的处理建议\n"
        "4. 按指定格式({})输出结果".format(", ".join(PRIORITY_CONTACTS), ", ".join(OUTPUT_FORMATS))
    )
    
    response = client.chat.completions.create(
        model="deepseek-r1:671b",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"待处理邮件：{json.dumps(processed_emails, ensure_ascii=False)}"},
            {"role": "user", "content": f"{prompt}\n请使用{output_format}格式回复"},
        ],
        temperature=0.6,
        max_tokens=4096,
        stream=False
    )
    
    result = response.choices[0].message.content
    
    # 格式后处理
    if output_format == "json":
        try:
            return json.loads(result)
        except json.JSONDecodeError:
            return {"error": "JSON格式转换失败", "raw_response": result}
    return result

def generate_email_report(email_excel_path: str, prompt: str, output_format: str = "text") -> Union[str, Dict]:
    """独立函数，可在非Web环境中使用"""
    try:
        df = pd.read_excel(email_excel_path)
        emails = df.to_dict("records")
        return ask(prompt, emails, output_format)
    except Exception as e:
        return f"Error: {str(e)}"
    
@app.route("/process_emails", methods=["POST"])
def process_emails():
    """处理邮件的主端点，支持多种输出格式"""
    try:
        # 获取请求参数
        data = request.get_json()
        email_excel_path = data.get("email_excel")
        prompt = data.get("prompt", "请分类这些邮件并标记优先级")
        output_format = data.get("output_format", "text")
        
        # 验证输出格式
        if output_format not in OUTPUT_FORMATS:
            return jsonify({
                "status": "error",
                "message": f"不支持的输出格式，请选择其中之一: {', '.join(OUTPUT_FORMATS)}"
            }), 400
        
        # 读取Excel文件
        df = pd.read_excel(email_excel_path)
        emails = df.to_dict("records")
        
        # 处理邮件
        result = ask(prompt, emails, output_format)
        
        # 返回响应
        return jsonify({
            "status": "success",
            "data": result,
            "format": output_format
        })
    
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

def generate_all_reports(email_excel_path: str, prompt: str, output_dir: str = "output") -> dict:
    """
    生成所有格式的报告并保存为文件
    返回生成的文件路径字典
    """
    # 确保输出目录存在
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # 生成时间戳用于文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = f"email_report_{timestamp}"
    
    # 读取原始邮件数据
    try:
        df = pd.read_excel(email_excel_path)
        emails = df.to_dict("records")
        scl_pdf = ""  # 可替换为实际文档内容
        
        # 生成所有格式的结果
        results = {
            "text": ask(prompt, emails, "text"),
            "json": ask(prompt, emails,  "json"),
            "markdown": ask(prompt, emails, "markdown"),
        }
        
        # 生成Excel格式结果（特殊处理）
        excel_data = process_for_excel(emails, results['json'])
        results['excel'] = excel_data
        
        # 保存所有文件
        file_paths = {}
        
        # 1. 保存文本文件
        text_path = Path(output_dir) / f"{base_filename}.txt"
        with open(text_path, 'w', encoding='utf-8') as f:
            f.write(results['text'])
        file_paths['text'] = str(text_path)
        
        # 2. 保存JSON文件
        json_path = Path(output_dir) / f"{base_filename}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(results['json'], f, ensure_ascii=False, indent=2)
        file_paths['json'] = str(json_path)
        
        # 3. 保存Markdown文件
        md_path = Path(output_dir) / f"{base_filename}.md"
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(results['markdown'])
        file_paths['markdown'] = str(md_path)
        
        # 4. 保存Excel文件
        excel_path = Path(output_dir) / f"{base_filename}.xlsx"
        results['excel'].to_excel(excel_path, index=False)
        file_paths['excel'] = str(excel_path)
        
        return {"status": "success", "files": file_paths}
    
    except Exception as e:
        return {"status": "error", "message": str(e)}

def process_for_excel(original_emails, json_result) -> pd.DataFrame:
    """
    准备Excel格式的输出数据
    返回DataFrame对象
    
    参数:
        original_emails: 原始邮件数据列表
        json_result: 可能是字符串或已解析的JSON对象
    """
    try:
        # 确保json_result是字典格式
        if isinstance(json_result, str):
            processed_data = json.loads(json_result)
        else:
            processed_data = json_result
        
        # 如果结果是包含'data'字段的响应结构
        if isinstance(processed_data, dict) and 'data' in processed_data:
            processed_data = processed_data['data']
        
        # 验证处理后的数据格式
        if not isinstance(processed_data, (list, dict)):
            raise ValueError("处理后的数据格式不正确")
        
        # 如果是字典形式，转换为列表
        if isinstance(processed_data, dict):
            processed_data = [processed_data]
        
        # 合并原始数据和处理结果
        excel_data = []
        for i, original in enumerate(original_emails):
            # 获取对应的处理结果
            processed = processed_data[i] if i < len(processed_data) else {}
            
            record = {
                **original,
                "AI处理_优先级": processed.get("优先级", "未标记"),
                "AI处理_分类": processed.get("分类", "未分类"),
                "AI处理_建议": processed.get("处理建议", "无"),
                "AI处理_摘要": processed.get("内容摘要", ""),
            }
            excel_data.append(record)
        
        return pd.DataFrame(excel_data)
    
    except Exception as e:
        print(f"Excel处理错误: {str(e)}")
        # 返回只包含原始数据的DataFrame作为降级方案
        return pd.DataFrame(original_emails)
    
if __name__ == "__main__":
    email_data = "D:\\课程作业（研）\\高级人工智能\\export_163mails.xlsx"
    output_directory = "D:\\课程作业（研）\\高级人工智能\\output"
    user_prompt = (
        "请完成以下任务：\n"
        "1. 按优先级和主题分类所有邮件\n"
        "2. 识别需要立即处理的邮件\n"
        "3. 为每类邮件提供处理建议\n"
        "4. 统计各类邮件的数量"
    )
    
    # 获取不同格式的结果
    text_result = generate_email_report(email_data, user_prompt, "text")
    json_result = generate_email_report(email_data, user_prompt, "json")
    markdown_result = generate_email_report(email_data, user_prompt, "markdown")
    # app.run(host='0.0.0.0', port=5000, debug=True)
    
    # print("文本格式结果:\n", text_result)
    # print("\nJSON格式结果:\n", json.dumps(json_result, indent=2, ensure_ascii=False))
    # print("\nMarkdown格式结果:\n", markdown_result)
    # 保存为文件
    # 生成所有报告
    result = generate_all_reports(email_data, user_prompt, output_directory)
    
    if result['status'] == 'success':
        print("成功生成所有报告文件：")
        for format_type, path in result['files'].items():
            print(f"{format_type.upper()}格式: {path}")
    else:
        print("处理失败:", result['message'])
