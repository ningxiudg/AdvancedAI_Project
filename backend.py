from flask import Flask, request, jsonify, send_from_directory, render_template
from flask_cors import CORS
import os
import json
import pandas as pd
from datetime import datetime
from pathlib import Path
from openai import OpenAI
from typing import List, Dict, Union

app = Flask(__name__)
CORS(app)  # 这将允许所有域的跨域请求
# 初始化Flask应用
app = Flask(__name__, template_folder="templates", static_folder="static")

# OpenAI客户端配置
client = OpenAI(
    base_url="https://uni-api.cstcloud.cn/v1",
    api_key="e6f0f5790dd3c0050e0bba50d85204a5da3f0566df60014b21776f0eb680167f",
)

# 配置参数
PRIORITY_CONTACTS = ["合作期刊编辑", "实验室合作伙伴", "重要客户"]
OUTPUT_FORMATS = ["text", "json", "html", "markdown"]
OUTPUT_DIR = "C:\\Users\\lusia\\Desktop\\email_agent\\output"  # 输出目录
EMAIL_DATA = (
    "C:\\Users\\lusia\\Desktop\\email_agent\\export_163mails.xlsx"  # 邮件数据路径
)


# 1. 创建前端路由
@app.route("/")
def index():
    """渲染主页面"""
    return render_template("index.html")  # 确保有templates/index.html文件


# 2. 处理邮件请求的路由
@app.route("/process_emails", methods=["POST"])
def process_emails():
    """处理邮件的主端点，支持多种输出格式"""
    try:
        # 获取请求参数
        data = request.get_json()
        prompt = data.get("prompt", "请分类这些邮件并标记优先级")
        output_format = data.get("output_format", "text")

        # 验证输出格式
        if output_format not in OUTPUT_FORMATS:
            return (
                jsonify(
                    {
                        "status": "error",
                        "message": f"不支持的输出格式，请选择其中之一: {', '.join(OUTPUT_FORMATS)}",
                    }
                ),
                400,
            )

        # 读取固定路径的Excel文件
        df = pd.read_excel(EMAIL_DATA)
        emails = df.to_dict("records")

        # 处理邮件
        result = ask(prompt, emails, output_format)

        # 返回响应
        return jsonify({"status": "success", "data": result, "format": output_format})

    except Exception as e:
        print(f"Error processing emails: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/generate_reports", methods=["POST"])
def generate_reports_route():
    try:
        data = request.json
        email = data.get("email", "")
        prompt = data.get("prompt", "请分类这些邮件并标记优先级")

        saved_result = load_saved_result(OUTPUT_DIR, email, prompt)
        if saved_result is not None:
            result = saved_result  # 直接使用历史结果
        else:
            # 无历史结果，调用核心函数生成新结果
            result = generate_all_reports(EMAIL_DATA, prompt, OUTPUT_DIR)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"report_result_{timestamp}.txt"
        result_path = os.path.join(OUTPUT_DIR, result_filename)
        if result.get("status") == "error":
            return jsonify(result), 500
        with open(result_path, "w", encoding="utf-8") as f:
            f.write(
                json.dumps(
                    {"email": email, "prompt": prompt, "result": result},
                    ensure_ascii=False,
                    indent=2,
                )
            )

        # 添加邮箱信息到返回结果
        result["email"] = email
        return jsonify(result)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/get_txt_content/<filename>")
def get_txt_content(filename):
    try:
        if ".." in filename or filename.startswith("/"):
            return jsonify({"status": "error", "message": "无效的文件路径"}), 400

        file_path = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(file_path):
            return jsonify({"status": "error", "message": "文件不存在"}), 404

        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()

        return jsonify({"status": "success", "filename": filename, "content": content})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# 3. 文件下载路由
@app.route("/download/<filename>")
def download_file(filename):
    """提供生成文件的下载"""
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


def generate_all_reports(
    email_excel_path: str, prompt: str, output_dir: str = "output"
) -> dict:
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
            "json": ask(prompt, emails, "json"),
            "markdown": ask(prompt, emails, "markdown"),
        }

        # 生成Excel格式结果（特殊处理）
        excel_data = process_for_excel(emails, results["json"])
        results["excel"] = excel_data

        # 保存所有文件
        file_paths = {}

        # 1. 保存文本文件
        text_path = Path(output_dir) / f"{base_filename}.txt"
        with open(text_path, "w", encoding="utf-8") as f:
            f.write(results["text"])
        file_paths["text"] = str(text_path)

        # 2. 保存JSON文件
        json_path = Path(output_dir) / f"{base_filename}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(results["json"], f, ensure_ascii=False, indent=2)
        file_paths["json"] = str(json_path)

        # 3. 保存Markdown文件
        md_path = Path(output_dir) / f"{base_filename}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(results["markdown"])
        file_paths["markdown"] = str(md_path)

        # 4. 保存Excel文件
        excel_path = Path(output_dir) / f"{base_filename}.xlsx"
        results["excel"].to_excel(excel_path, index=False)
        file_paths["excel"] = str(excel_path)

        text_content = ""
        try:
            with open(text_path, "r", encoding="utf-8") as f:
                text_content = f.read()
        except Exception as e:
            print(f"读取文本文件失败: {str(e)}")
        return {
            "status": "success",
            "files": file_paths,
            "text_content": text_content,  # 添加文本内容到返回值
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}


def load_saved_result(OUTPUT_DIR, target_email, target_prompt):
    """
    从 OUTPUT_DIR 中读取已保存的 report_result_*.txt 文件，筛选匹配的 result
    :param OUTPUT_DIR: 结果保存目录
    :param target_email: 请求中的邮箱
    :param target_prompt: 请求中的 prompt
    :return: 匹配的 result（字典）或 None（无匹配/读取失败）
    """
    # 筛选目录下所有符合命名规则的文件（report_result_时间戳.txt）
    saved_files = [
        f
        for f in os.listdir(OUTPUT_DIR)
        if f.startswith("report_result_") and f.endswith(".txt")
    ]
    if not saved_files:
        return None  # 无历史文件

    # 按时间戳倒序排序（优先取最新文件）
    saved_files.sort(
        reverse=True, key=lambda x: x.split("_")[2] + x.split("_")[3]
    )  # 按时间戳排序

    for filename in saved_files:
        file_path = os.path.join(OUTPUT_DIR, filename)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                file_content = f.read()
                data = json.loads(file_content)  # 解析文件内容

                # 检查邮箱和 prompt 是否匹配（可根据业务调整匹配逻辑）
                if (
                    data.get("email") == target_email
                    and data.get("prompt") == target_prompt
                ):
                    return data.get("result")  # 返回保存的 result
        except Exception as e:
            print(f"读取历史文件失败：{filename}，错误：{str(e)}")
    return None  # 无匹配文件或读取失败


def generate_email_report(
    email_excel_path: str, prompt: str, output_format: str = "text"
) -> Union[str, Dict]:
    """独立函数，可在非Web环境中使用"""
    try:
        df = pd.read_excel(email_excel_path)
        emails = df.to_dict("records")
        return ask(prompt, emails, output_format)
    except Exception as e:
        return f"Error: {str(e)}"


def ask(
    prompt: str, triples_excel: List[Dict], output_format: str = "text"
) -> Union[str, Dict]:
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
        "4. 按指定格式({})输出结果".format(
            ", ".join(PRIORITY_CONTACTS), ", ".join(OUTPUT_FORMATS)
        )
    )

    response = client.chat.completions.create(
        model="deepseek-r1:671b",
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"待处理邮件：{json.dumps(processed_emails, ensure_ascii=False)}",
            },
            {"role": "user", "content": f"{prompt}\n请使用{output_format}格式回复"},
        ],
        temperature=0.6,
        max_tokens=4096,
        stream=False,
    )

    result = response.choices[0].message.content

    # 格式后处理
    if output_format == "json":
        try:
            return json.loads(result)
        except json.JSONDecodeError:
            return {"error": "JSON格式转换失败", "raw_response": result}
    return result


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
        if isinstance(processed_data, dict) and "data" in processed_data:
            processed_data = processed_data["data"]

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
    app.run(host="0.0.0.0", port=5001, debug=True)
