from flask import Flask, request, jsonify, send_from_directory, render_template
from flask_cors import CORS
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
from openai import OpenAI
from typing import List, Dict, Union
from mail import mail_main

app = Flask(__name__)
CORS(app)  # 这将允许所有域的跨域请求
# 初始化Flask应用
app = Flask(__name__, template_folder="templates", static_folder="static")

CORS(
    app,
    origins=["http://127.0.0.1:5001/"],
    methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
    supports_credentials=True,
    max_age=3600,
)


@app.after_request
def after_request(response):
    if "Connection" in response.headers:
        del response.headers["Connection"]
    response.headers["Connection"] = "keep-alive"
    return response


# OpenAI客户端配置
client = OpenAI(
    # base_url="https://uni-api.cstcloud.cn/v1",
    base_url = "https://xxx",
    api_key = "xxx"
)

# 配置参数
PRIORITY_CONTACTS = ["稿件接收", "实验室合作伙伴", "重要客户", "项目进展"]
OUTPUT_FORMATS = ["text", "markdown"]
OUTPUT_DIR = "output"  # 输出目录
EMAIL_DATA = "mail\\export_163mails.xlsx"  # 邮件数据路径


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
        output_format = data.get("output_format", "markdown")

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


# 传输email_list
@app.route("/generate_reports", methods=["POST"])
def generate_reports():
    try:
        if not request.json:
            return jsonify({"status": "error", "message": "无效的JSON请求"}), 400

        data = request.json
        email = data.get("email")
        cookie = data.get("email")
        prompt = data.get("prompt", "请分类这些邮件并标记优先级")

        # 获取邮件列表
        email_list = mail_main(cookie)

        # 判断是否收到“暂无未读邮件.”
        if isinstance(email_list, str) and email_list.strip() == "暂无未读邮件。":
            return jsonify(
                {
                    "status": "success",
                    "data": email_list,
                    "format": "text",
                    "emails_processed": 0,
                }
            )

        # 正常处理邮件数据
        output_format = "markdown"  # 默认格式
        result = ask(prompt, email_list, output_format)
        print(result)

        # 保存结果到文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"report_result_{timestamp}.md"
        result_path = os.path.join(OUTPUT_DIR, result_filename)

        with open(result_path, "w", encoding="utf-8") as f:
            f.write(result)

        # 返回响应
        return jsonify(
            {
                "status": "success",
                "data": result,
                "format": output_format,
                "emails_processed": len(email_list),
            }
        )

    except Exception as e:
        return (
            jsonify(
                {
                    "status": "error",
                    "message": f"处理失败: {str(e)}",
                    "error_type": type(e).__name__,
                }
            ),
            500,
        )


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

        # 生成所有格式的结果
        results = {
            "text": ask(prompt, emails, "text"),
            # "json": ask(prompt, emails, "json"),
            "markdown": ask(prompt, emails, "markdown"),
        }

        # 保存所有文件
        file_paths = {}

        # 1. 保存文本文件
        text_path = Path(output_dir) / f"{base_filename}.txt"
        with open(text_path, "w", encoding="utf-8") as f:
            f.write(results["text"])
        file_paths["text"] = str(text_path)

        # 2. 保存JSON文件
        # json_path = Path(output_dir) / f"{base_filename}.json"
        # with open(json_path, "w", encoding="utf-8") as f:
        #     json.dump(results["json"], f, ensure_ascii=False, indent=2)
        # file_paths["json"] = str(json_path)

        # 3. 保存Markdown文件
        md_path = Path(output_dir) / f"{base_filename}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(results["markdown"])
        file_paths["markdown"] = str(md_path)

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
    saved_files = [
        f
        for f in os.listdir(OUTPUT_DIR)
        if f.startswith("report_result_") and f.endswith(".md")
    ]
    if not saved_files:
        return None

    saved_files.sort(reverse=True, key=lambda x: x.split("_")[2] + x.split("_")[3])

    for filename in saved_files:
        file_path = os.path.join(OUTPUT_DIR, filename)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                file_content = f.read().strip()
                # 如果你还想做 email/prompt 匹配，可以解析旧文件名或保留部分元数据
                return file_content  # 返回纯 Markdown 内容
        except Exception as e:
            print(f"读取历史文件失败：{filename}，错误：{str(e)}")
    return None


def ask(
    prompt: str, email_list: List[Dict], output_format: str = "markdown"
) -> Union[str, Dict]:
    """
    处理邮件列表并返回指定格式结果

    :param prompt: 处理提示
    :param email_list: 邮件数据列表（格式如email.png所示）
    :param output_format: 输出格式（此处强制为markdown）
    :return: markdown格式字符串或包含错误信息的字典
    """
    try:
        system_prompt = """
                作为邮件处理助手，接下来请严格按以下模板生成邮件摘要。直接提取邮件信息填入模板：
                **发件人**: [原始发件人姓名和邮箱]
                **优先级**: [优先级说明]
                **分类标签**: [标签1] | [标签2]
                **处理建议**:
                1. [具体建议1]
                2. [具体建议2]
                3. [具体建议3]
                **附件状态**: [附件描述]

                ▌具体填充规则：
                1. **分类标签**：
                - 使用 emoji + 文字组合（例：📚 学术会议）
                - 根据内容选择 1-3 个标签，用竖线分隔
                - 备选标签库：
                    • 📚 学术会议 • 🗓️ 截止提醒 • 📬 邀请函 
                    • 📊 数据报告 • ❓ 问题咨询 • 📝 材料提交
                2. 处理建议：
                - 生成2-4条可操作建议
                - 必须包含研究方向匹配性检查
                - 必须包含资质/截止期验证
                3. 附件状态：
                - 有附件："检测到附件：共X个（示例：filename.pdf）"
                - 无附件："未检测到附件"
                """

        # 1. 预处理邮件数据
        user_input = f"邮件内容：\n"
        for idx, email in enumerate(email_list, 1):
            sender = email.get("发件人", "")
            subject = email.get("邮件标题", "无标题")
            content = email.get("正文摘要", "")
            # 强制转换为字符串，防止 list 类型报错
            if isinstance(subject, list):
                subject = " ".join(subject)
            if isinstance(content, list):
                content = " ".join(content)
            subject = str(subject)
            content = str(content)
            user_input += f"邮件 {idx}:\n"
            user_input += f"发件人: {sender}\n"
            user_input += f"标题: {subject}\n"
            user_input += f"正文摘要: {content}\n\n"

            # api
            response = client.chat.completions.create(
                model="deepseek-r1:671b",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_input},
                ],
                temperature=0.6,
                max_tokens=4096,
                stream=False,
            )

            suggestion = response.choices[0].message.content
            # 分类邮件
            # category = classify_email(subject, content, sender)
            # suggestion = generate_suggestion(priority, category)
            # processed_emails.append({
            #     "处理建议": suggestion,
            # })

        return suggestion

    except Exception as e:
        return {
            "status": "error",
            "message": f"邮件处理失败: {str(e)}",
            "error": str(e),
        }


def summarize_content(content: str, max_length: int = 100) -> str:
    """生成内容摘要"""
    content = " ".join(content.split())  # 去除多余空白
    if len(content) <= max_length:
        return content
    return content[:max_length] + "..."


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
