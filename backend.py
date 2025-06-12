from flask import Flask, request, jsonify, send_from_directory, render_template
from flask_cors import CORS
import os
import json
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

CORS(app, origins=["http://192.168.43.157:5001"],
    methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
    supports_credentials=True,
    max_age=3600)

@app.after_request
def after_request(response):
    if 'Connection' in response.headers:
        del response.headers['Connection']
    response.headers['Connection'] = 'keep-alive'
    return response

# OpenAI客户端配置
client = OpenAI(
    base_url="https://uni-api.cstcloud.cn/v1",
    api_key="e6f0f5790dd3c0050e0bba50d85204a5da3f0566df60014b21776f0eb680167f",
)

# 配置参数
PRIORITY_CONTACTS = ["合作期刊编辑", "实验室合作伙伴", "重要客户"]
OUTPUT_FORMATS = ["text", "markdown"]
OUTPUT_DIR = "D:\\课程作业（研）\\高级人工智能\\email_agent\\AdvancedAI_Project-main\\output"#"C:\\Users\\lusia\\Desktop\\email_agent\\output"  # 输出目录
EMAIL_DATA = (
    "mail\\export_163mails.xlsx"  # 邮件数据路径
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
            return jsonify({
                "status": "success",
                "data": email_list,
                "format": "text",
                "emails_processed": 0
            })

        # 正常处理邮件数据
        output_format = "markdown"  # 默认格式
        result = ask(prompt, email_list, output_format)
        print(result)

        # 保存结果到文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"report_result_{timestamp}.md"
        result_path = os.path.join(OUTPUT_DIR, result_filename)

        with open(result_path, "w", encoding="utf-8") as f:
            f.write(json.dumps({
                "email": email,
                "prompt": prompt,
                "result": result
            }, ensure_ascii=False, indent=2))

        # 返回响应
        return jsonify({
            "status": "success",
            "data": result,
            "format": output_format,
            "emails_processed": len(email_list)
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"处理失败: {str(e)}",
            "error_type": type(e).__name__
        }), 500


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
    """
    从 OUTPUT_DIR 中读取已保存的 report_result_*.txt 文件，筛选匹配的 result
    :param OUTPUT_DIR: 结果保存目录
    :param target_email: 请求中的邮箱
    :param target_prompt: 请求中的 prompt
    :return: 匹配的 result（字典）或 None（无匹配/读取失败）
    """
    # 筛选目录下所有符合命名规则的文件
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

def ask(prompt: str, email_list: List[Dict], output_format: str = "markdown") -> Union[str, Dict]:
    """
    处理邮件列表并返回指定格式结果
    
    :param prompt: 处理提示
    :param email_list: 邮件数据列表（格式如email.png所示）
    :param output_format: 输出格式（此处强制为markdown）
    :return: markdown格式字符串或包含错误信息的字典
    """
    try:
        # 1. 预处理邮件数据
        processed_emails = []
        for email in email_list:
            # 解析邮件数据
            sender = email.get("发件人", "")
            subject = email.get("邮件标题", "无标题")
            content = email.get("正文", "")
            has_attachment = email.get("附件", False)
            # 强制转换为字符串，防止 list 类型报错
            if isinstance(subject, list):
                subject = " ".join(subject)
            if isinstance(content, list):
                content = " ".join(content)
            subject = str(subject)
            content = str(content)

            # 确定优先级
            priority = "normal"
            if "紧急" in subject or "紧急" in content:
                priority = "high"
            elif any(contact in sender for contact in PRIORITY_CONTACTS):
                priority = "high"
            
            # 分类邮件
            category = classify_email(subject, content, sender)
            
            # 生成处理建议
            suggestion = generate_suggestion(priority, category)
            
            processed_emails.append({
                "发件人": sender,
                "邮件标题": subject,
                "优先级": priority,
                "分类": category,
                "处理建议": suggestion,
                "正文摘要": summarize_content(content),
                "附件": "有" if has_attachment else "无"
            })

        # 2. 生成markdown格式结果
        markdown_output = f"# 邮件处理报告\n\n**处理提示**: {prompt}\n\n"
        markdown_output += f"**处理邮件总数**: {len(processed_emails)}\n\n"
        
        for idx, email in enumerate(processed_emails, 1):
            markdown_output += f"## 邮件 {idx}: {email['邮件标题']}\n"
            markdown_output += f"- **发件人**: {email['发件人']}\n"
            markdown_output += f"- **优先级**: {email['优先级']}\n"
            markdown_output += f"- **分类**: {email['分类']}\n"
            markdown_output += f"- **附件**: {email['附件']}\n"
            markdown_output += f"- **处理建议**: {email['处理建议']}\n"
            markdown_output += f"\n**正文摘要**:\n{email['正文摘要']}\n\n"
            markdown_output += "---\n\n"
        
        return markdown_output
        
    except Exception as e:
        return {
            "status": "error",
            "message": f"邮件处理失败: {str(e)}",
            "error": str(e)
        }


def classify_email(subject: str, content: str, sender: str) -> str:
    """邮件分类逻辑"""
    subject = subject.lower()
    content = content.lower()
    
    if "会议" in subject or "meeting" in subject:
        return "会议相关"
    elif "报告" in subject or "report" in subject:
        return "工作报告"
    elif "期刊" in subject or "paper" in subject or "call for papers" in content:
        return "学术相关"
    elif "ucas.ac.cn" in sender:
        return "学校邮件"
    elif "618" in subject or "促销" in subject or "天猫" in subject or "淘宝" in subject or "京东" in subject or "拼多多" in subject or "唯品会" in subject or "苏宁易购" in subject:
        return "商业广告"
    else:
        return "其他"


def generate_suggestion(priority: str, category: str) -> str:
    """生成处理建议"""
    if priority == "high":
        return "立即处理，24小时内回复"
    elif category == "会议相关":
        return "添加到日历并回复确认"
    elif category == "工作报告":
        return "阅读并归档"
    elif category == "学术相关":
        return "转发给相关研究人员并跟进"
    elif category == "学校邮件":
        return "学校邮件，请仔细阅读"
    else:
        return "标准处理，3个工作日内回复"


def summarize_content(content: str, max_length: int = 100) -> str:
    """生成内容摘要"""
    content = ' '.join(content.split())  # 去除多余空白
    if len(content) <= max_length:
        return content
    return content[:max_length] + "..."


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
