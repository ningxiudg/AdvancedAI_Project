from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import pyperclip
import pandas as pd
import os


# 打开浏览器 免登录 进入邮箱首页
def open_browser_with_cookies(cookie):
    driver_path = "actual driver path"
    # driver_path = "xxx/xxx/chromedriver.exe"
    service = Service(executable_path=driver_path)

    options = webdriver.ChromeOptions()
    base_dir = os.getcwd()
    relative_download_dir = "export"
    download_dir = os.path.join(base_dir, relative_download_dir)
    os.makedirs(download_dir, exist_ok=True)
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://mail.163.com/")
    cookie_str = cookie
    for cookie in cookie_str.split(";"):
        cookie = cookie.strip()
        if "=" in cookie:
            name, value = cookie.split("=", 1)
            driver.add_cookie({"name": name, "value": value, "domain": ".163.com"})
    driver.refresh()
    driver.maximize_window()
    return driver


# 进入未读邮件列表
def click_to_mail_list(driver):
    not_read = driver.find_element(By.ID, "GF0")
    not_read.click()
    print("click GF0[未读消息] success")
    time.sleep(5)


# 点击每封邮件
def click_to_mail(driver):
    driver.refresh()
    time.sleep(5)
    emails = driver.find_elements(By.CLASS_NAME, "dP0")
    if not emails:
        return False
    emails[0].click()
    return True


# 保存邮件信息至字典
"""
    邮件标题
    发件人
    收件人
    时间
    附件
    正文
"""


def save_mail(driver, index, mail_list):
    driver.refresh()
    time.sleep(5)
    print(f"正在处理第{index}封邮件...")
    dict = {}

    # 发件人 收件人 时间
    infos = driver.find_elements(By.CLASS_NAME, "ig0")
    sender, receiver, t = cope_srt(infos[:3])
    dict["发件人"] = sender
    dict["收件人"] = receiver
    dict["时间"] = t

    # 附件
    for info in infos:
        text = info.text.strip()
        if "附件" in text:
            filename = cope_attachment(driver)
            dict["附件"] = filename
        else:
            dict["附件"] = False

    # 邮件标题
    title = driver.find_element(By.XPATH, "//h1[@title='邮件标题']")
    title = title.text.split("\n")
    dict["邮件标题"] = title

    # 正文
    html = cope_html(driver)
    dict["正文"] = html

    print(dict)
    mail_list.append(dict)
    driver.back()


# 处理 发件人 收件人 时间 字段
def cope_srt(infos):
    [sender, receiver, t] = infos[:3]
    sender = sender.text.split("\n")
    sender = sender[1].strip()
    receiver = receiver.text.split("\n")
    receiver = receiver[1].strip()
    t = t.text.split("\n")
    t = t[1].strip()
    return sender, receiver, t


# 处理 附件
def cope_attachment(driver):
    element = driver.find_element(By.CLASS_NAME, "lh0")
    filename = element.text.split("\n")
    filename = filename[0].strip()
    actions = ActionChains(driver)
    actions.move_to_element(element).perform()
    sub_element = driver.find_element(By.XPATH, "//a[@target='downloadFrame']")
    sub_element.click()
    return filename


# 处理 正文 字段 该text属性受保护 无法直接获取
def cope_html(driver):
    element = driver.find_element(By.CLASS_NAME, "nu0")
    actions = ActionChains(driver)
    actions.move_to_element(element).click()
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL)
    actions.key_down(Keys.CONTROL).send_keys("c").key_up(Keys.CONTROL)
    actions.perform()
    time.sleep(1)
    html = pyperclip.paste()
    html = html.strip()
    return html


# 输出至 excel
def export_to_xlsx(infos, file):
    df = pd.DataFrame(infos)
    df.to_excel(file, index=False, engine="openpyxl")


def mail_main(cookie):
    # 打开浏览器 免登录 进入邮箱首页
    driver = open_browser_with_cookies(cookie)
    # 进入未读邮件列表
    click_to_mail_list(driver)
    index = 0
    mail_list = []
    # 处理每封未读邮件
    while click_to_mail(driver):
        save_mail(driver, index + 1, mail_list)
        index += 1
    if index:
        print(f"处理完毕. 共处理{index}封未读邮件.")
        # 输出至 excel
        export_file = "export/export_163mails.xlsx"
        export_to_xlsx(mail_list, export_file)
        print(f"输出至{export_file}")
    else:
        print(f"暂无未读邮件.")
    driver.quit()
    return mail_list

