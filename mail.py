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
def open_browser_with_cookies():
    driver_path = "C:/Users/86158/chromedriver-win64/chromedriver-win64/chromedriver.exe"
    service = Service(executable_path=driver_path)

    options = webdriver.ChromeOptions()
    base_dir = os.getcwd()
    relative_download_dir = "export"
    download_dir = os.path.join(base_dir, relative_download_dir)
    os.makedirs(download_dir, exist_ok=True)
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://mail.163.com/")
    cookie_str = "NTES_P_UTID=cVz61pNK7RHIYhqB4UDfF6LTYqeT9Sz8|1747963263; NTES_PASSPORT=EbaCrCYGTKGzauqx4WVyvAmuho4INUvRZyXWk03uQ5R8a7s5atXR4yb_dBBB6bCFEfBYI.Ento2KbMxnWQWC_hIfFZ0wHDzOnyBx8SPKawSbiBdJgiRhpcQI26eDKm.RYN0UtomIlXPkQinKusNYfGN9wtSATHT20EWWiE_9aaVKTMYauyxHGS6Zqk4JzhOIn; P_INFO=agenttest0514@163.com|1747963263|1|mail163|00&99|zhj&1747212624&unireg#zhj&330100#10#0#0|&0|unireg|agenttest0514@163.com; stats_session_id=0d681780-adaa-4ca0-a379-2ba4705df696; mail_idc=''; MAIL_SDID=1243131484874293248; MAIL_MISC=agenttest0514; MAIL_PINFO='agenttest0514@163.com|1747963263|1|mail163|00&99|zhj&1747212624&unireg#zhj&330100#10#0#0|&0|unireg|agenttest0514@163.com'; MAIL_LOCATION=gz; MAIL_PASSPORT_INFO=agenttest0514@163.com|1747963263|1; secu_info=1; locale=; face=js6; mail_style=js6; mail_uid=agenttest0514@163.com; mail_host=mail.163.com; p_user=agenttest0514; p_user.sig=GuZ87icIrKLa9bhcDgBwcFgZkrI; mailteam_info={'isOwner':0}; mailteam_info.sig=bOvgSFNB4qCEtQtSQ067MBvZL6c; _ntes_nuid=ea92d68c187ee1a1da5e92d33918ccd5; starttime=; S_INFO=1748158108|1|0&60##|agenttest0514; NTES_SESS=ph9FUOqWgKnyXQUOx2N0T6w6MAR6J_lc8.BK94xCrXuS_q13_H8IY.6uMSSSa6hLiWMkQ9NXIljTtb_ybVrzcw285PW7bJbZxlrsQ55OJ3srIN1KxVYz34KLbNAMm.XfNW5DprgTIMUe7342BHrxRKKxM7w4KtZmcfUgY615IsntIV.RKYgmdnH6r_5F9W0nJv2LLtqAdiZvVq3s5AIDjb2k0; MAIL_SESS=ph9FUOqWgKnyXQUOx2N0T6w6MAR6J_lc8.BK94xCrXuS_q13_H8IY.6uMSSSa6hLiWMkQ9NXIljTtb_ybVrzcw285PW7bJbZxlrsQ55OJ3srIN1KxVYz34KLbNAMm.XfNW5DprgTIMUe7342BHrxRKKxM7w4KtZmcfUgY615IsntIV.RKYgmdnH6r_5F9W0nJv2LLtqAdiZvVq3s5AIDjb2k0; MAIL_SINFO=1748158108|1|0&60##|agenttest0514; Coremail=fb92d5d011e82%TLztlHJhXKsCDCmLUuhhtlupxCksEbAC%wmsvr-40-115.mail.163.com; MAIL_ENTRY_INFO=1|24|mail163|mail163_letter|122.224.251.228|d5c131fb0e5a6b9da8c977f41a50b51c_v1||1243131484874293248|1748178562|0|0; MAIL_ENTRY_CS=c0a7b9ef7a08d06c97e2ce4eb1570345; mixmailTokens=1D1f_zmJDZBKRCndGDQ6vN8za3tyi6mlJENf6Mbpv5LH8TLU8RjsqRPk0j1wxm95Fa2k2XSA-ZrWfAIqCC-PaW_i29_ZSRS8R_U0nKDjLaLwOpz9FfnZlMaVS7ofVURXHyvvFW3rcnut2Zy9fz8cD3UoA; cm_last_info='dT1hZ2VudHRlc3QwNTE0JTQwMTYzLmNvbSZkPWh0dHBzJTNBJTJGJTJGbWFpbC4xNjMuY29tJTJGanM2JTJGbWFpbi5qc3AlM0ZzaWQlM0RUTHp0bEhKaFhLc0NEQ21MVXVoaHRsdXB4Q2tzRWJBQyZzPVRMenRsSEpoWEtzQ0RDbUxVdWhodGx1cHhDa3NFYkFDJmg9aHR0cHMlM0ElMkYlMkZtYWlsLjE2My5jb20lMkZqczYlMkZtYWluLmpzcCUzRnNpZCUzRFRMenRsSEpoWEtzQ0RDbUxVdWhodGx1cHhDa3NFYkFDJnc9aHR0cHMlM0ElMkYlMkZtYWlsLjE2My5jb20mbD0tMSZ0PS0xJmFzPXRydWU='; MAIL_PASSPORT=tYqKGuiBZtesVYvlyjimmiVmhl7fY1jNnznaoqsxEYuTybvYyrnuZzS.wjjjGS8PtmjkFUtcrMK3SOecaEa8.7FmPNqp4VLDczjeTRl3ypRSHjwQ9Hu7iWEFKGhV31UukdqCrM1F0nloEHc3xvdkmId2prRgB4BKqtaaHt.2yy53BOkyxze4IRGN6oZQL7DFc; mail_entry_sess=760e62ddb2c67e55dce592690010d5ac275e34f4e13fdd8cc592f5ded63bf7611aa654992723ce1d4711e0e8ada11faeecba3d7c0952eabd57802ccdbe2846a7f290caa218f174bbd7b50928e9250e6cac17fadce2b3e22fb35c2fdfb1a7d05f162495e73b9e87cf51b95199ab8492120be957469d10e87d224098cb8278c545639ef4b30b19569b0d73987f3cc30b1b37a1b31c5cb47a7e179677d9cfe2cab2e5a020e7eb04b91cf4bc0d6e7317f21b2ee289d132a0a4eb9071c3705a11be53; JSESSIONID=E817945C9AE4FC7200E6915E403861C0; Coremail.sid=TLztlHJhXKsCDCmLUuhhtlupxCksEbAC; NTES_PC_EMAIL_IP=%E6%9D%AD%E5%B7%9E%7C%E6%B5%99%E6%B1%9F"
    for cookie in cookie_str.split(';'):
        cookie = cookie.strip()
        if '=' in cookie:
            name, value = cookie.split('=', 1)
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
'''
    邮件标题
    发件人
    收件人
    时间
    附件
    正文
'''
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


if __name__ == "__main__":
    # 打开浏览器 免登录 进入邮箱首页
    driver = open_browser_with_cookies()
    # 进入未读邮件列表
    click_to_mail_list(driver)
    index = 0
    mail_list = []
    # 处理每封未读邮件
    while click_to_mail(driver):
        save_mail(driver, index+1, mail_list)
        index += 1
    if index:
        print(f"处理完毕. 共处理{index}封未读邮件.")
        # 输出至 excel
        export_file = 'export/export_163mails.xlsx'
        export_to_xlsx(mail_list, export_file)
        print(f"输出至{export_file}")
    else:
        print(f"暂无未读邮件.")
    input("enter to be terminated")
    driver.quit()
