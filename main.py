import requests
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from time import sleep
# from flask import Flask, request
import tempfile
import time
import pickle
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# app = Flask(__name__)

# Cấu hình Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("teamstaskautomation-482c1370cde2.json", scope)
client = gspread.authorize(creds)

# URL của Sheet tổng quan và Sheet cá nhân
total_sheet_url = "https://docs.google.com/spreadsheets/d/1SWSVjinG8kefQB18YqIk3gDttaKOrhUwnxjh7OmEpXs/edit?gid=0#gid=0"  # Thay bằng URL của "Task_Tong_Quan"
personal_sheet_url = "https://docs.google.com/spreadsheets/d/1E8g3inBy8IlzdUaM8nzBQL5SVOMp5NDX9Arh8KneWnc/edit?gid=0#gid=0"  # Thay bằng URL của "Task_Ca_Nhan_Tran_Nu_Ho_Na"


@app.route('/webhook/send-teams-task', methods=['POST'])
def send_message():
    data = request.get_json()
    task = data.get("task")
    assignee = data.get("assignee")
    deadline = data.get("deadline")
    email = data.get("email")
    your_name ="Tran Nu Ho Na"
    
    # Đọc file task tổng quan
    try:
        total_sheet = client.open_by_url(total_sheet_url).sheet1
        tasks = [row for row in total_sheet.get_all_values() if row]  # Lấy tất cả hàng
        print("Tasks from sheet:", tasks)  # Debug
    except Exception as e:
        print("Lỗi: Không đọc được Sheet tổng quan:", e)
        return {"status": "Lỗi", "message": "Không đọc được Sheet tổng quan"}, 500

    # Lọc task của bạn và cập nhật Sheet cá nhân
    personal_tasks = [task for task in tasks[1:] if task and task[1].strip() == your_name.strip()]  # Bỏ header, so sánh cột 2
    if personal_tasks:
        personal_sheet = client.open_by_url(personal_sheet_url).sheet1
        personal_sheet.clear()  # Xóa dữ liệu cũ
        personal_sheet.append_rows(personal_tasks)  # Thêm task mới
        print(f"Đã cập nhật Sheet cá nhân cho {your_name}. Tasks: {personal_tasks}")
    #setup chrome so don't have to log in everytime
    if personal_tasks:
        # Setup Chrome
        chrome_options = Options()
        chrome_options.add_argument("--user-data-dir=C:\\Temp\\selenium-profile-1")
        chrome_options.add_argument("--profile-directory=Profile 1")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        from selenium.webdriver.chrome.service import Service
        driver = webdriver.Chrome(service=Service('D:\\test\\teams-message-automation\\chromedriver-win64\\chromedriver.exe'), options=chrome_options)
        wait = WebDriverWait(driver, 60)
        try:
            driver.get("https://teams.microsoft.com")
            sleep(5)
            
            #load cookies from previous session before opening Teams
            try:
                # with open("/Users/vietnguyen/Desktop/teams_cookies.pkl","rb") as cookie_file:
                with open("D:\\test\\teams-message-automation\\teams_cookies.pkl", "rb") as cookie_file:
                    cookies = pickle.load(cookie_file)
                    for cookie in cookies:
                        driver.add_cookie(cookie)
                driver.get("https://teams.microsoft.com")
            except Exception as e:
                print("Warning: Could not load cookies:",e)

        # Navigate to Teams main chat page
            driver.get("https://teams.microsoft.com/_#/chat")
            sleep(8)
            
            # #Locate Ray's existing chat
            # chat_xpaths = [
            #     '//span[contains(text(), "Tran Nu Ho Na")]',
            #     '//div[contains(@data-tid, "chat-list")]//span[contains(text(), "Tran Nu Ho Na")]',
            #     '//button[contains(@aria-label, "Tran Nu Ho Na")]',
            # ]

            # Wait for chat list sidebar to load and search for "Ray"
            try:
                your_chat = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@title="Tran Nu Ho Na"]')))
                your_chat.click()
                print(" Opened existing chat with {your_name}.")
            except Exception as e:
                print(" Could not find {your_name} in chat list. Make sure {your_name} is pinned or recently chatted.") 
                raise e

            sleep(5)
            
            # Wait for the message box and click
            message_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]')))
            ActionChains(driver).move_to_element(message_box).click().perform()
            sleep(1)
        
            # Format and send the message
            # Format and send the message for each personal task
            for task in personal_tasks:
                formatted_message = f"Task: {task[0]}\nAssigned to: {task[1]}\nDeadline: {task[2]}"
                for line in formatted_message.split('\n'):
                    message_box.send_keys(line)
                    message_box.send_keys(Keys.SHIFT, Keys.ENTER)
                    sleep(0.2)
                message_box.send_keys(Keys.RETURN)
                print(f"Message sent for task: {task[0]}")
                sleep(3)
        except Exception as e:
            print("Error sending message:",e)
            raise e
                
        finally:
            driver.quit()
    else:
        print("Không có task nào cho bạn.")   
    return {"status": "Message sent" if personal_tasks else "No tasks"}, 200

if __name__=='__main__':
    app.run(port=5001)
#   app.run(host="0.0.0.0", port=5001)
