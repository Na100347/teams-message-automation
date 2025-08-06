from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from time import sleep
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pickle
# Cấu hình Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("teamstaskautomation-482c1370cde2.json", scope)
client = gspread.authorize(creds)

# URL của Sheet tổng quan và Sheet cá nhân
total_sheet_url = "https://docs.google.com/spreadsheets/d/1SWSVjinG8kefQB18YqIk3gDttaKOrhUwnxjh7OmEpXs/edit?gid=0#gid=0"  # Thay bằng URL thực tế
personal_sheet_url = "https://docs.google.com/spreadsheets/d/1E8g3inBy8IlzdUaM8nzBQL5SVOMp5NDX9Arh8KneWnc/edit?gid=0#gid=0"  # Thay bằng URL thực tế

def sync_tasks():
    your_name = "Tran Nu Ho Na"

    # Đọc dữ liệu từ Sheet tổng quan
    try:
        total_sheet = client.open_by_url(total_sheet_url).sheet1
        tasks = [row for row in total_sheet.get_all_values() if row]
        print("Tasks from sheet:", tasks)
    except Exception as e:
        print("Lỗi: Không đọc được Sheet tổng quan:", e)
        return

    # Lọc task của bạn và cập nhật Sheet cá nhân
    your_tasks = [task for task in tasks[1:] if task and task[1].strip() == your_name.strip()]
    if your_tasks:
        try:
            personal_sheet = client.open_by_url(personal_sheet_url).sheet1
            personal_sheet.clear()
            personal_sheet.append_rows(your_tasks)
            print(f"Đã cập nhật Sheet cá nhân cho {your_name}. Tasks: {your_tasks}")
        except gspread.exceptions.APIError as e:
            print(f"Lỗi quyền hoặc kết nối Sheet cá nhân: {e}")
            personal_sheet.append_rows(your_tasks)
        except Exception as e:
            print(f"Lỗi cập nhật Sheet cá nhân: {e}")
            return

    # Gửi tin nhắn nếu có task
    if your_tasks:
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        from selenium.webdriver.chrome.service import Service
        driver = webdriver.Chrome(service=Service('chromedriver'), options=chrome_options)
        wait = WebDriverWait(driver, 60)

        try:
            driver.get("https://teams.microsoft.com")
            sleep(5)

            try:
                with open("teams_cookies.pkl", "rb") as cookie_file:
                    cookies = pickle.load(cookie_file)
                    for cookie in cookies:
                        driver.add_cookie(cookie)
                driver.get("https://teams.microsoft.com")
            except FileNotFoundError:
                print("Cảnh báo: Không tìm thấy teams_cookies.pkl, cần tạo lại cookie.")
            except Exception as e:
                print("Cảnh báo: Không tải được cookie:", e)

            driver.get("https://teams.microsoft.com/_#/chat")
            sleep(8)

            try:
                your_chat = wait.until(EC.element_to_be_clickable((By.XPATH, f'//span[@title="{your_name}"]')))
                your_chat.click()
                print(f"Đã mở chat với {your_name}.")
            except Exception as e:
                print(f"Lỗi: Không tìm thấy {your_name} trong danh sách chat.")
                raise e

            sleep(5)

            message_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]')))
            ActionChains(driver).move_to_element(message_box).click().perform()
            sleep(1)

            for task in your_tasks:
                message = f"Task: {task[0]}\nAssigned to: {task[1]}\nDeadline: {task[2]}"
                for line in message.split('\n'):
                    message_box.send_keys(line)
                    message_box.send_keys(Keys.SHIFT, Keys.ENTER)
                    sleep(0.2)
                message_box.send_keys(Keys.RETURN)
                print(f"Đã gửi tin nhắn cho task: {task[0]}")
                sleep(3)

        except Exception as e:
            print("Lỗi khi gửi tin nhắn:", e)
        finally:
            driver.quit()
    else:
        print(f"Không có task nào cho {your_name}.")

if __name__ == "__main__":
    sync_tasks()
