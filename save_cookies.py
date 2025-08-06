from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pickle
import time

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)

driver.get("https://teams.microsoft.com")
print("Vui lòng đăng nhập vào Microsoft Teams (bao gồm MFA nếu có) và nhấn Enter khi hoàn tất...")
input()

cookies = driver.get_cookies()
with open("teams_cookies.pkl", "wb") as f:
    pickle.dump(cookies, f)
print("Cookie đã được lưu. Kiểm tra domain:", [cookie['domain'] for cookie in cookies])

driver.quit()
