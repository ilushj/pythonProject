from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

# 初始化浏览器
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.get("https://qidian.ekangonline.com/spiam/login.html")
print("初始URL:", driver.current_url)

# 定义选择器（替换为实际值）
username_selector = "account"
password_selector = "password"
login_button_selector = "login"

try:
    # 等待并输入用户名
    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, username_selector))
    )
    username = input("请输入用户名: ")
    username_field.send_keys(username)
    print("用户名输入成功")

    # 等待并输入密码
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, password_selector))
    )
    password = input("请输入密码: ")
    password_field.send_keys(password)
    print("密码输入成功")

    # 等待并点击登录按钮
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, login_button_selector))
    )
    login_button.click()
    print("登录按钮点击成功")

    # 检查登录结果（假设登录后有欢迎元素）
    try:
        param1 = '//img[@class="layui-nav-img" and @src="images/icon/avatar_default.png"]'
        user_icon = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, param1))
        )
        print("登录成功，当前URL:", driver.current_url)
        param2 = '//a[@onclick="exec_menu_click(this)" and @data-url="pages/ekang/dataTotal.html?status=manager&type=2"]'
        menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, param2))
        )
        menu.click()
        time.sleep(2)

        # 切换到 iframe
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "content_iframe"))
        )
        driver.switch_to.frame(iframe)
        print("已切换到 iframe: content_iframe")

        # 定位“查询月”并设置值
        try:
            query_month_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "search_month"))
            )
            print("找到 search_month 元素")
            date_value = "2023-01-2023-12"
            driver.execute_script("arguments[0].value = arguments[1];", query_month_field, date_value)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", query_month_field)
        except TimeoutException:
            print("在 iframe 中未找到 search_month，检查 iframe 是否正确")
            print("当前 iframe 源码:", driver.page_source[:1000])

        # 点击“查询”按钮
        try:
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@lay-event='search']"))
            )
            search_button.click()
            print("查询按钮点击成功")
        except TimeoutException:
            print("在 iframe 中未找到查询按钮")
        time.sleep(12)

        try:
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@lay-event='export']"))
            )
            search_button.click()
            print("导出按钮点击成功")

            # 等待提示框并确保可见
            error_popup = WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "layui-layer-content"))
            )
            error_message = error_popup.text.strip()  # 去除首尾空白
            if error_message:
                print(f"查询出错，提示信息: {error_message}")
            else:
                print("查询出错，但提示信息为空")
        except TimeoutException:
            print("在 iframe 中未找到导出按钮")

    except TimeoutException:
        print("登录失败，检查是否有错误提示")
        try:
            error = driver.find_element(By.CLASS_NAME, "error")  # 替换为实际错误提示选择器
            print("错误信息:", error.text)
        except NoSuchElementException:
            print("未找到错误提示，可能页面未变化")

except TimeoutException as e:
    print(f"超时错误: {e}")
    print("当前URL:", driver.current_url)
except NoSuchElementException as e:
    print(f"元素未找到: {e}")

