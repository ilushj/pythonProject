from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
import time


# 指定下载目录（请替换为你的实际路径）
download_dir = r"D:\数据总表"  # 替换为你的目录

# 确保下载目录存在，如果不存在则创建
if not os.path.exists(download_dir):
    os.makedirs(download_dir)


# 配置 ChromeOptions 以设置下载相关参数
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,  # 设置默认下载目录为指定路径
    "download.prompt_for_download": False,       # 禁用下载提示，直接开始下载
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})


# 初始化浏览器驱动
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)


# 打开登录页面
driver.get("https://qidian.ekangonline.com/spiam/login.html")
print("初始URL:", driver.current_url)


# 定义页面元素的选择器（需根据实际页面元素进行替换）
username_selector = "account"
password_selector = "password"
login_button_selector = "login"


try:
    # 等待用户名输入框出现并输入用户名
    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, username_selector))
    )
    username = input("请输入用户名: ")
    username_field.send_keys(username)
    print("用户名输入成功")

    # 等待密码输入框出现并输入密码
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, password_selector))
    )
    password = input("请输入密码: ")
    password_field.send_keys(password)
    print("密码输入成功")

    # 等待登录按钮可点击并点击
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, login_button_selector))
    )
    login_button.click()
    print("登录按钮点击成功")

    # 检查登录是否成功（假设登录成功后有特定的欢迎元素）
    try:
        param1 = '//img[@class="layui-nav-img" and @src="images/icon/avatar_default.png"]'
        user_icon = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, param1))
        )
        print("登录成功，当前URL:", driver.current_url)

        # 点击特定的菜单以进入数据页面
        param2 = '//a[@onclick="exec_menu_click(this)" and @data-url="pages/ekang/dataTotal.html?status=manager&type=2"]'
        menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, param2))
        )
        menu.click()
        time.sleep(2)

        # 切换到 iframe 中操作
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "content_iframe"))
        )
        driver.switch_to.frame(iframe)
        print("已切换到 iframe: content_iframe")

        # 定位“查询月”输入框并设置值
        try:
            query_month_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "search_month"))
            )
            print("找到 search_month 元素")
            date_value = "2023-01 - 2023-12"
            driver.execute_script("arguments[0].value = arguments[1];", query_month_field, date_value)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", query_month_field)
        except TimeoutException:
            print("在 iframe 中未找到 search_month，检查 iframe 是否正确")
            print("当前 iframe 源码:", driver.page_source[:1000])

        # 点击“查询”按钮进行数据查询
        try:
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@lay-event='search']"))
            )
            search_button.click()
            print("查询按钮点击成功")

            # 等待加载图标出现，判断查询是否开始
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "layui-icon-loading"))
                )
                print("检测到加载图标 (layui-icon-loading)")
            except TimeoutException:
                print("未检测到加载图标 (layui-icon-loading)，可能查询立即完成或选择器错误")

            # 等待加载图标消失，判断查询是否完成
            try:
                WebDriverWait(driver, 30).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, "layui-icon-loading"))
                )
                print("加载图标 (layui-icon-loading) 已消失，查询完成")
            except TimeoutException:
                print("加载图标 (layui-icon-loading) 未消失，查询可能超时或选择器错误")
        except TimeoutException:
            print("在 iframe 中未找到查询按钮")

        time.sleep(10)

        # 点击“导出”按钮导出数据
        try:
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@lay-event='export']"))
            )
            search_button.click()
            print("导出按钮点击成功")
            # 简单等待 10 秒，等待文件下载完成
            time.sleep(10)
            print(f"文件已下载到: {download_dir}")
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

time.sleep(20)
# 关闭浏览器
driver.quit()