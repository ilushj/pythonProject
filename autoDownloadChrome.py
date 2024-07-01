from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
import time

# 设置 ChromeDriver 和下载目录
driver_path = r'D:\chrome-win64'  # ChromeDriver 的实际路径
download_directory = r'd:\千服日报'  # 替换为你的下载目录

# 配置 Chrome 选项以指定下载目录
options = webdriver.ChromeOptions()
prefs = {'download.default_directory': download_directory}
options.add_experimental_option('prefs', prefs)
# options.add_argument('--headless')  # 添加headless参数
options.add_argument('--start-maximized')  # 最大化窗口
# 初始化 ChromeDriver
driver = webdriver.Chrome(options=options)


try:
    # 打开指定的网页并登录
    url = 'https://qidian.ekangonline.com/spiam/index.html'  # 替换为实际的登录页面 URL
    driver.get(url)

    # 等待用户名字段加载并输入用户名
    username_selector = 'account'  # 替换为实际的用户名字段 ID 或选择器
    password_selector = 'password'  # 替换为实际的密码字段 ID 或选择器
    login_button_selector = 'login'  # 替换为实际的登录按钮 ID 或选择器

    try:
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, username_selector))
        )
        username_field.send_keys('zhanjie')  # 替换为实际的用户名

        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, password_selector))
        )
        password_field.send_keys('Qidianbx123')  # 替换为实际的密码

        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, login_button_selector))
        )
        login_button.click()
    except TimeoutException:
        print("登录页面加载超时或元素未找到")
        driver.quit()
        exit(1)

    # 等待登录完成，检测登录成功
    try:
        # 使用 WebDriverWait 等待特定元素可见，以确认登录成功。param1为头像图片元素
        param1 = '//img[@class="layui-nav-img" and @src="images/icon/avatar_default.png"]'
        user_icon = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, param1))
        )
        print("登录成功")
    except TimeoutException:
        print("登录失败：未找到用户图标")
        driver.quit()
        exit(1)
    try:
        # param2 是日报表菜单元素
        param2 = '//a[@onclick="exec_menu_click(this)" and @data-url="pages/customer/importrecordday_upload.html"]'
        menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, param2))
        )
        menu.click()
        # 等待页面加载完成
        time.sleep(50)  # 等待页面加载，可以根据实际情况调整时间
        try:
            iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'content_iframe'))
            )
            driver.switch_to.frame(iframe)
            print("切换到 iframe 成功")

            # 打印 iframe 内部的页面源码以调试
            # iframe_source = driver.page_source
            # print(iframe_source)

            try:

                download_links = WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located((By.XPATH,
                                                         '//a[@class="layui-btn layui-btn-xs layui-btn-primary" '
                                                         'and @lay-event="export" and @title="人员清单"]'))
                )
                print(f"找到 {len(download_links)} 个下载链接")

                if not download_links:
                    print("未找到任何下载链接")
                else:
                    # 点击所有下载链接
                    for index, link in enumerate(download_links):
                        try:
                            # 滚动到元素
                            driver.execute_script("arguments[0].scrollIntoView();", link)
                            time.sleep(1)  # 等待滚动完成

                            link.click()
                            print(f"点击了第 {index + 1} 个下载链接")
                        except ElementClickInterceptedException:
                            print("点击被拦截，尝试使用 JavaScript 点击")
                            driver.execute_script("arguments[0].click();", link)
                            print("使用 JavaScript 点击了第一个下载链接")

                        time.sleep(2)  # 等待下载开始
            except NoSuchElementException:
                print("未找到下载链接元素")
            except TimeoutException:
                print("下载链接加载超时")

        except TimeoutException:
            print("iframe 加载超时或未找到")

    except NoSuchElementException:
        print("未找到下载链接元素")
    except TimeoutException:
        print("菜单加载超时")
    # 找到所有的下载链接并点击

except TimeoutException:
    print("页面加载超时")
except Exception as e:
    print(f"发生未知错误: {e}")
finally:
    # 关闭浏览器
    driver.quit()
