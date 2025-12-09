from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService  # 重命名以避免冲突
from selenium.webdriver.edge.options import Options
import time
import os
import sys

# 【新增】从 webdriver_manager 导入 EdgeChromiumDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# --- 用户配置区 ---
# incite账号 (请切换为自己的)
ACCOUNT = r"1832411836@qq.com"
# 填写对应密码
PSWD = r"jsb123456700*"
# IcDoc_Org起止年份
INI_YEAR = 2015
END_YEAR = 2025
# 下载文件的保存路径 (重要: 请确保该文件夹存在)
DOWNLOAD_DIR = os.path.join(os.getcwd(), "InCites_Downloads")

# 【已移除】手动驱动路径配置，现由 webdriver-manager 自动管理
# EDGE_DRIVER_PATH = r"msedgedriver.exe"

# --- 全局常量 ---
TIMEOUT = 30  # 操作的超时，单位秒
DOWNLOAD_TIMEOUT = 60  # 下载操作的超时，单位秒
LOCATION = "CHINA MAINLAND"
LOGIN_URL = "https://access.clarivate.com/login?app=incites"
ORG_URL = "https://incites.clarivate.com/zh/#/analysis/0/organization"
AGREE_URL = "https://access.clarivate.com/login?app=incites"

# 确保下载目录存在
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


def get_browser():
    """【已修改】使用 EdgeChromiumDriverManager 自动配置驱动"""
    edge_options = Options()

    # 设置下载目录
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    edge_options.add_experimental_option("prefs", prefs)

    # 禁用 WebDriver 标识 (可选，用于反爬)
    edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    edge_options.add_experimental_option('useAutomationExtension', False)

    # 【核心修改】自动下载或获取 Edge WebDriver 的路径
    service = EdgeService(EdgeChromiumDriverManager().install())

    print(f"初始化 Edge 浏览器，下载路径设置为: {DOWNLOAD_DIR}")
    return webdriver.Edge(service=service, options=edge_options)


def login(b):
    """登录 InCites 平台"""
    try:
        print("开始登录...")
        b.get(LOGIN_URL)

        # 显式等待登录输入框出现
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="email"]'))
        )

        # 查找输入框并输入账号和密码
        email_input = b.find_element(By.CSS_SELECTOR, 'input[formcontrolname="email"]')
        password_input = b.find_element(By.CSS_SELECTOR, 'input[formcontrolname="password"]')
        sign_in_button = b.find_element(By.ID, "signIn-btn")

        email_input.send_keys(ACCOUNT)
        password_input.send_keys(PSWD)
        sign_in_button.click()

        # 检查是否成功跳转到 InCites 主页
        WebDriverWait(b, TIMEOUT + 15).until(
            EC.url_contains("/analysis/0")
        )
        print("登录成功并跳转到分析页面。")

        # 尝试点击 'accept all cookies' 按钮
        try:
            cookie_button = WebDriverWait(b, 5).until(
                EC.element_to_be_clickable((By.ID, r'onetrust-accept-btn-handler'))
            )
            cookie_button.click()
            print("已接受 Cookie 协议。")
        except:
            print("未发现 Cookie 弹窗或已处理。")

        return True
    except Exception as ex:
        print(f"登录失败: {ex}")
        return False


def apply_initial_filters(b):
    """
    @description 初始化IcDoc_Org的通用配置: ESI分类体系、起止年份、CHINA MAINLAND
    @param b: webdriver
    @return: bool
    """
    try:
        print("开始设置初始化筛选条件...")

        # 1. 设置分类体系为 ESI (研究方向 -> ESI)

        # 点击 '研究方向'
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH,
                                        "/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[3]/ul/li[15]/button/span"))
        ).click()

        # 点击 '学科分类体系' 下拉菜单
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//cl-select-filter//nav/ul/li/a"))
        ).click()

        # 点击 'esi' 选项
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//cl-select-filter//nav/ul/li/ul/li[3]/label"))
        ).click()

        # 点击确认
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//ic-sidebar-filter//div[4]/div[1]/button[2]"))
        ).click()
        print("已设置学科分类体系为 ESI。")

        # 2. 设置年份范围

        # 显式等待年份输入框出现
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'cl-range-input input[placeholder="起始年份"]'))
        )

        # 使用 JS 脚本设置年份 (更稳定)
        b.execute_script(
            f"document.querySelector('cl-range-input input[placeholder=\"起始年份\"]').value='{INI_YEAR}';"
            f"document.querySelector('cl-range-input input[placeholder=\"起始年份\"]').dispatchEvent(new Event('input', {{ bubbles : true }}));"
        )
        b.execute_script(
            f"document.querySelector('cl-range-input input[placeholder=\"截止年份\"]').value='{END_YEAR}';"
            f"document.querySelector('cl-range-input input[placeholder=\"截止年份\"]').dispatchEvent(new Event('input', {{ bubbles : true }}));"
        )
        print(f"已设置年份范围为 {INI_YEAR}-{END_YEAR}。")

        # 3. 设置 LOCATION 为 CHINA MAINLAND

        # 点击 '地点' 筛选按钮 (假设是 li[5]，与原代码一致)
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH,
                                        "/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[3]/ul/li[5]/button/span"))
        ).click()

        # 等待地点搜索框出现
        location_input_css = 'cl-multi-select input[placeholder="输入要添加的地点"]'
        location_input = WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, location_input_css))
        )

        location_input.send_keys(LOCATION)

        # 等待搜索结果出现，并点击第一个结果
        first_result_xpath = "//ul[@class='cl-multi-select__options__list']/li[1]/cl-text-highlight/span"
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, first_result_xpath))
        ).click()

        # 点击确认
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//ic-sidebar-filter//div[4]/div[1]/button[2]"))
        ).click()
        print(f"已设置地点筛选为 {LOCATION}。")

        # 等待数据表格刷新
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cl-sortable-table__row__value"))
        )
        print("初始化筛选条件设置完毕。")
        return True
    except Exception as ex:
        print(f"初始化配置失败: {ex}")
        return False


def set_indicators(b):
    """设置所有指标（排除第六组）"""
    try:
        print("开始设置下载指标...")

        # 1. 点击 '添加指标' 按钮
        add_indicator_button_xpath = "//ic-add-indicator-menu//button"
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, add_indicator_button_xpath))
        ).click()

        # 等待指标选择弹窗出现
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cl-popover-modal__body"))
        )

        # 2. 遍历并选择指标
        fs = range(1, 8)
        excl = [6]  # 排除第六组指标

        for i in fs:
            fieldset_xpath = f"//ic-add-indicator-menu//fieldset[{i}]"
            fieldset_element = b.find_element(By.XPATH, fieldset_xpath)
            check_boxes = fieldset_element.find_elements(By.TAG_NAME, "input")

            for box in check_boxes:
                if i not in excl:
                    # 选中所有非排除组的指标
                    if not box.is_selected():
                        box.send_keys(Keys.SPACE)  # 使用 Keys.SPACE 触发点击更可靠
                else:
                    # 取消选中排除组的指标
                    if box.is_selected():
                        box.send_keys(Keys.SPACE)

        # 3. 点击 '确认' 按钮
        confirm_button_xpath = "//ic-add-indicator-menu//cl-popover-modal-content//button[text()='确认']"
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, confirm_button_xpath))
        ).click()

        # 等待数据表格刷新
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cl-sortable-table__row__value"))
        )
        print("指标设置完毕。")
        return True
    except Exception as ex:
        print(f"设置指标失败: {ex}")
        return False


def init(b):
    """初始化浏览器，登录，设置通用筛选条件和指标"""
    if not login(b):
        return False

    # 1. 跳转到 '机构' 分析页面
    b.get(ORG_URL)

    # 2. 等待页面加载完成 (等待数据表格出现)
    try:
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cl-sortable-table__row__value"))
        )
    except:
        print("警告: 机构数据表格加载超时，尝试继续初始化。")

    # 3. 设置通用筛选条件 (年份、ESI、CHINA MAINLAND)
    if not apply_initial_filters(b):
        return False

    # 4. 设置下载指标
    if not set_indicators(b):
        return False

    return True


def apply_journal_filter(b, name):
    """
    @description 应用期刊名称筛选条件
    @param b: webdriver
    @param name: 期刊名称
    """
    print(f"-> 正在应用期刊筛选: {name}")
    try:
        # 1. 点击 '来源出版物' 筛选按钮
        journal_filter_button_xpath = "//ic-sidebar-filters//ul/li[4]/button"
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, journal_filter_button_xpath))
        ).click()

        # 2. 等待搜索框出现
        search_input_css = 'cl-multi-select input[placeholder="输入要添加的来源出版物"]'
        search_input = WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, search_input_css))
        )

        # 3. 输入期刊名称
        search_input.clear()
        search_input.send_keys(name)

        # 4. 等待搜索结果出现，并点击第一个结果
        first_result_xpath = "//ul[@class='cl-multi-select__options__list']/li[1]/cl-text-highlight/span"
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, first_result_xpath))
        ).click()

        # 5. 点击确认
        WebDriverWait(b, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//ic-sidebar-filter//div[4]/div[1]/button[2]"))
        ).click()

        # 等待数据表格刷新
        WebDriverWait(b, TIMEOUT).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cl-sortable-table__row__value"))
        )
        print(f"-> 期刊筛选 {name} 应用成功。")
        return True
    except Exception as ex:
        print(f"-> 期刊筛选 {name} 应用失败: {ex}")
        # 如果失败，尝试清理已打开的筛选侧栏
        try:
            # 点击取消按钮关闭侧栏 (假设 Cancel 按钮是第一个)
            cancel_button_xpath = "//ic-sidebar-filter//div[4]/div[1]/button[1]"
            WebDriverWait(b, 5).until(EC.element_to_be_clickable((By.XPATH, cancel_button_xpath))).click()
        except:
            pass
        return False


def download(b, name):
    """
    @description 下载当前筛选条件下（某个期刊）的所有机构数据
    @param b: webdriver
    @param name: 期刊名称
    """
    print(f"-> 正在下载 {name} 的数据...")

    # 1. 应用期刊筛选
    if not apply_journal_filter(b, name):
        raise Exception(f"Failed to apply journal filter for {name}")

    # 2. 点击表格右上角的下载按钮
    download_btn_xpath = "//ic-analysis-table//div[@class='cl-tab-pane__toolbar']/ic-analysis-table-download/cl-popover-modal/div/cl-popover-modal-button/button"
    WebDriverWait(b, TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, download_btn_xpath))
    ).click()

    # 3. 等待下载弹窗出现
    WebDriverWait(b, TIMEOUT).until(
        EC.presence_of_element_located((By.CLASS_NAME, "cl-popover-modal-content"))
    )

    # 4. 设置文件名 (使用期刊名称作为前缀)
    filename_input = b.find_element(By.ID, 'file-name')
    filename_input.clear()

    # 避免文件名包含非法字符
    safe_name = name.replace('/', '_').replace(':', '_').replace('\\', '_')
    filename_input.send_keys(f"ORG_Data_{safe_name}")

    # 5. 设置下载格式为 CSV (使用 JS 脚本操作 Select 元素)
    b.execute_script(
        """
        var selectElement = document.querySelector('select[formcontrolname="downloadOption"]'); 
        selectElement.value = 'csv'; 
        selectElement.dispatchEvent(new Event('change', { bubbles: true }));
        """
    )

    # 6. 点击下载按钮 (在弹窗内部)
    download_confirm_xpath = "//cl-popover-modal-content//button[text()='下载']"
    WebDriverWait(b, TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, download_confirm_xpath))
    ).click()

    # 7. 等待下载完成 (在 Edge WebDriver 中，文件下载后，通常没有明确的完成信号，
    # 只能依赖于文件系统监控或延时，这里使用一个较长的等待)
    print(f"-> 等待下载 {safe_name}.csv 到 {DOWNLOAD_DIR}...")
    time.sleep(DOWNLOAD_TIMEOUT)

    # 8. 清理筛选条件

    # 点击右上角已应用的期刊筛选的关闭按钮 (假设它是第二个应用筛选条件)
    close_filter_btn_xpath = "/html/body/app-root/main/ic-analysis/div/div[1]/cl-applied-filters/div/ul/li[2]/span/button"
    WebDriverWait(b, TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, close_filter_btn_xpath))
    ).click()

    # 等待数据表格刷新
    WebDriverWait(b, TIMEOUT).until(
        EC.presence_of_element_located((By.CLASS_NAME, "cl-sortable-table__row__value"))
    )
    print(f"-> {name} 下载完成，已清除筛选条件。")


def main():
    print("--- InCites 机构数据自动化下载脚本 ---")
    print(f"下载目录: {DOWNLOAD_DIR}")

    # 提示用户输入文件路径
    while True:
        try:
            p = input("请输入 CSV 文件路径并回车: ")
            f = open(p, "r", encoding="latin-1")
            # 读取并清理期刊名称列表
            name_list = [i.strip().replace('"', '') for i in f.readlines() if i.strip()]
            f.close()
            if not name_list:
                raise ValueError("CSV 文件中没有有效的期刊名称。")
            break
        except FileNotFoundError:
            print(f"错误: 找不到文件 {p}，请检查路径。")
        except ValueError as e:
            print(f"错误: {e}")
        except Exception as e:
            print(f"读取文件时发生意外错误: {e}")
            sys.exit(1)

    # 初始化浏览器并登录
    browser = None
    try:
        browser = get_browser()
        if not init(browser):
            print("初始化失败，程序退出。")
            return
    except Exception as e:
        print(f"浏览器或初始化失败: {e}")
        if browser:
            browser.quit()
        return

    # 开始下载循环
    fa = 0
    leng = len(name_list)

    for c, name in enumerate(name_list, 1):
        try:
            print(f"\n--- ({c}/{leng}) 任务: {name} ---")
            download(browser, name)
        except Exception as exc:
            fa += 1
            print(f"!!! ({c}/{leng}) 下载失败: {name}. 正在尝试重新初始化浏览器...")
            try:
                # 重新初始化并继续
                if browser:
                    browser.quit()
                browser = get_browser()
                if not init(browser):
                    raise Exception("重新初始化失败。")
            except Exception as reinit_exc:
                print(f"致命错误: 无法重新初始化浏览器。{reinit_exc}")
                break

    print("\n--- 任务总结 ---")
    print(f"总任务数: {leng}")
    print(f"成功数: {leng - fa}")
    print(f"失败数: {fa}")
    print("脚本完成。")
    time.sleep(5)

    if browser:
        browser.quit()


if __name__ == "__main__":
    main()