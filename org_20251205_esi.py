# 只需要安装selenium这个库。
# 配置好所有参数后，运行脚本，根据提示输入csv文件的路径回车即可。
"""
csv文件的格式如下：
NATURE COMMUNICATIONS
SCIENTIFIC REPORTS
SCIENCE OF THE TOTAL ENVIRONMENT
......
"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import time

# 用Edge浏览器的webdriver实现，请自行百度如何为selenium配置Edge的webdriver。理论上可以用Chrome，我没有测试过。
BROWSER = webdriver.Edge
# incite账号 切换为自己的
ACCOUNT = r"1832411836@qq.com"

"""
填写对应密码
"""

PSWD = r"jsb123456700*"
# IcDoc_Org起止年份
INI_YEAR = 2015
END_YEAR = 2025
# 除下载以外的操作的超时，单位秒
TIMEOUT = 30
# 下载操作的超时，单位秒
DOWNLOAD_TIMEOUT = 60
LOCATION = "CHINA MAINLAND"
LOGIN_URL = "https://access.clarivate.com/login?app=incites"
JOUR_URL = "https://incites.clarivate.com/zh/#/analysis/0/journal"
ORG_URL = "https://incites.clarivate.com/zh/#/analysis/0/organization"
AGREE_URL = "https://access.clarivate.com/login?app=incites"
def login(b):
    try:
        # Step 1: 打开登录页面
        b.get(LOGIN_URL)
        WebDriverWait(b, TIMEOUT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Step 2: 检查是否重定向到隐私条款页面
        # if b.current_url.startswith(AGREE_URL):
        #     print("Detected redirection to privacy agreement page. Agreeing to terms...")
        #
        #     # 等待条款页面加载完成
        #     WebDriverWait(b, TIMEOUT).until(
        #         EC.presence_of_element_located(
        #             (By.XPATH, "/html/body/microui-app/div/div/microui-base/div/div[1]/div/base-login/div/cross-border-data-acknowledgement/div[3]/button[2]")
        #         )
        #     )
        #
        #     # 勾选第一个复选框
        #     checkbox1 = b.find_element(By.XPATH, "/html/body/microui-app/div/div/microui-base/div/div[1]/div/base-login/div/cross-border-data-acknowledgement/section/div[1]/mat-checkbox/label/span[1]")
        #     checkbox1.click()
        #     print("Checked first agreement box.")
        #
        #     # 勾选第二个复选框
        #     checkbox2 = b.find_element(By.XPATH, "/html/body/microui-app/div/div/microui-base/div/div[1]/div/base-login/div/cross-border-data-acknowledgement/section/div[2]/mat-checkbox/label/span[1]")
        #     checkbox2.click()
        #     print("Checked second agreement box.")
        #
        #     # 点击同意按钮
        #     agree_button = b.find_element(By.XPATH, "/html/body/microui-app/div/div/microui-base/div/div[1]/div/base-login/div/cross-border-data-acknowledgement/div[3]/button[2]")
        #     agree_button.click()
        #     print("Clicked agree button.")
        #
        #     # 等待页面重定向回登录页面
        #     print("Redirected to login page.")

        # Step 3: 填写账号和密码
        b.get(LOGIN_URL)
        e = WebDriverWait(b, TIMEOUT).until(EC.presence_of_all_elements_located)
        cmds = [
            # Enter account and password
            r'document.querySelector("#mat-input-1").value = "{}";'.format(ACCOUNT),
            r'document.querySelector("#mat-input-1").dispatchEvent(new Event("input", { bubbles : true }));',
            r'document.querySelector("#mat-input-0").value = "{}";'.format(PSWD),
            r'document.querySelector("#mat-input-0").dispatchEvent(new Event("input", { bubbles : true }));',
            # Login
            r'document.getElementById("signIn-btn").click();'
        ]
        for i in cmds:
            b.execute_script(i)
        return True
    except Exception as ex:
        print(ex.with_traceback(None))
        return False


def init_org(b):
    """
    @description 初始化IcDoc_Org配置
    @param b: webdriver
    @return: bool
    @author:lzt
    @time: 2025/12/9
    """
    try:
        # select type:ios
        cmds = [
            # 点击“研究方向”
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[3]/ul/li[15]/button/span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue?.click();",
            # 点击“学科分类体系”下拉搜索
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[1]/div/cl-select-filter/div/fieldset/div/span/nav/ul/li/a', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue?.click();",
            # 点击“esi”选项
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[1]/div/cl-select-filter/div/fieldset/div/span/nav/ul/li/ul/li[3]/label', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue?.click();",
            # 点击确认
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue?.click();",
        ]
        for i in cmds:
            b.execute_script(i)
        time.sleep(3)
        cmds = [
            # cookie
            # r"document.querySelector('#onetrust-accept-btn-handler').click();",
            # select years
            r"document.querySelector('#analysis-sidebar-filters > ic-sidebar-filters > div.ic-sidebar-filters > div.ic-analysis-sidebar__header > fieldset > ic-date-range-filter > div > span > nav > ul > li > ul > li:nth-child(4) > label').click();",
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/div/span/cl-range-input/div/div/div[1]/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value='{}';".format(
                INI_YEAR),
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/div/span/cl-range-input/div/div/div[2]/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value='{}';".format(
                END_YEAR),
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/div/span/cl-range-input/div/div/div[1]/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles : true }));",
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/div/span/cl-range-input/div/div/div[2]/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles : true }));",
        ]
        for i in cmds:
            b.execute_script(i)
        cmds = [
            # select source
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[1]/div[3]/ul/li[5]/button/span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();",
            r"document.querySelector('#multiselect-location').focus();",
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value='CHINA MAINLAND';",
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles : true }));",
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/div/ul/li[1]/cl-text-highlight/span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();",
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();",
        ]
        for index, cmd in enumerate(cmds):
            if index == 4:
                time.sleep(3)
            b.execute_script(cmd)
        # Wait for blank
        e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.CLASS_NAME, "cl-sortable-table__row__value"))
        # Click accept cookie button
        try:
            b.execute_script(r"document.querySelector('#onetrust-accept-btn-handler').click();")
        except Exception as e:
            pass
        # Add target
        b.execute_script(
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/section/div/div/div[1]/div/ic-add-indicator-menu/cl-popover-modal/div/cl-popover-modal-button/button', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();")
        e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.CLASS_NAME, "cl-popover-modal__body"))
        fs = range(1, 8)
        excl = [6]
        a = '/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/section/div/div/div[1]/div/ic-add-indicator-menu/cl-popover-modal/div/cl-popover-modal-content/div/div/div[2]/fieldset[{}]'
        for i in fs:
            s = a.format(i)
            e = b.find_element(By.XPATH, s)
            check_boxes = e.find_elements(By.TAG_NAME, "input")
            for box in check_boxes:
                if i not in excl:
                    if not box.is_selected():
                        box.send_keys(Keys.SPACE)
                        # print(b.is_selected())
                else:
                    if box.is_selected():
                        box.send_keys(Keys.SPACE)
                        # print(b.is_selected())
        cmds = [
            r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/section/div/div/div[1]/div/ic-add-indicator-menu/cl-popover-modal/div/cl-popover-modal-content/div/div/div[3]/button[2]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();",
        ]
        for i in cmds:
            b.execute_script(i)
        # Wait for blank
        e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.CLASS_NAME, "cl-sortable-table__row__value"))
        return True
    except Exception as ex:
        print(ex.with_traceback(None))
        return False


def init(b):
    try:
        if not login(b): return False

        # click 'accept all cookies
        try:
            e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.ID,
                                                                         r'onetrust-accept-btn-handler'))
        except Exception as e:
            print("未找到cookie元素，直接进入下一页")

        # ORG
        b.get(ORG_URL)
        # Wait for blank
        e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.CLASS_NAME, "cl-sortable-table__row__value"))
        if not init_org(b): return False
        return True
    except Exception as e:
        print(e.with_traceback(None))
        return False


def download(b, name):
    # Enter filename and download
    cmds = [
        # Enter filename
        r"document.querySelector('#multiselect-ic-analysis__header__search__input').value = String.raw `{}`;".format(name),
        r"document.querySelector('#multiselect-ic-analysis__header__search__input').dispatchEvent(new Event('input', { bubbles : true }));",
        r"setTimeout(function() {document.querySelector('#cl-multi-select__options__item--0').click();}, 1500);"
    ]
    for i in cmds:
        b.execute_script(i)
    # Wait for download label refresh
    time.sleep(3)
    e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.XPATH,
                                                                 r'/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/section/div/div/div[2]/div/cl-sortable-table/div[2]/table/tbody[2]/tr/td[2]/cl-sortable-table-cell/span/span/button'))
    e.click()
    # Wait for download window
    e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.XPATH,
                                                                 r'/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[2]/ic-document-list/div[1]/span[4]/ic-document-list-download/cl-popover-modal/div/cl-popover-modal-button/button'))
    cmds = [
        # Click download button
        r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[2]/ic-document-list/div[1]/span[4]/ic-document-list-download/cl-popover-modal/div/cl-popover-modal-button/button', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();",
        r"var filename = document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[1]/div[2]/h2/span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.textContent; document.querySelector('#file-name').value = filename;",
        r"document.querySelector('#file-name').dispatchEvent(new Event('input', { bubbles : true }));",
    ]
    for i in cmds:
        b.execute_script(i)
    # Wait for popped window
    e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.CLASS_NAME, "cl-popover-modal"))
    cmds = [
        # Enter filename to download
        r"var filename = document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[1]/div[2]/h2/span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.textContent; document.querySelector('#file-name').value = filename;",
        r"document.querySelector('#file-name').dispatchEvent(new Event('input', { bubbles : true }));",
        r"var selectElement = document.querySelector('select[formcontrolname=\"downloadOption\"]'); selectElement.value = 'csv'; selectElement.dispatchEvent(new Event('change', { bubbles: true }));",
        r"document.evaluate('/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[2]/ic-document-list/div[1]/span[4]/ic-document-list-download/cl-popover-modal/div/cl-popover-modal-content/div/div[2]/form/fieldset/div[3]/button', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();"
    ]
    for i in cmds:
        b.execute_script(i)
    # Until download button appear, which means download completed
    e = WebDriverWait(b, DOWNLOAD_TIMEOUT).until(lambda x: x.find_element(By.XPATH,
                                                                         r'/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[2]/ic-document-list/div[1]/span[4]/ic-document-list-download/cl-popover-modal/div/cl-popover-modal-button/button'))
    # Close download window
    e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.XPATH,
                                                                 r'/html/body/app-root/main/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[1]/button'))
    # e.click()
    # Refresh
    b.refresh()
    # Close tag
    e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.XPATH,
                                                                 r'/html/body/app-root/main/ic-analysis/div/div[1]/cl-applied-filters/div/ul/li[2]/span/button'))
    e.click()
    # Refresh
    b.refresh()
    # Wait for blank
    e = WebDriverWait(b, TIMEOUT).until(lambda x: x.find_element(By.CLASS_NAME, "cl-sortable-table__row__value"))


def main():
    print("Enter file path: ", end="")
    p = input()
    f = open(p, "r", encoding="latin-1")
    name_list = [i.rstrip() for i in f.readlines() if i.rstrip() != ""]
    f.close()

    bs = [BROWSER()]
    while not init(bs[0]):
        bs[0].close()
        bs[0] = BROWSER()
    fa = 0
    leng = len(name_list)
    c = 0
    for name in name_list:
        b = bs[0]
        c += 1
        name = name.replace('"', '')
        try:
            print("({}/{}) Downloading {}".format(c, leng, name))
            download(b, name)
        except Exception as exc:
            fa += 1
            print(exc)
            print("({}/{}) Download error at {}".format(c, leng, name))
            b.close()
            bs[0] = BROWSER()
            while (not init(bs[0])):
                bs[0].close()
                bs[0] = BROWSER()

    print("Download completed, {} failed, {} succeeded, {} totally".format(fa, leng - fa, leng))
    time.sleep(5)


if __name__ == "__main__":
    main()