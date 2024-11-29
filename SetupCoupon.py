import time
import pyautogui
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotInteractableException, TimeoutException, NoSuchElementException


# 打开浏览器对应网页
def open_browser(url):
    # 配置 Chrome 下载路径
    chrome_options = Options()
    # 启用无头模式（如果需要）
    # chrome_options.add_argument("--headless")
    # 设置用户数据目录
    chrome_options.add_argument(r'--user-data-dir=C:\Users\83825\AppData\Local\Google\Chrome\User Data')
    # 排除自动化的标记
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])

    # 启动 Chrome 浏览器，使用webdriver-manager来自动管理驱动
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # 打开目标网页
    driver.get(url)

    # 等待页面加载
    time.sleep(3)

    return driver


# 关闭浏览器
def close_browser(driver):
    # 关闭浏览器
    driver.quit()


# 遇到问题首先进行三轮刷新操作，若还是无法正常进行则报错
def retry_operation(driver, operation, max_retries=3, *args, **kwargs):
    """
    通用的重试机制，执行某个操作并在失败时刷新页面重试，最多尝试 max_retries 次。

    :param driver: 浏览器驱动对象
    :param operation: 要执行的操作函数
    :param max_retries: 最大刷新次数
    :param *args: 传递给操作函数的位置参数
    :param **kwargs: 传递给操作函数的关键字参数
    :return: 操作成功的返回值
    :raise Exception: 如果重试超过最大次数仍然失败，抛出异常
    """
    attempts = 0

    while attempts < max_retries:
        try:
            # 执行操作
            return operation(driver, *args, **kwargs)

        except Exception as e:
            # 如果操作失败，刷新页面并增加尝试次数
            print(f"尝试 {attempts + 1} 失败，刷新页面...")
            print(f"错误类型: {e.__class__.__name__}")
            time.sleep(3)
            driver.refresh()
            time.sleep(3)  # 等待页面刷新完成
            attempts += 1

    # 如果尝试超过最大次数仍然没有成功，抛出异常
    raise Exception(f"操作失败，已尝试刷新 {max_retries} 次，操作仍然失败。")


def click_button(driver, xpath):
    """
    点击按钮的操作
    """
    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    driver.execute_script("arguments[0].click();", button)
    print(f"成功点击按钮：{xpath}")
    time.sleep(3)  # 等待页面刷新完成


def fill_input(driver, xpath, input_text):
    """
    填写输入框的操作
    """
    input_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
    try:
        # 全选并删除内容
        driver.execute_script("arguments[0].click();", input_field)  # 确保输入框处于焦点
        input_field.send_keys(Keys.CONTROL + "a")  # 全选文本
        input_field.send_keys(Keys.DELETE)  # 删除选中的内容

        # 填入文本
        input_field.send_keys(input_text)
        print(f"成功填写输入框：{xpath}")
    # 若输入框不可直接用send_keys填写，则模拟键盘逐字输入
    except ElementNotInteractableException:
        print("ElementNotInteractableException: 使用模拟逐字输入")

        # 创建 ActionChains 实例
        actions = ActionChains(driver)

        # 确保焦点在输入框上
        actions.move_to_element(input_field).click().perform()

        # 模拟逐字符输入
        for char in input_text:
            actions.send_keys(char).perform()  # 每次输入一个字符并执行
            time.sleep(0.1)  # 适当延迟，模拟逐字输入
        print(f"成功填写输入框（模拟输入）：{xpath}")


# 模拟滑动元素
def scroll():
    try:
        # 由于无法定位到选择门店那个列表框，无法正常执行滑轮操作，但网页是动态加载，无法获取地区目标元素，只能采用简单粗暴的绝对定位进行模拟滑轮
        x = 1500
        y = 1000
        print(f"将鼠标移动到门店列表所在位置: X={x}, Y={y}")
        pyautogui.moveTo(x, y)
        pyautogui.scroll(-500)  # 向下滚动
        time.sleep(2)
        return True
    except Exception as e:
        print("滚动元素时出错:", e)
        return False


def unfold_and_check(driver, region, operation):
    """
    专门对门店选择列表的操作函数
    :param driver: 浏览器驱动对象
    :param region: 传入的地区-省份或地市名
    :param operation: 要执行的操作-展开列表unfold或是勾选check
    """
    # 定位到 title="region" 的 span 元素
    span = driver.find_element(By.XPATH, f'//span[@title="{region}"]')
    # 获取该 span 元素的父节点
    parent = span.find_element(By.XPATH, 'parent::*')
    # 若操作为展开列表unfold则span[2]，操作为勾选check则span[3]
    if operation == "unfold":
        o = 2
        print("成功展开省份地市列表")
    elif operation == "check":
        o = 3
        print("成功勾选对应地区")
    # 根据判断传入的操作类型决定要点击的区域
    operation_span = parent.find_element(By.XPATH, f'.//span[{o}]')
    driver.execute_script("arguments[0].click();", operation_span)
    time.sleep(1)


# 边滑动边找元素勾选
def scroll_tree_list(driver, region, operation):
    # 最大循环10次，找不到元素就报错
    scroll_count = 0
    while scroll_count < 10:
        try:
            # 直接勾选对应地区
            unfold_and_check(driver, region, operation)
            # 若成功勾选到目标，则退出循环
            break
        # 若找不到对应元素，说明信息有误或未加载，需要向下滑轮确认
        except NoSuchElementException:
            scroll()
            scroll_count += 1

        # 如果滑轮10次还是找不到元素则报错
        if scroll_count >= 10:
            raise NoSuchElementException(f"没有找到{region}选项框")


def coupon_rule(driver, coupon_list):
    # 填写优惠券名称
    # 若没有地市级则填写省级
    if coupon_list['地市'] == "" or "NaN":
        region = coupon_list['省份']
    else:
        region = coupon_list['地市']
    fill_input(driver, '//*[@id="activity_title"]', f"测试测试{coupon_list['品牌']}{coupon_list['机型']}"
                                                    f"（-{coupon_list['减多少']}）{region}")

    # 点击领取时间区域
    click_button(driver, '//*[@id="marketing-tool-modify-step-wrap"]/div[3]/div/div[2]/div[1]/div/div/form/div[1]/div/'
                         'div/div/div/div/div[2]/div/div[2]/div/div[2]/div/div/div')
    # 点击本月按钮
    click_button(driver, '//div[contains(text(), "本月")]')
    # 填写有效天数
    fill_input(driver, '//*[@id="marketing_base_info_valid_period"]', days_until_end_of_month)
    # 填写满减金额
    fill_input(driver, '//*[@id="marketing_base_info_threshold"]', int(coupon_list['满多少']))   # 满
    fill_input(driver, '//*[@id="marketing_base_info_credit"]', int(coupon_list['减多少']))   # 减
    # 填写单门店发放量
    fill_input(driver, '//*[@id="marketing_base_info_total_amount"]', 1)
    print("完成优惠券活动信息填写。")
    # 点击下一步
    click_button(driver, '//*[@id="marketing-tool-modify-step-wrap"]/div[4]/div/div/div[2]/div[1]/button')


def coupon_range(driver, coupon_list):
    # 点击添加商品按钮
    click_button(driver, '//*[@id="marketing-tool-modify-step-wrap"]/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div/button')
    # 输入商品id
    fill_input(driver, '//*[@id="product_ids"]/div[1]/div[2]', str(coupon_list['商品ID']))
    # 回车
    actions = ActionChains(driver)
    actions.send_keys(Keys.RETURN).perform()
    time.sleep(1)
    # 勾选商品
    item_checkbox = driver.find_elements(By.XPATH, f'//tr[@data-row-key="{coupon_list["商品ID"]}"]//input[@type="checkbox"]')[0]
    item_checkbox.click()
    # click_button(driver, f'//tr[@data-row-key="{coupon_list[4]}"]//input[@type="checkbox"]')  # 不希望等待太久所以弃用
    # 点击下一步
    click_button(driver, '//div[@class="auxo-drawer-footer-buttons"]/button[1]')

    """
    由于选择门店中的列表需要等滑动到对应位置才进行加载
    所以需要模拟循环模拟滑动直至出现对应地区的名称
    """
    # 判断是分地市还是整个省份
    if coupon_list['地市'] == "" or "NaN":
        # 若整个省份
        province = coupon_list['省份']
        # 滑动至对应省份处并勾选
        scroll_tree_list(driver, province, "check")

    else:
        # 若要分地市
        city = coupon_list['地市']
        province = coupon_list['省份']
        scroll_tree_list(driver, province, "unfold")
        scroll_tree_list(driver, city, "check")

    # 点击提交
    click_button(driver, '//div[@class="auxo-drawer-footer-buttons"]/button[1]')
    # 判断是否弹出提醒窗口
    try:
        click_button(driver, '//div[@class="auxo-modal-footer-btn-wrapper"]/button[1]')
    except TimeoutException:
        pass

    # 主界面点击提交
    click_button(driver, '//button[@class="auxo-btn auxo-btn-primary"]')


def setup_coupon(driver, coupon_list):
    # 点击立即新建按钮
    retry_operation(driver, click_button, 3, '//*[@id="marketing-root"]/div[1]/div[2]/div[1]/div[2]/button/a')
    # 填写优惠券活动信息
    retry_operation(driver, coupon_rule, 3, coupon_list)
    # 选择对应门店和商品
    retry_operation(driver, coupon_range, 3, coupon_list)
    print("当前券设置好了")
    time.sleep(10)


# 获取今天的日期
today = datetime.today()

# 获取下个月的第一天
first_day_of_next_month = (today + relativedelta(months=+1)).replace(day=1)

# 获取本月最后一天
last_day_of_month = first_day_of_next_month - timedelta(days=1)

# 计算今天到月底的天数
days_until_end_of_month = (last_day_of_month - today).days

