# @Author:慕白
import json
import os
import random
import time
from pathlib import Path
from typing import Union, Tuple
import cv2
import numpy as np
from PIL import Image
from ddddocr import DdddOcr
from selenium.common import NoSuchElementException
from selenium.webdriver import Chrome, ActionChains
from selenium.webdriver import ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class OpenError(Exception):
    """
    打开网页异常
    """
    ...


class BasePage:
    """
    Selenium封装
    """

    def __init__(self, display: bool = True):
        """
        驱动谷歌浏览器
        """
        if not isinstance(display, bool):
            raise TypeError('display must be a boolean')
        options = ChromeOptions()
        # 开发者模式
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        # 取消自动化控制语句
        options.add_argument('--disable-blink-features=AutomationControlled')
        prefs = {
            # 取消浏览器弹窗
            'profile.default_content_setting_values': {'notifications': 2},
            # 取消保存密码提示框
            'credentials_enable_service': False,
            'profile.password_manager_enabled': False
        }
        options.add_experimental_option('prefs', prefs)
        # 跳过安全证书验证
        options.set_capability('acceptInsecureCerts', True)
        if display:
            # 结束后保留浏览器页面
            options.add_experimental_option('detach', True)
            self.driver = Chrome(options=options, service=Service(ChromeDriverManager().install()))
        else:
            # 隐藏浏览器页面
            options.add_argument('--headless')
            self.driver = Chrome(options=options, service=Service(ChromeDriverManager().install()))
        # 规避检测
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                                      Object.defineProperty(navigator,'webdriver',{
                                        get: () => undefined
                                      })
                                      """
        })
        self.driver.maximize_window()

    def open(self, url: str):
        """
        打开网页
        """
        if not isinstance(url, str):
            raise TypeError(f'{url} must be a string')
        try:
            self.driver.get(url)
        except Exception:
            raise OpenError(f'failed to open {url}')

    def cookie_login(self, domain=True):
        """
        cookie登录
        """
        if not isinstance(domain, bool):
            raise TypeError(f'{domain} must be a boolean')
        if not os.path.exists('./cookie'):
            os.mkdir('./cookie')
        if not os.path.exists('./cookie/cookie.json'):
            with open('./cookie/cookie.json', 'w') as f:
                json.dump([], f)
        self.driver.delete_all_cookies()
        with open('./cookie/cookie.json', 'r') as f:
            cookie_list = json.load(f)
            for cookie in cookie_list:
                if "expiry" in cookie.keys():
                    cookie.pop("expiry")
                if domain:
                    if "domain" in cookie.keys():
                        cookie.pop("domain")
                self.driver.add_cookie(cookie)
        self.driver.refresh()

    def wait(self, times: Union[int, float]):
        """
        隐式加载
        """
        if not isinstance(times, (int, float)):
            raise TypeError(f'{times} must be an integer or a float')
        self.driver.implicitly_wait(times)

    def wait_element(self, element: Tuple[str, str]):
        """
        显示加载
        """
        WebDriverWait(self.driver, 10).until(lambda _: self.position(element))

    def wait_alert(self):
        """
        等待弹窗加载
        """
        WebDriverWait(self.driver, 10).until(expected_conditions.alert_is_present())

    @staticmethod
    def stop(times: Union[int, float]):
        """
        等待加载
        """
        if not isinstance(times, (int, float)):
            raise TypeError(f'{times} must be an integer or a float')
        time.sleep(times)

    def position(self, element: Tuple[str, str]):
        """
        定位单一元素
        """
        if not isinstance(element, tuple):
            raise TypeError(f'{element} must be a (string,string)')
        return self.driver.find_element(*element)

    def positions(self, element: Tuple[str, str]):
        """
        定位多个元素
        """
        if not isinstance(element, tuple):
            raise TypeError(f'{element} must be a (string,string)')
        return self.driver.find_elements(*element)

    def position_list(self, element: Tuple[str, str]):
        """
        定位下拉选项框
        """
        return Select(self.position(element))

    def click(self, element: Tuple[str, str]):
        """
        点击元素
        """
        self.position(element).click()

    @staticmethod
    def click_web_element(element: WebElement):
        """
        点击web元素
        """
        if not isinstance(element, WebElement):
            raise TypeError(f'{element} must be a WebElement')
        element.click()

    def click_list(self, element: Tuple[str, str], index: int):
        """
        点击下拉选项框的内容
        """
        if not isinstance(index, int):
            raise TypeError(f'{index} must be an integer')
        options = self.position_list(element)
        if index < -len(options.options) - 1 or index > len(options.options):
            raise ValueError(f'{index} must be between -{self.get_pages() - 1} '
                             f'and -1 or between 0 and {self.get_pages()}')
        options.select_by_index(index)

    def submit(self, element: Tuple[str, str]):
        """
        回车确认
        """
        self.position(element).submit()

    def input(self, element: Tuple[str, str], content: str):
        """
        输入内容
        """
        if not isinstance(content, str):
            raise TypeError(f'{content} must be a string')
        self.position(element).send_keys(content)

    def clear(self, element: Tuple[str, str]):
        """
        清除内容
        """
        self.position(element).clear()

    def get_html(self):
        """
        获取源代码
        """
        return self.driver.page_source

    def get_title(self):
        """
        获取标题
        """
        return self.driver.title

    def get_pages(self):
        """
        获取窗口个数
        """
        return len(self.driver.window_handles)

    def get_len(self, element: Tuple[str, str]):
        """
        获取元素长度
        """
        return len(self.positions(element))

    def switch_to_front_page(self):
        """
        切换到上一个页面
        """
        self.driver.back()

    def switch_to_next_page(self):
        """
        切换到下一个页面
        """
        self.driver.forward()

    def switch_to_frame(self, frame: Union[int, tuple]):
        """
        切换到frame弹窗
        """
        if type(frame) is int:
            self.driver.switch_to.frame(frame)
            return
        self.driver.switch_to.frame(self.position(frame))

    def switch_to_forward_frame(self):
        """
        切换到上一个frame弹窗
        """
        self.driver.switch_to.parent_frame()

    def switch_to_main_page(self):
        """
        切换到主页面
        """
        self.driver.switch_to.default_content()

    def switch_page(self, index: int):
        """
        切换页面
        """
        if not isinstance(index, int):
            raise TypeError(f'{index} must be an integer')
        handles = self.driver.window_handles
        if index < -len(handles) - 1 or index > len(handles):
            raise ValueError(f'{index} must be between -{len(handles) - 1} '
                             f'and -1 or between 0 and {len(handles)}')
        self.driver.switch_to.window(handles[index])

    def switch_to_alert(self):
        """
        切换到alert弹窗
        """
        return self.driver.switch_to.alert

    def accept_alert(self):
        """
        接受alert弹窗
        """
        self.switch_to_alert().accept()

    def dismiss_alert(self):
        """
        拒绝alert弹窗
        """
        self.switch_to_alert().dismiss()

    def input_alert(self, content: str):
        """
        在alert弹窗中输入内容
        """
        if not isinstance(content, str):
            raise TypeError(f'{content} must be a string')
        self.switch_to_alert().send_keys(content)

    def close(self, index: int = -1):
        """
        关闭页面
        """
        self.switch_page(index)
        self.driver.close()

    def quit(self):
        """
        退出浏览器
        """
        self.driver.quit()

    def if_exist(self, element: Tuple[str, str]):
        """
        判断元素是否存在
        """
        try:
            self.position(element)
            return True
        except NoSuchElementException:
            return False

    def if_alert_exist(self):
        """
        判断alert弹窗是否存在
        """
        return expected_conditions.alert_is_present()(self.driver)

    def if_show(self, element: Tuple[str, str]):
        """
        判断元素是否可见
        """
        return self.position(element).is_displayed()

    def if_click(self, element: Tuple[str, str]):
        """
        判断元素是否可点击
        """
        return self.position(element).is_enabled()

    def if_selected(self, element: Tuple[str, str]):
        """
        判断元素是否被选中
        """
        return self.position(element).is_selected()

    def click_frame(self, frame: Tuple[str, str]):
        """
        点击frame弹窗
        """
        if not isinstance(frame, Tuple):
            raise TypeError(f'{frame} must be a (string,string)')
        WebDriverWait(self.driver, 1000).until(
            expected_conditions.element_to_be_clickable
            (frame)).click()

    def get_security_code(self, element: Tuple[str, str], dpi: float = 1.5):
        """
        获取验证码
        """
        if not isinstance(dpi, float):
            raise TypeError(f'{dpi} must be a float')
        if not os.path.exists('./images'):
            os.mkdir('./images')
        self.save_screenshot('./images/page.png')
        code = self.position(element)
        loc = code.location
        size = code.size
        left = loc['x'] * dpi
        top = loc['y'] * dpi
        right = loc['x'] * dpi + size['width'] * dpi
        bottom = loc['y'] * dpi + size['height'] * dpi
        val = (left, top, right, bottom)
        page_img = Image.open('./images/page.png')
        security_code_img = page_img.crop(val)
        security_code_img.save('./images/security_code.png')
        ocr = DdddOcr(old=True, show_ad=False)
        with open('./images/security_code.png', 'rb') as f:
            security_code = ocr.classification(f.read())
        return security_code

    def get_slider_distance(self, slider: Tuple[str, str], background: Tuple[str, str], dpi: float = 1.5):
        """
        获取滑块距离
        """
        if not isinstance(dpi, float):
            raise TypeError(f'{dpi} must be a float')
        if not os.path.exists('./images'):
            os.mkdir('./images')
        self.save_screenshot('./images/page.png')
        sli = self.position(slider)
        bg = self.position(background)
        slider_loc = sli.location
        slider_size = sli.size
        slider_left = slider_loc['x'] * dpi
        slider_top = slider_loc['y'] * dpi
        slider_right = slider_loc['x'] * dpi + slider_size['width'] * dpi
        slider_bottom = slider_loc['y'] * dpi + slider_size['height'] * dpi
        slider_val = (slider_left, slider_top, slider_right, slider_bottom)
        bg_loc = bg.location
        bg_size = bg.size
        bg_left = bg_loc['x'] * dpi
        bg_top = bg_loc['y'] * dpi
        bg_right = bg_loc['x'] * dpi + bg_size['width'] * dpi
        bg_bottom = bg_loc['y'] * dpi + bg_size['height'] * dpi
        bg_val = (bg_left, bg_top, bg_right, bg_bottom)
        page_img = Image.open('./images/page.png')
        slider_img = page_img.crop(slider_val)
        slider_img.save('./images/slider.png')
        bg_img = page_img.crop(bg_val)
        bg_img.save('./images/bg.png')
        with open('./images/slider.png', 'rb') as f:
            slider_img_bytes = f.read()
        with open('./images/bg.png', 'rb') as f:
            bg_img_bytes = f.read()
        ocr = DdddOcr(det=False, ocr=False, show_ad=False)
        result = ocr.slide_match(slider_img_bytes, bg_img_bytes, simple_target=True)
        distance = result['target'][0]
        return distance

    def get_slider_distance1(self, slider: Tuple[str, str], background: Tuple[str, str], dpi: float = 1.5):
        """
        获取滑块距离
        """
        if not isinstance(dpi, float):
            raise TypeError(f'{dpi} must be a float')
        if not os.path.exists('./images'):
            os.mkdir('./images')
        self.save_screenshot('./images/page.png')
        sli = self.position(slider)
        bg = self.position(background)
        slider_loc = sli.location
        slider_size = sli.size
        slider_left = slider_loc['x'] * dpi
        slider_top = slider_loc['y'] * dpi
        slider_right = slider_loc['x'] * dpi + slider_size['width'] * dpi
        slider_bottom = slider_loc['y'] * dpi + slider_size['height'] * dpi
        slider_val = (slider_left, slider_top, slider_right, slider_bottom)
        bg_loc = bg.location
        bg_size = bg.size
        bg_left = bg_loc['x'] * dpi
        bg_top = bg_loc['y'] * dpi
        bg_right = bg_loc['x'] * dpi + bg_size['width'] * dpi
        bg_bottom = bg_loc['y'] * dpi + bg_size['height'] * dpi
        bg_val = (bg_left, bg_top, bg_right, bg_bottom)
        page_img = Image.open('./images/page.png')
        slider_img = page_img.crop(slider_val)
        slider_img.save('./images/slider.png')
        bg_img = page_img.crop(bg_val)
        bg_img.save('./images/bg.png')
        slider_pic = cv2.imread('./images/slider.png', 0)
        background_pic = cv2.imread('./images/bg.png', 0)
        cv2.imwrite('./images/slider_grey.png', background_pic)
        cv2.imwrite('./images/bg_grey.png', slider_pic)
        slider_pic = cv2.imread('./images/slider_grey.png')
        slider_pic = cv2.cvtColor(slider_pic, cv2.COLOR_BGR2GRAY)
        slider_pic = abs(255 - slider_pic)
        cv2.imwrite('./images/slider_grey.png', slider_pic)
        slider_pic = cv2.imread('./images/slider_grey.png')
        background_pic = cv2.imread('./images/bg_grey.png')
        result = cv2.matchTemplate(slider_pic, background_pic, cv2.TM_CCOEFF_NORMED)
        top, left = np.unravel_index(result.argmax(), result.shape)
        left = int(left)
        return left

    def get_cookie(self):
        """
        获取cookie
        """
        if not os.path.exists('./cookie'):
            os.mkdir('./cookie')
        with open('./cookie/cookie.json', 'w+') as f:
            json.dump(self.driver.get_cookies(), f, ensure_ascii=False)
        return self.driver.get_cookies()

    def execute_js(self, js: str, element: WebElement = None):
        """
        执行js代码
        输入值:'document.getElementById("xxx").value="xxx"'/'document.getElementsByClassName/Id("xxx")[x].value="xxx"'
        点击元素:'document.getElementById("xxx").click()' / 'document.getElementsByClassName/Id("xxx")[0].click()'
        显示隐藏元素:'document.getElementById("xxx").style.display="block"'/'document.getElementsByClassName/Id("xxx")[x].style.display="block"'
        获取长度:'return document.getElementById("xxx").length'/return document.getElementsByClassName/Id("xxx").length'
        修改元素属性:'document.getElementById("xxx").setAttribute("xxx","xxx")/'document.getElementsByClassName/Id("xxx")[x].setAttribute("xxx","xxx")'
        """
        if not isinstance(js, str):
            raise TypeError(f'{js} must be a string')
        if element:
            return self.driver.execute_script(js, element)
        return self.driver.execute_script(js)

    def js_click(self, element: Tuple[str, str]):
        """
        js点击元素
        """
        js = 'arguments[0].click()'
        self.execute_js(js, self.position(element))

    def js_input(self, element: Tuple[str, str], value: str):
        """
        js给元素输入值
        """
        if not isinstance(value, str):
            raise TypeError(f'{value} must be a string')
        js = f'arguments[0].value="{value}"'
        self.execute_js(js, self.position(element))

    def js_modify(self, element: Tuple[str, str], attribute: str, value: Union[str, int]):
        """
        js修改元素属性
        """
        if not isinstance(attribute, str):
            raise TypeError(f'{attribute} must be a string')
        if not isinstance(value, (str, int)):
            raise TypeError(f'{value} must be a string or an integer')
        if isinstance(value, str):
            js = f'arguments[0].setAttribute("{attribute}","{value}")'
            self.execute_js(js, self.position(element))
        if isinstance(value, int):
            js = f'arguments[0].setAttribute("{attribute}",{value})'
            self.execute_js(js, self.position(element))

    def scroll_load(self, page: int):
        """
        滚动加载
        """
        if not isinstance(page, int):
            raise TypeError(f'{page} must be an integer')
        flag = 0
        while True:
            self.scroll_to_bottom()
            self.stop(1)
            flag += 1
            if flag == page:
                break

    def scroll_until_exist(self, element: Tuple[str, str]):
        """
        滚动页面直到元素出现
        """
        js = 'arguments[0].scrollIntoView(false)'
        self.execute_js(js, self.position(element))

    def scroll_until_web_element_exist(self, element: WebElement):
        """
        滚动页面直到web元素出现
        """
        js = 'arguments[0].scrollIntoView(false)'
        self.execute_js(js, element)

    def scroll_slow_to_bottom(self):
        """
        缓慢滑动页面到底部
        """
        js = '''{let he=   setInterval(()=>{
        document.documentElement.scrollTop+=30;
        if(document.documentElement.scrollTop>=(document.documentElement.scrollHeight-document.documentElement.scrollWidth)){clearInterval(he);}},100);}'''
        self.execute_js(js)

    def scroll_slow_to_top(self):
        """
        缓慢滑动页面到顶部
        """
        js = '''{let he=   setInterval(()=>{
                document.documentElement.scrollTop-=30;
                if(document.documentElement.scrollTop==0){clearInterval(he);}},100);}'''
        self.execute_js(js)

    def scroll_to_bottom(self):
        """
        滚动页面到底部
        """
        js = 'window.scrollTo(0,document.body.scrollHeight)'
        self.execute_js(js)

    def scroll_to_top(self):
        """
        滚动页面到顶部
        """
        js = 'window.scrollTo(0,document.body.scrollTop)'
        self.execute_js(js)

    def show_display(self, element: Tuple[str, str]):
        """
        显示隐藏元素
        """
        js = 'arguments[0].style.display="block"'
        self.execute_js(js, self.position(element))

    def alert_warning(self, content: str):
        """
        弹窗提示
        """
        if not isinstance(content, str):
            raise TypeError(f'{content} must be a string')
        self.execute_js(f"alert('{content}')")

    def refresh(self):
        """
        刷新网页
        """
        self.driver.refresh()

    def save_screenshot(self, file: Union[str, Path]):
        """
        保存屏幕截图
        """
        if not isinstance(file, (str, Path)):
            raise TypeError(f'{file} must be a string or a Path')
        self.driver.save_screenshot(file)

    def move(self, x: Union[int, float] = 0, y: Union[int, float] = 0, element: Tuple[str, str] = None):
        """
        移动鼠标坐标
        """
        if not isinstance(x, (int, float)):
            raise TypeError(f'{x} must be an integer or a float')
        if not isinstance(y, (int, float)):
            raise TypeError(f'{y} must be an integer or a float')
        action = ActionChains(self.driver)
        if element is None:
            action.move_by_offset(x, y)
            action.perform()
            return
        action.move_to_element(self.position(element))
        action.perform()

    def move_slider(self, slider: Tuple[str, str], distance: Union[int, float]):
        """
        移动滑块
        """
        if not isinstance(distance, (int, float)):
            raise TypeError(f'{distance} must be an integer or a float')
        self.click_left_hold(slider)
        for _ in range(7):
            self.move(random.uniform(distance / 8, distance / 7.5), 0)
            self.stop(random.random())
        self.release_left()

    def click_left(self, element: Tuple[str, str] = None):
        """
        点击鼠标左键
        """
        action = ActionChains(self.driver)
        if element is None:
            action.click()
            action.perform()
            return
        action.move_to_element(self.position(element)).click(self.position(element))
        action.perform()

    def click_right(self, element: Tuple[str, str] = None):
        """
        点击鼠标右键
        """
        action = ActionChains(self.driver)
        if element is None:
            action.context_click()
            action.perform()
            return
        action.move_to_element(self.position(element)).context_click(self.position(element))
        action.perform()

    def click_left_hold(self, element: Tuple[str, str] = None):
        """
        点击鼠标左键,不松开
        """
        action = ActionChains(self.driver)
        if element is None:
            action.click_and_hold()
            action.perform()
            return
        action.move_to_element(self.position(element)).click_and_hold(self.position(element))
        action.perform()

    def click_double_left(self, element: Tuple[str, str] = None):
        """
        双击鼠标左键
        """
        action = ActionChains(self.driver)
        if element is None:
            action.double_click()
            action.perform()
            return
        action.move_to_element(self.position(element)).double_click(self.position(element))
        action.perform()

    def drag_to(self, element1: Tuple[str, str], element2: Tuple[str, str]):
        """
        把元素拖拽到元素上
        """
        action = ActionChains(self.driver)
        action.drag_and_drop(self.position(element1), self.position(element2))
        action.perform()

    def drag(self, element: Tuple[str, str], x: int, y: int):
        """
        拖拽元素
        """
        if not isinstance(x, int):
            raise TypeError(f'{x} must be an integer or a float')
        if not isinstance(y, int):
            raise TypeError(f'{y} must be an integer or a float')
        action = ActionChains(self.driver)
        action.move_to_element(self.position(element)).drag_and_drop_by_offset(self.position(element), x, y)
        action.perform()

    def release_left(self, element: Tuple[str, str] = None):
        """
        松开鼠标左键
        """
        action = ActionChains(self.driver)
        if element is None:
            action.release()
            action.perform()
            return
        action.move_to_element(self.position(element)).release(self.position(element))
        action.perform()

    def remove(self):
        """
        清除鼠标操作
        """
        action = ActionChains(self.driver)
        action.reset_actions()

    def dubug(self):
        """
        调试
        """
        js = 'debugger'
        self.execute_js(js)
