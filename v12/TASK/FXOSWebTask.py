# FXOS WEB 巡检任务

# 导入标准库
from typing import Dict

# 导入第三方库
from playwright.sync_api import sync_playwright

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG, DEFAULT_PAGE_GOTO_TIMEOUT, DEFAULT_SELECTOR_TIMEOUT, BLOCK_RES_TYPES, require_keys, decrypt_password

# FXOS WEB巡检任务类：通过浏览器自动化技术对FXOS设备进行WEB界面巡检
class FXOSWebTask(BaseTask):
    
    # 初始化FXOS WEB巡检任务：设置登录凭据、设备URL列表和自动化参数
    def __init__(self):
        super().__init__("FXOS WEB巡检")
        
        # 验证FXOSWebTask专用配置
        require_keys(CONFIG, ["FXOSWebTask"], "root")
        require_keys(CONFIG["FXOSWebTask"], ["username", "password", "devices"], "FXOSWebTask")
        
        FXOS_CFG = CONFIG["FXOSWebTask"]
        self.USERNAME = FXOS_CFG["username"]
        self.PASSWORD = decrypt_password(FXOS_CFG["password"])
        self.DEVICE_URLS: Dict[str, str] = FXOS_CFG["devices"]
        self.EXPECTED_XPATH = (
            '/html/body/div[6]/div/div/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/'
            'table/tbody/tr/td[1]/table/tbody/tr/td[5]/div'
        )
        self.AUTO_PRESS_ENTER: bool = bool(FXOS_CFG.get("auto_press_enter", False))
        self.ENTER_RETRIES: int = int(FXOS_CFG.get("enter_retries", 5))
        self.ENTER_INTERVAL_MS: int = int(FXOS_CFG.get("enter_interval_ms", 400))

    # 返回要巡检的FXOS设备URL列表
    def items(self):
        return list(self.DEVICE_URLS.items())

    # 自动处理页面继续按钮和对话框：通过键盘操作和元素点击自动跳过确认步骤
    def _NUDGE_CONTINUE(self, PAGE) -> None:
        if not self.AUTO_PRESS_ENTER:
            return
        try:
            # 处理页面对话框：自动接受弹出的对话框
            def _ON_DIALOG(DIALOG):
                try:
                    DIALOG.accept()
                except Exception:
                    pass

            PAGE.once("dialog", _ON_DIALOG)
        except Exception:
            pass

        for _ in range(self.ENTER_RETRIES):
            try:
                PAGE.keyboard.press("Enter")
            except Exception:
                pass
            PAGE.wait_for_timeout(self.ENTER_INTERVAL_MS)
            try:
                PAGE.keyboard.press("Tab")
                PAGE.keyboard.press("Enter")
            except Exception:
                pass
            PAGE.wait_for_timeout(self.ENTER_INTERVAL_MS)

            for SELECTOR in (
                    "text=Continue", "text=Proceed", "text=OK", "text=Confirm",
                    "text=继续", "text=确认", "text=确定",
                    "xpath=//button[contains(.,'Continue') or contains(.,'Proceed') or contains(.,'OK') or contains(.,'确认') or contains(.,'继续') or contains(.,'确定')]",
                    "xpath=//a[contains(.,'Continue') or contains(.,'Proceed') or contains(.,'OK') or contains(.,'确认') or contains(.,'继续') or contains(.,'确定')]",
            ):
                try:
                    ELEMENT = PAGE.query_selector(SELECTOR)
                    if ELEMENT:
                        ELEMENT.click()
                        PAGE.wait_for_timeout(self.ENTER_INTERVAL_MS)
                except Exception:
                    pass

    # 执行单个FXOS设备的WEB巡检：自动登录并验证页面加载
    def run_single(self, item):
        DEVICE_NAME, URL = item
        # 从设备名中提取站点名（如HX00-FXOS-01 -> HX00）
        SITE_NAME = DEVICE_NAME.split('-')[0] if '-' in DEVICE_NAME else DEVICE_NAME
        with sync_playwright() as PLAYWRIGHT:
            BROWSER = PLAYWRIGHT.chromium.launch(headless=True)
            CONTEXT = BROWSER.new_context(ignore_https_errors=True)
            # ↓↓↓ 新增：统一超时 & 拦截图片/媒体/字体请求，降低负载
            CONTEXT.set_default_timeout(DEFAULT_PAGE_GOTO_TIMEOUT)
            CONTEXT.route("**/*", lambda
                ROUTE: ROUTE.abort() if ROUTE.request.resource_type in BLOCK_RES_TYPES else ROUTE.continue_())

            PAGE = CONTEXT.new_page()
            try:
                PAGE.goto(URL, timeout=DEFAULT_PAGE_GOTO_TIMEOUT)
                try:
                    PAGE.wait_for_selector('xpath=/html/body/center/div/form/a[1]', timeout=5000)
                    PAGE.click('xpath=/html/body/center/div/form/a[1]')
                except Exception:
                    pass

                self._NUDGE_CONTINUE(PAGE)

                PAGE.fill('xpath=/html/body/center/div/form/div[3]/input[1]', self.USERNAME)
                PAGE.fill('xpath=/html/body/center/div/form/div[3]/input[2]', self.PASSWORD)
                PAGE.click('xpath=/html/body/center/div/form/a[2]')

                self._NUDGE_CONTINUE(PAGE)

                PAGE.wait_for_selector(f'xpath={self.EXPECTED_XPATH}', timeout=DEFAULT_SELECTOR_TIMEOUT)
                self.add_result(Level.OK, f"站点{SITE_NAME}防火墙{DEVICE_NAME}网页登录成功")
            except Exception as ERROR:
                self.add_result(Level.WARN, f"站点{SITE_NAME}防火墙{DEVICE_NAME}网页登录失败: {ERROR}")
            finally:
                try:
                    CONTEXT.close()
                except Exception:
                    pass
                BROWSER.close()
