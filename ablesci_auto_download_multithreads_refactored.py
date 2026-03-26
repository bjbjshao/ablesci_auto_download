import sys
import time
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# =========================
# 配置
# =========================
BASE_URL = "https://www.ablesci.com/"
DOWNLOAD_DIR = Path.cwd() / "文献下载"
COOKIES_FILE = Path("cookies.txt")
LOG_FILE = Path("ablesci_run.log")

HEADLESS = True
WAIT_TIMEOUT = 20
LONG_WAIT_TIMEOUT = 28800
DOWNLOAD_SETTLE_TIMEOUT = 120
STARTUP_INTERVAL = 3
MAX_THREADS_LIMIT = 5

SELECTORS = {
    "nav_query_page": "body > div.able-header.header-bg-assist > div > div > a",
    "input_identifier": "#onekey",
    "submit_button": "#assist-create-form > div.alert.alert-success > div.layui-form-item.layui-row > div:nth-child(2) > div > button",
    "submit_confirm": "#layui-layer2 > div.layui-layer-btn.layui-layer-btn- > a.layui-layer-btn0",
    "ask_detail": "a.layui-layer-btn0",
    "article_title": "#LAY_ucm > div:nth-child(1) > div.assist-detail.layui-row > div > table > tbody > tr:nth-child(1) > td.assist-title > div:nth-child(1)",
    "download_link": "a[title='点击下载']",
    "accept_button": "#uploaded-file-handle > button:nth-child(1)",
    "accept_confirm": "#layui-layer2 > div.layui-layer-btn.layui-layer-btn- > a.layui-layer-btn0",
    "accept_finish": "#layui-layer4 > div.layui-layer-btn.layui-layer-btn- > a",
    "credits": "#user-point-now",
}

XPATHS = {
    "still_submit": "//a[text()='仍然提交']",
    "confirm_ok": "//a[text()='确定']",
}


# =========================
# 日志
# =========================
def setup_logger() -> logging.Logger:
    logger = logging.getLogger("ablesci")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


logger = setup_logger()


# =========================
# 数据结构
# =========================
@dataclass
class TaskResult:
    identifier: str
    success: bool
    article_title: str = ""
    message: str = ""


# =========================
# 工具函数
# =========================
def ensure_download_dir():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("下载目录: %s", DOWNLOAD_DIR)


def load_cookie_text() -> str:
    if not COOKIES_FILE.exists():
        raise FileNotFoundError(f"未找到 cookies 文件: {COOKIES_FILE}")
    text = COOKIES_FILE.read_text(encoding="utf-8").strip()
    if not text:
        raise ValueError("cookies.txt 为空")
    return text


def parse_cookie_text(cookie_text: str) -> List[dict]:
    cookies = []
    for part in [x.strip() for x in cookie_text.split(";") if x.strip()]:
        if "=" not in part:
            continue
        name, value = part.split("=", 1)
        name = name.strip()
        value = value.strip()
        if name:
            cookies.append({
                "name": name,
                "value": value,
                "domain": ".ablesci.com"
            })
    return cookies


def get_application_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def build_options() -> Options:
    options = Options()

    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.managed_default_content_settings.images": 2,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    return options


def create_driver() -> WebDriver:
    options = build_options()
    try:
        return webdriver.Chrome(options=options)
    except Exception as e:
        logger.warning("自动启动 ChromeDriver 失败: %s", e)

    possible_paths = [
        get_application_path() / "chromedriver.exe",
        Path.cwd() / "chromedriver.exe",
        Path("chromedriver.exe"),
    ]

    for path in possible_paths:
        if path.exists():
            service = Service(executable_path=str(path))
            return webdriver.Chrome(service=service, options=options)

    raise FileNotFoundError("ChromeDriver 未找到")


def safe_text(driver: WebDriver, css: str, default: str = "") -> str:
    try:
        return driver.find_element(By.CSS_SELECTOR, css).text.strip()
    except Exception:
        return default


def wait_for_no_partial_downloads(timeout: int = DOWNLOAD_SETTLE_TIMEOUT) -> bool:
    start = time.time()
    while time.time() - start < timeout:
        partials = list(DOWNLOAD_DIR.glob("*.crdownload"))
        if not partials:
            return True
        time.sleep(2)
    return False


def list_recent_files(limit: int = 10):
    files = []
    for file in DOWNLOAD_DIR.iterdir():
        if file.is_file():
            files.append((file.name, file.stat().st_ctime))
    files.sort(key=lambda x: x[1], reverse=True)
    return files[:limit]


# =========================
# 客户端
# =========================
class AbleSciClient:
    def __init__(self, driver: WebDriver, thread_name: str):
        self.driver = driver
        self.wait = WebDriverWait(driver, WAIT_TIMEOUT)
        self.short_wait = WebDriverWait(driver, 2)
        self.long_wait = WebDriverWait(driver, LONG_WAIT_TIMEOUT)
        self.thread_name = thread_name

    def log(self, msg: str, *args):
        logger.info("[%s] " + msg, self.thread_name, *args)

    def open_home(self):
        self.driver.get(BASE_URL)

    def inject_cookies(self, cookie_text: str):
        self.open_home()
        time.sleep(1)

        cookies = parse_cookie_text(cookie_text)
        if not cookies:
            raise ValueError("未解析到有效 cookies")

        for cookie in cookies:
            try:
                self.driver.add_cookie(cookie)
            except Exception as e:
                self.log("添加 cookie 失败: %s - %s", cookie.get("name"), e)

        self.driver.refresh()
        time.sleep(2)

    def click_css(self, selector: str):
        element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
        element.click()
        return element

    def js_click_xpath(self, xpath: str, long_wait: bool = False):
        waiter = self.long_wait if long_wait else self.wait
        element = waiter.until(EC.presence_of_element_located((By.XPATH, xpath)))
        self.driver.execute_script("arguments[0].click();", element)
        return element

    def maybe_click_xpath(self, xpath: str):
        try:
            element = self.short_wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            return False

    def go_to_query_page(self):
        self.log("进入查询页")
        self.click_css(SELECTORS["nav_query_page"])

    def submit_identifier(self, identifier: str):
        self.log("提交标识符: %s", identifier)

        input_box = self.wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, SELECTORS["input_identifier"]))
        )
        input_box.clear()
        input_box.send_keys(identifier)

        self.click_css(SELECTORS["submit_button"])
        time.sleep(2)
        self.click_css(SELECTORS["submit_confirm"])
        time.sleep(0.5)

        # 可选弹窗
        self.maybe_click_xpath(XPATHS["still_submit"])

    def open_detail_page(self):
        self.log("打开求助详情")
        detail_button = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, SELECTORS["ask_detail"]))
        )
        self.driver.execute_script("arguments[0].click();", detail_button)
        time.sleep(2)

    def get_article_title(self, fallback: str) -> str:
        title = safe_text(self.driver, SELECTORS["article_title"], fallback)
        self.log("文献标题: %s", title)
        return title

    def wait_for_result(self):
        self.log("等待平台返回结果")
        self.js_click_xpath(XPATHS["confirm_ok"], long_wait=True)
        self.log("确认文献")
        time.sleep(1)

    def review_if_needed(self):
        clicked = self.maybe_click_xpath(XPATHS["confirm_ok"])
        if clicked:
            self.log("已执行查看/审核确认")

    def download_file(self) -> Optional[str]:
        self.log("等待下载链接")
        link = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, SELECTORS["download_link"]))
        )
        url = link.get_attribute("href")
        self.log("下载地址: %s", url)
        link.click()
        time.sleep(2)
        return url

    def accept_result(self):
        self.log("执行采纳流程")
        self.click_css(SELECTORS["accept_button"])
        self.click_css(SELECTORS["accept_confirm"])
        self.click_css(SELECTORS["accept_finish"])

    def get_credits(self) -> str:
        return safe_text(self.driver, SELECTORS["credits"], "未知")


# =========================
# 单任务
# =========================
def process_identifier(identifier: str, cookie_text: str, thread_index: int) -> TaskResult:
    driver = None
    thread_name = f"线程-{thread_index}"

    try:
        logger.info("[%s] 开始处理: %s", thread_name, identifier)
        driver = create_driver()
        client = AbleSciClient(driver, thread_name)
        client.inject_cookies(cookie_text)

        client.go_to_query_page()
        client.submit_identifier(identifier)
        client.open_detail_page()

        title = client.get_article_title(identifier)
        client.wait_for_result()
        client.review_if_needed()
        client.download_file()
        client.accept_result()

        logger.info("[%s] 已触发下载: %s", thread_name, title)
        return TaskResult(
            identifier=identifier,
            success=True,
            article_title=title,
            message="下载流程已触发"
        )

    except TimeoutException:
        logger.exception("[%s] 处理超时: %s", thread_name, identifier)
        return TaskResult(
            identifier=identifier,
            success=False,
            message="处理超时"
        )
    except Exception as e:
        logger.exception("[%s] 处理失败: %s", thread_name, identifier)
        return TaskResult(
            identifier=identifier,
            success=False,
            message=str(e)
        )
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
            logger.info("[%s] 浏览器已关闭", thread_name)


# =========================
# 获取积分
# =========================
def fetch_credits(cookie_text: str) -> str:
    driver = None
    try:
        driver = create_driver()
        client = AbleSciClient(driver, "积分查询")
        client.inject_cookies(cookie_text)
        return client.get_credits()
    except Exception:
        logger.exception("获取积分失败")
        return "无法获得剩余积分"
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


# =========================
# 主程序
# =========================
def main():
    try:
        ensure_download_dir()
        cookie_text = load_cookie_text()

        raw = input("请输入 DOI 或 PMID，多个请用分号(;)分隔：\n").strip()
        identifiers = [x.strip() for x in raw.split(";") if x.strip()]
        if not identifiers:
            logger.info("未输入有效 DOI/PMID，程序结束")
            return

        max_threads = min(MAX_THREADS_LIMIT, len(identifiers))
        logger.info("共检测到 %s 个任务，最大并发线程数 %s", len(identifiers), max_threads)

        results: List[TaskResult] = []

        with ThreadPoolExecutor(max_workers=max_threads) as executor:
            futures = {}
            for idx, identifier in enumerate(identifiers, start=1):
                future = executor.submit(process_identifier, identifier, cookie_text, idx)
                futures[future] = identifier
                time.sleep(STARTUP_INTERVAL)

            for future in as_completed(futures):
                result = future.result()
                results.append(result)

        success_list = [r for r in results if r.success]
        failed_list = [r for r in results if not r.success]

        logger.info("等待下载任务稳定")
        wait_for_no_partial_downloads()

        logger.info("最近下载文件:")
        for file_name, ctime in list_recent_files():
            logger.info("- %s | %s", file_name, time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ctime)))

        credits = fetch_credits(cookie_text)

        logger.info("处理结果摘要")
        logger.info("成功: %s/%s", len(success_list), len(results))
        if success_list:
            for item in success_list:
                logger.info("成功: %s (%s)", item.identifier, item.article_title)
        if failed_list:
            for item in failed_list:
                logger.info("失败: %s (%s)", item.identifier, item.message)

        logger.info("当前剩余积分: %s", credits)

    except Exception:
        logger.exception("程序执行过程中发生错误")
    finally:
        input("\n所有操作已完成，按回车键关闭...")


if __name__ == "__main__":
    main()
