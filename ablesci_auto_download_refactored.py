import logging
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# =========================
# 配置区
# =========================

BASE_URL = "https://www.ablesci.com/"
DOWNLOAD_DIR = Path.cwd() / "文献下载"
COOKIES_FILE = Path("cookies.txt")
LOG_FILE = Path("ablesci_run.log")

HEADLESS = True          # 后台运行
WAIT_TIMEOUT = 20
LONG_WAIT_TIMEOUT = 28800   # 8小时
DOWNLOAD_START_WAIT = 10
FINAL_EXTRA_WAIT = 10


SELECTORS = {
    "nav_query_page": "body > div.able-header.header-bg-assist > div > div > a",
    "input_identifier": "#onekey",
    "submit_button": "#assist-create-form > div.alert.alert-success > div.layui-form-item.layui-row > div:nth-child(2) > div > button",
    "submit_confirm": "#layui-layer2 > div.layui-layer-btn.layui-layer-btn- > a.layui-layer-btn0",
    "ask_detail": "#layui-layer5 > div.layui-layer-btn.layui-layer-btn- > a",
    "article_title": "#LAY_ucm > div:nth-child(1) > div.assist-detail.layui-row > div > table > tbody > tr:nth-child(1) > td.assist-title > div:nth-child(1)",
    "result_confirm": "#layui-layer1 > div > div:nth-child(3) > div.layui-layer-btn.layui-layer-btn- > a",
    "review_button": "#layui-layer1 > div.layui-layer-btn.layui-layer-btn- > a",
    "download_link": "a[title='点击下载']",
    "accept_button": "#uploaded-file-handle > button:nth-child(1)",
    "accept_confirm": "#layui-layer2 > div.layui-layer-btn.layui-layer-btn- > a.layui-layer-btn0",
    "accept_finish": "#layui-layer4 > div.layui-layer-btn.layui-layer-btn- > a",
    "credits": "#user-point-now",
}


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
# 日志
# =========================

def setup_logger() -> logging.Logger:
    logger = logging.getLogger("ablesci")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s"
    )

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
# 基础工具
# =========================

def ensure_download_dir():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("下载目录: %s", DOWNLOAD_DIR)


def load_cookies_text() -> str:
    if not COOKIES_FILE.exists():
        raise FileNotFoundError(f"未找到 cookies 文件: {COOKIES_FILE}")
    content = COOKIES_FILE.read_text(encoding="utf-8").strip()
    if not content:
        raise ValueError("cookies.txt 为空")
    return content


def parse_cookie_string(cookie_text: str) -> List[dict]:
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
        logger.info("尝试自动启动 ChromeDriver")
        return webdriver.Chrome(options=options)
    except Exception as e:
        logger.warning("自动启动失败: %s", e)

    app_path = get_application_path()
    possible_paths = [
        app_path / "chromedriver.exe",
        Path.cwd() / "chromedriver.exe",
        Path("chromedriver.exe"),
    ]

    for path in possible_paths:
        if path.exists():
            logger.info("使用本地 ChromeDriver: %s", path)
            service = Service(executable_path=str(path))
            return webdriver.Chrome(service=service, options=options)

    raise FileNotFoundError("ChromeDriver 未找到")


def wait_click(wait: WebDriverWait, selector: str):
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
    element.click()
    return element


def wait_visible(wait: WebDriverWait, selector: str):
    return wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selector)))


def safe_text(driver: WebDriver, selector: str, default: str = "") -> str:
    try:
        return driver.find_element(By.CSS_SELECTOR, selector).text.strip()
    except Exception:
        return default


def list_recent_files(limit: int = 10):
    files = []
    for file in DOWNLOAD_DIR.iterdir():
        if file.is_file():
            files.append((file.name, file.stat().st_ctime))
    files.sort(key=lambda x: x[1], reverse=True)
    return files[:limit]


def wait_for_download_settle(timeout: int = 120):
    """
    简单等待下载完成：
    若目录里还有 .crdownload，则继续等。
    """
    start = time.time()
    while time.time() - start < timeout:
        partials = list(DOWNLOAD_DIR.glob("*.crdownload"))
        if not partials:
            return True
        time.sleep(2)
    return False


# =========================
# 业务流程
# =========================

class AbleSciClient:
    def __init__(self, driver: WebDriver):
        self.driver = driver
        self.wait = WebDriverWait(driver, WAIT_TIMEOUT)
        self.long_wait = WebDriverWait(driver, LONG_WAIT_TIMEOUT)

    def open_home(self):
        logger.info("打开首页")
        self.driver.get(BASE_URL)

    def inject_cookies(self, cookie_text: str):
        self.open_home()
        time.sleep(1)

        cookies = parse_cookie_string(cookie_text)
        if not cookies:
            raise ValueError("未解析出有效 cookies")

        added = 0
        for cookie in cookies:
            try:
                self.driver.add_cookie(cookie)
                added += 1
            except Exception as e:
                logger.warning("添加 cookie 失败: %s, %s", cookie.get("name"), e)

        logger.info("成功添加 cookie 数量: %s", added)
        self.driver.refresh()
        time.sleep(2)

    def go_to_query_page(self):
        logger.info("进入查询页")
        wait_click(self.wait, SELECTORS["nav_query_page"])

    def submit_identifier(self, identifier: str):
        logger.info("提交标识符: %s", identifier)
        input_box = wait_visible(self.wait, SELECTORS["input_identifier"])
        input_box.clear()
        input_box.send_keys(identifier)

        wait_click(self.wait, SELECTORS["submit_button"])
        wait_click(self.wait, SELECTORS["submit_confirm"])

    def open_detail_page(self):
        logger.info("打开求助详情页")
        wait_click(self.wait, SELECTORS["ask_detail"])

    def wait_for_result(self):
        logger.info("等待平台返回结果")
        wait_click(self.long_wait, SELECTORS["result_confirm"])

    def confirm_and_review(self):
        logger.info("确认并查看审核页")
        wait_click(self.wait, SELECTORS["review_button"])

    def get_article_title(self, fallback: str) -> str:
        title = safe_text(self.driver, SELECTORS["article_title"], fallback)
        logger.info("文献标题: %s", title)
        return title

    def download_file(self) -> Optional[str]:
        logger.info("等待下载链接")
        element = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, SELECTORS["download_link"]))
        )
        url = element.get_attribute("href")
        logger.info("下载链接: %s", url)
        element.click()
        time.sleep(DOWNLOAD_START_WAIT)
        return url

    def accept_result(self):
        logger.info("执行采纳流程")
        wait_click(self.wait, SELECTORS["accept_button"])
        wait_click(self.wait, SELECTORS["accept_confirm"])
        wait_click(self.wait, SELECTORS["accept_finish"])

    def recover_home(self):
        logger.info("尝试恢复到首页")
        try:
            self.driver.get(BASE_URL)
            time.sleep(2)
        except Exception as e:
            logger.warning("恢复首页失败: %s", e)

    def get_credits(self) -> str:
        return safe_text(self.driver, SELECTORS["credits"], "未知")


def process_one(client: AbleSciClient, identifier: str) -> TaskResult:
    try:
        logger.info("开始处理: %s", identifier)

        client.go_to_query_page()
        client.submit_identifier(identifier)
        client.open_detail_page()

        title = client.get_article_title(identifier)
        client.wait_for_result()
        client.confirm_and_review()
        client.download_file()
        client.accept_result()

        logger.info("已触发下载: %s", title)
        return TaskResult(
            identifier=identifier,
            success=True,
            article_title=title,
            message="下载流程已触发"
        )

    except TimeoutException:
        logger.exception("处理超时: %s", identifier)
        client.recover_home()
        return TaskResult(
            identifier=identifier,
            success=False,
            message="处理超时"
        )
    except Exception as e:
        logger.exception("处理失败: %s", identifier)
        client.recover_home()
        return TaskResult(
            identifier=identifier,
            success=False,
            message=str(e)
        )


# =========================
# 主函数
# =========================

def main():
    driver = None
    try:
        ensure_download_dir()
        cookie_text = load_cookies_text()

        logger.info("启动浏览器")
        driver = create_driver()
        client = AbleSciClient(driver)

        client.inject_cookies(cookie_text)

        raw = input("请输入 DOI 或 PMID，多个请用分号(;)分隔：\n").strip()
        identifiers = [x.strip() for x in raw.split(";") if x.strip()]
        if not identifiers:
            logger.info("未输入有效 DOI/PMID，程序结束")
            return

        logger.info("任务总数: %s", len(identifiers))

        results: List[TaskResult] = []

        for index, identifier in enumerate(identifiers, start=1):
            logger.info("处理进度: %s/%s", index, len(identifiers))
            result = process_one(client, identifier)
            results.append(result)

        success_list = [r for r in results if r.success]
        failed_list = [r for r in results if not r.success]

        if success_list:
            logger.info("等待下载稳定")
            wait_for_download_settle(timeout=120)
            time.sleep(FINAL_EXTRA_WAIT)

            logger.info("最近下载文件:")
            for name, ctime in list_recent_files():
                logger.info("- %s | %s", name, time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ctime)))

        logger.info("处理结果摘要")
        logger.info("成功: %s/%s", len(success_list), len(results))

        if success_list:
            logger.info("成功列表: %s", ", ".join(r.identifier for r in success_list))
        if failed_list:
            logger.info("失败列表: %s", ", ".join(f"{r.identifier}({r.message})" for r in failed_list))

        credits = client.get_credits()
        logger.info("剩余积分: %s", credits)

    except Exception:
        logger.exception("程序执行过程中发生错误")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        input("\n所有操作已完成，按回车键关闭...")


if __name__ == "__main__":
    main()
