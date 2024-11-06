from typing import Literal
from botcity.web import Browser, WebBot
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager, IEDriverManager
from webdriver_manager.firefox import GeckoDriverManager
import time

def choosing_browser(browser: Literal["edge", "firefox", "chrome", "ie"] = "edge"):
    """
    Obtém a versão compatível do driver para o navegador especificado.

    :param browser: Navegador ('edge', 'chrome', 'firefox', ou 'ie')
    :return: O tipo de navegador e o caminho para o driver instalado.
    """
    match browser.lower():
        case "edge":
            return Browser.EDGE, EdgeChromiumDriverManager().install()
        case "chrome":
            return Browser.CHROME, ChromeDriverManager().install()
        case "firefox":
            return Browser.FIREFOX, GeckoDriverManager().install()
        case "ie":
            return Browser.INTERNET_EXPLORER, IEDriverManager().install()
        case _:
            raise TypeError(
                "Expected type 'Literal['edge','firefox','chrome','ie']',",
                f"got 'Literal['{browser}']' instead",
            )

def open_system_in_ie_module(webbot, url: str, browser_choice="edge"):

    try:
        webbot.browser, webbot.driver_path = choosing_browser(browser_choice)
        webbot.browse(url)

        time.sleep(3)

    except Exception as e:
        webbot.browser, webbot.driver_path = choosing_browser("ie")
        webbot.browse(url)
        time.sleep(3)

if __name__ == "__main__":
    webbot = WebBot()

    # Configure whether or not to run on headless mode
    webbot.headless = False

    system_url = "https://www.botcity.dev"

    open_system_in_ie_module(webbot, system_url, browser_choice="chrome")

    # Wait 3 seconds before closing
    webbot.wait(3000)

    webbot.stop_browser()
