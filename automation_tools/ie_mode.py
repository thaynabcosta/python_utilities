from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
from botcity.web.browsers.ie import default_options

EDGE_PATH = r'resources\msedgedriver.exe'

def configure_edge():
    edge_options = default_options()
    edge_options.add_additional_option("ie.edgechromium", True)  # Compatibilidade com IE
    edge_options.add_additional_option("ignoreProtectedModeSettings", True)
    edge_options.add_additional_option("ie.edgepath", EDGE_PATH)  # Caminho do Edge
    return edge_options


desktop_bot = DesktopBot()
webbot = WebBot()
    
webbot.headless = False
webbot.browser = Browser.EDGE
webbot.driver_path = EDGE_PATH 
webbot.options = configure_edge()

webbot.start_browser()
 
webbot.browse("https://exemplo.com/inicial")

webbot.navigate_to("https://exemplo.com/sistema_legado")
    
webbot.stop_browser()
