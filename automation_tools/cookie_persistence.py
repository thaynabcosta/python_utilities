from botcity.web import WebBot, Browser

def export_cookies(bot):
    return bot.browser.get_cookies()

def import_cookies(bot, cookies):
    for cookie in cookies:

        cookie.pop('sameSite', None)
        bot.browser.add_cookie(cookie)

bot_edge = WebBot()
bot_edge.browser = Browser.EDGE
bot_edge.start_browser()
bot_edge.browse("https://www.seu-site.com/")
# [...]
cookies = export_cookies(bot_edge)
bot_edge.wait(5000)

bot_ie = WebBot()
bot_ie.browser = Browser.IE
bot_ie.start_browser()
bot_ie.browse("https://www.seu-site.com/")

# Importar os cookies
import_cookies(bot_ie, cookies)
bot_ie.refresh()

bot_edge.wait(5000)
bot_ie.wait(5000)

bot_edge.stop_browser()
bot_ie.stop_browser()
