from selenium import webdriver

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('useAutomationExtension', False)  # 防止检测
chrome_options.add_argument("--mute-audio")  # 静音
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])  # 防止检测、禁止打印日志
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument('--ignore-certificate-errors')  # 忽略证书错误
chrome_options.add_argument('--ignore-ssl-errors')  # 忽略ssl错误
chrome_options.add_argument('–log-level=3')
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument('--remote-debugging-port=9222')
driver = webdriver.Chrome(executable_path='C:/Users/Administrator/AppData/Local/Google/Chrome/Application/chromedriver.exe',options=chrome_options)
driver.get('https://wzgzt.jygl.sinopec.com/')