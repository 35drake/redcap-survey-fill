# Import tools that I don't understand:
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# The URL of the Redcap survey:
url = "https://redcap.partners.org/redcap/plugins/survey_token/survey_token_login.php?pid=18168&hash=988968e7-e9d0-4581-9c1e-0ddd3f5b8036"

# Load a headless browser
options = webdriver.ChromeOptions()
options.headless = True
with webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options) as driver:
	driver.get(url)
	# Show that it's working:
	print("Page URL:", driver.current_url) 
	print("Page Title:", driver.title) #Redcap site has no title...






