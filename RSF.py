# This is a program that helps you submit several Redcap forms. It's useful when the Redcap
#	website loads very slowly, because you can have the program transfer spreadsheet
#	data into many Redcap submissions; you won't have to wait for the page to slowly load
#	each time, because the program will run and do the waiting for you.
# I made this as a way to learn Selenium. See a guide here: 
#	https://www.scaler.com/topics/selenium-tutorial/form-input-in-selenium/
# The most important function here is driver.find_element(attribute,name). It's how you locate
#	HTML elements for the program to work with. Use Inspect Element on a webpage to find
#	the attribute of your element (i.e. "id" or "class" or "style") and the actual 
#	name/value of its attribute (i.e. an element's id could be "username" if it's the
#	username text entry box on the website).
# Debugging reminder: To debug a program in the python terminal, execute the program with
#	"python -m pdb RSF.py". You'll type in "n" every time you want the next line to be
#	executed, and in the middle of program execution you can send commands like a new line
#	of program code. Type "exit" to prematurely exit. Also, note that Python has trouble
#	handling single-letter variable names, so during debug don't send a command like 
#	"a = driver.find_element("id","blah").

# This will do the login portion of the program
def login(driver):
	if 0:
		# Log in by taking credentials, then pressing the "LOG IN" button
		driver.find_element("id","username").send_keys(input("Username: "))
		driver.find_element("id","password").send_keys(getpass.getpass("Password: "))
	else:
		driver.find_element("id","username").send_keys("ds001")
		driver.find_element("id","password").send_keys("fake")
	login_button = driver.find_element("id","login_btn")
	login_button.click()  # Select the button by clicking on it

	# You should be logged in now; show the title of the webpage you've gotten to now
	# print("Page Title:", driver.title)
	# print("\n\n")

# Keep clicking a button until the click works (use this for a button that's slow to load)
def spam(driver,id):
	while 1:
		try:
    			driver.find_element("id",id).click() # click the button
		except WebDriverException:
			print("The button is still not clickable yet. Waiting.\n")
			time.sleep(1) # Wait another second, since it isn't clickable yet
    		


import time # For a pause with time.sleep()
import getpass # For a hidden password input through getpass.getpass()

# Import tools that I don't understand:
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException


# Load a headless browser and set the URL of the Redcap survey:
options = webdriver.ChromeOptions() ; options.headless = True ; url = "https://redcap.partners.org/redcap/plugins/survey_token/survey_token_login.php?pid=18168&hash=988968e7-e9d0-4581-9c1e-0ddd3f5b8036"




with webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options) as driver:
	# Load the webpage
	driver.get(url)

	login(driver)

	# Show that it's reached the page properly:
	# print("Page URL:", driver.current_url) 
	# print("\n\n")
		
	# Type in the employee's ID number:
	# driver.find_element("id","search_field").send_keys(input("Employee ID: "))
	driver.find_element("id","search_field").send_keys("100634497")

	# I can't figure out how to select the option from the drop-down list, so let's click this one
	#	weird "leftdiv" on the page, which isn't a button but is some clickable region I
	#	guess (but the click is designed to do nothing). Anyway, this makes the website \
	# 	respond by automatically selecting the (hopefully) only employee in the dropdown list.
	driver.find_element("id","respirator_fit_testing-left").click()

	spam(driver,"cpsubmit") #spam the Select Employee button until it's actually clickable



	# End program
	unused_var = input("Program done. Press ENTER to close.\n")




