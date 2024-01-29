# This is a program that helps you submit several Redcap forms. It's useful when the Redcap
#	website loads very slowly, because you can have the program transfer spreadsheet
#	data into many Redcap submissions; you won't have to wait for the page to slowly load
#	each time, because the program will run and do the waiting for you.
# I made this as a way to learn Selenium. See a guide here: 
#	https://www.scaler.com/topics/selenium-tutorial/form-input-in-selenium/
# Make sure Selenium is installed and up-to-date (use pip, ie with "pip install --upgrade selenium").
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

# Spam a button (identified via id) until it is clickable. Then, keep spamming it until it is 
#	considered no longer clickable. Trust me, both steps are necessary for some of these buttons.
def spam_sandwich(driver,id):
	while 1:
		try:
			driver.find_element("id",id).click() # click the button
			break
		except WebDriverException:
			pass
	while 1:
		try:
    			driver.find_element("id",id).click() # click the button
		except WebDriverException:
			break

# The same as the spam_sandwich() function except it clicks a clickable element (identified via xpath) 
#	that is not necessarily a button. Also, once the button is clicked, the function ends instead of 
#	waiting until it is no longer clickable.
def spam_by_x(driver,xpath):
	while 1:
		try:
			driver.find_element("xpath",xpath).click() # click the button
			break
		except WebDriverException:
			pass
	while 0:
		try:
    			driver.find_element("xpath",xpath).click() # click the button
		except WebDriverException:
			break


import time # For a pause with time.sleep()
import getpass # For a hidden password input through getpass.getpass()

# Import tools that I don't understand:
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService

from selenium.common.exceptions import WebDriverException

# So I can navigate the webpage with tab and space keys:
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains



# Load a headless browser and set the URL of the Redcap survey:
options = webdriver.ChromeOptions() ; options.headless = True ; url = "https://redcap.partners.org/redcap/plugins/survey_token/survey_token_login.php?pid=18168&hash=988968e7-e9d0-4581-9c1e-0ddd3f5b8036"


# I used to use the following two lines, but they only work on some computers (perhaps a problem with Chrome versions?)
#from webdriver_manager.chrome import ChromeDriverManager
#with webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options) as driver:
# Now, instead of the above two lines, I use the below two lines (the "if 1:" is to keep indentation valid). See here for what led me to this solution: https://stackoverflow.com/questions/76553813/messageunable-to-obtain-chromedriver-using-selenium-manager
for count in range(1):
	driver = webdriver.Chrome()
	
	# Load the webpage
	driver.get(url)

	login(driver)

	# Show that it's reached the page properly:
	# print("Page URL:", driver.current_url) 
	# print("\n\n")
		
	# Wait until the text input field loads, then type in the employee's ID number:
	try:
		driver.find_element("id","search_field").send_keys("100634497")
	except WebDriverException:
		time.sleep(1)

	# I think if I take a closer look at the Selenium tutorial, I'll figure out how to select from
	#	the dropdown list and print the options.

	# I can't figure out how to select the option from the drop-down list, so let's click this one
	#	weird "leftdiv" on the page, which isn't a button but is some clickable region I
	#	guess (but the click is designed to do nothing). Anyway, this makes the website \
	# 	respond by automatically selecting the (hopefully) only employee in the dropdown list.
	driver.find_element("id","respirator_fit_testing-left").click()

	# Spam the Select Employee button until it's actually clickable. Unfortunately, Selenium still
	#	thinks it's "clickable" even after the website already registers the click. So, we have
	#	 to click the button until it's clickable, and then eventually becomes non-clickable 			
	# 	again. Only then can we be sure it was registered and we can move on to the next button.
	spam_sandwich(driver,"cpsubmit")
	
	# Do the same for the "Respirator Fit Testing" button:
	spam_sandwich(driver,"respirator_fit_testing-submit")

	# At this point, the website may take 10+ seconds to load the survey but so far it's always 
	#	worked. Now that the survey's up, mine the data that autopopulated in the ID #, 
	#	Last Name, and First Name fields. I can't find these elements by id for 
	# 	some reason (I don't really understand HTML), so I'm finding them by name
	#	(inspect element and search the HTML code for "name=___" on the element to see
	#	what I'm talking about).
	print("\n")
	print("The website autopopulated the Employee ID as:",driver.find_element("name","employeeid").get_attribute("value"),"\n")
	print("The website autopopulated the Last Name as:",driver.find_element("name","last_name").get_attribute("value"),"\n")
	print("The website autopopulated the First Name ID as:",driver.find_element("name","first_name").get_attribute("value"),"\n")
	
	# Type in the employee's ID number (I don't think there's a need to wait for 
	#	loading at all, but I made the program do it just in case.) Delete the autopopulated response first.
	try:
		driver.find_element("name","fittestdate").clear()
		# driver.find_element("name","fittestdate").send_keys(input("Fit Test Date (MM-DD-YYYY): "))	
		driver.find_element("name","fittestdate").send_keys("01-01-2024")	
	except WebDriverException:
		time.sleep(1)

	# actions = ActionChains(driver)

	# The next step is loading up all the xpath variables. These may change over time but are unlikely to:
	Small_xpath = """//*[@id="fitresult-tr"]/td[2]/i/i/span/div/div[1]/label/span"""
	Failed_xpath = """//*[@id="fitresult-tr"]/td[2]/i/i/span/div/div[2]/label/span"""
	MedReg_xpath = """//*[@id="fitresult-tr"]/td[2]/i/i/span/div/div[3]/label/span"""
	Yes_xpath = """//*[@id="fittestpassed-tr"]/td[2]/i/span/div/div[1]/label/span"""
	No_xpath = """//*[@id="fittestpassed-tr"]/td[2]/i/span/div/div[2]/label/span"""
	FacialHair_xpath = """//*[@id="fittestfailreason-tr"]/td[2]/i/span/div/div[1]/label/span"""
	FacialSize_xpath = """//*[@id="fittestfailreason-tr"]/td[2]/i/span/div/div[2]/label/span"""
	Mask1860S_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[3]/label/span"""
	Mask1860_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[3]/label/span"""
	Mask1870Plus_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[9]/label/span"""
	HalyardSmall_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[20]/label/span"""
	HalyardRegular_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[22]/label/span"""
	N95_xpath = """//*[@id="respclass-tr"]/td[2]/i/i/span/div/div[1]/label/span"""
	PAPR_xpath = """//*[@id="respclass-tr"]/td[2]/i/i/span/div/div[2]/label/span"""
	Submit_xpath = """//*[@id="questiontable"]/tbody/tr[22]/td/table/tbody/tr/td/button"""



	# Answer the question for Fit Test Result. Just a note, if you're looking for the xpath of something
	#	like the "Failed" button on the survey; Ctrl+F the Inspect Element and it will be the 2nd of 
	#	the two viable elements in the HTML script. Chrome will highlight the word itself inside the
	#	survey button, not the whole button -- when you select it in the HTML.
	
	# Small
	if 0:	
		spam_by_x(driver,Small_xpath)

	# Failed
	if 0: 
		spam_by_x(driver,Failed_xpath)

	# Med/Reg
	if 1: 
		spam_by_x(driver,MedReg_xpath)

	# Answer: Was the respirator test successful?
	
	# Yes
	if 1: 	
		spam_by_x(driver,Yes_xpath)
	
	# No
	if 0: 
		spam_by_x(driver,No_xpath)

	# Answer (if you just said "no"): What was the failed reason?

	# Facial hair
	if 0: 
		spam_by_x(driver,FacialHair_xpath)
	# Facial Size, Shape, or Structure
	if 0: 
		spam_by_x(driver,FacialSize_xpath)
	
	# Answer: Respirator Model?	

	# 3M 1860S
	if 0: 
		spam_by_x(driver,Mask1860S_xpath)
	# 3M 1860
	if 1: 
		spam_by_x(driver,Mask1860_xpath)
	# 3M 1870+
	if 0: 
		spam_by_x(driver,Mask1870Plus_xpath)

	# Halyard Fluidshield #46827 S
	if 0: 
		spam_by_x(driver,HalyardSmall_xpath)

	# Halyard Fluidshield 46727 R
	if 0: 
		spam_by_x(driver,HalyardRegular_xpath)
	
	# Answer: Will we choose N95 or PAPR? Choose N95 for a Fit Test pass, and choose PAPR if failed all options.
	# N95
	if 1: 
		spam_by_x(driver,N95_xpath)
	# PAPR
	if 0: 
		spam_by_x(driver,PAPR_xpath)

	# Hit submit
	if 0:
		spam_by_x(driver,Submit_xpath)

	# End program
	unused_var = input("Program done. Press ENTER to close.\n")
	print("Closing...\n")



