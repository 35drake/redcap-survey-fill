# This is a program that helps you submit several Redcap forms. It's useful when the Redcap
#	website loads very slowly, because you can have the program transfer spreadsheet
#	data into many Redcap submissions; you won't have to wait for the page to slowly load
#	each time, because the program will run and do the waiting for you.
# I made this as a way to learn Selenium. See a guide here: 
#	https://www.scaler.com/topics/selenium-tutorial/form-input-in-selenium/
# Make sure Selenium is installed and up-to-date (use pip).
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

# Spam a button until it is clickable. Then, keep spamming it until it is considered no longer
#	clickable. Trust me, both steps are necessary for some of these buttons.
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

# This function navigates the webpage and allows you to press buttons that Selenium can't normally press.
#	This is accomplished by TAB and SPACE presses. It will press TAB a certain number of times, then press
#	SPACE to hit the button. Note! the function will then SHIFT+TAB to reset you back to your original
#	position. If you want your ending position to actually move across the page, then use tab_move().
def hit_button_with_space(actions,tabs):
	tabs = int(tabs)
	for count in range(0,tabs):
		actions = actions.send_keys(Keys.TAB)
	actions = actions.send_keys(Keys.SPACE)
	actions.perform()
	time.sleep(1)
	for count in range(0,tabs):
		actions = actions.key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT)
	actions.perform() # This performs the actions and also clears them out from the actions stack
	time.sleep(1)

def tab_move(actions,tabs):
	tabs = int(tabs)
	if tabs < 0: #It should only be < 0 for debug purposes
		for count in range(0,-tabs):
			actions = actions.key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT)
			actions.perform()
	else:	
		for count in range(0,tabs):
			actions = actions.send_keys(Keys.TAB)	
			actions.perform()


import time # For a pause with time.sleep()
import getpass # For a hidden password input through getpass.getpass()

# Import tools that I don't understand:
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException

# So I can navigate the webpage with tab and space keys:
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains



# Load a headless browser and set the URL of the Redcap survey:
options = webdriver.ChromeOptions() ; options.headless = True ; url = "https://redcap.partners.org/redcap/plugins/survey_token/survey_token_login.php?pid=18168&hash=988968e7-e9d0-4581-9c1e-0ddd3f5b8036"




with webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options) as driver:
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
	#	 to click the button until it's clickable, and then eventually becomes non-clickable 			# 	again. Only then can we be sure it was registered and we can move on to the next button.
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
		driver.find_element("name","fittestdate").send_keys(input("Fit Test Date (MM-DD-YYYY): "))	
	except WebDriverException:
		time.sleep(1)


	actions = ActionChains(driver)

	# Move to the Fit Test Result question:
	tab_move(actions,2)

	# Make your answer 
	hit_button_with_space(actions,input( "What size? (0 = Small, 1 = Failed, 2 = Med/Red)  " ))

	# Move to the next question:
	tab_move(actions,4)

	# Make your answer 
	hit_button_with_space(actions,input(  "Was the test successful? (0 = Yes, 1 = No)  " )  )

	# Move to the next question:
	tab_move(actions,3)

	# Note that here the survey changes depending on your answer. If you click "No", an extra
	#	question pops up. However the survey seems broken as it doesn't ever pop up that 
	#	question if you navigate with tab+space. Oh well, I'm just gonna ignore it for now.

	# Answer Respirator Model:
	hit_button_with_space(actions,15) # 0 to 23, depending on model. 15 should get you to "Envo"

	# Move to the next question:
	tab_move(actions,2)

	# Answer Respirator Class:
	hit_button_with_space(actions,input("Respirator Class (0 for N95, 1 for PAPR): "))
 
#Actually, WAIT! The Respirator Class question doesn't even come up with tab+space; you need an actual click.
# I cannot use the tab+space strategy then, and will instead now use xpath to find elements that I couldn't
# locate before (because they had no obvious HTML id). And I'll click them, which is possible even though I
# didn't realize it before.



	# Hit submit

	# End program
	unused_var = input("Program done. Press ENTER to close.\n")
	print("Closing...\n")


#driver.find_element("id","opt-model_MM_2200")
# ac = ActionChains(driver)
# ac.move_to_element(elem).move_by_offset(x_offset, y_offset).click().perform()


