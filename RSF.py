# Drake Sorkhab, February 2024

submitting_for_real = 1 # Set to 0 if you're debugging or 1 if you're running

# SUMMARY AND NOTES
# This is a program that helps you submit several Redcap forms. It's useful when the Redcap
#	website loads very slowly, because you can have the program transfer spreadsheet
#	data into many Redcap submissions; you won't have to wait for the page to slowly load
#	each time, because the program will run and do the waiting for you.
# Note that this program is currently set up to submit a specific RedCap form; you'll have to 
#	alter the source code and sample spreadsheet to make it work with other forms.
#	If RedCap updates or changes the format of its login page or survey pages, even more 
#	of the program will have to be changed.
# I learned the basics of Selenium while creating this program. See a guide here: 
#	https://www.scaler.com/topics/selenium-tutorial/form-input-in-selenium/
# Make sure Selenium is installed and up-to-date (use pip, ie with "pip install --upgrade selenium").
#	Several other libraries are needed too (see the imported libraries in the next section).
# The most important Selenium function here is driver.find_element(attribute,name). It's how you locate
#	HTML elements for the program to work with. Use Inspect Element on a webpage to find
#	the attribute of your element (i.e. "id" or "xpath") and the actual 
#	name/value of its attribute (i.e. an element's id could be "username" if it's the
#	username text entry box on the website). I started by using id for all survey buttons
#	and interactable objects, but some buttons were programmed into RedCap strangely
#	and without an id, so I had to find them by xpath.
# Debugging reminder: To debug a program in the python terminal, execute the program with
#	"python -m pdb RSF.py". You'll type in "n" every time you want the next line to be
#	executed, and in the middle of program execution you can send commands like a new line
#	of program code. Type "exit" to prematurely exit. Also, note that Python has trouble
#	handling single-letter variable names, so during debug don't send a command like 
#	"a = driver.find_element("id","blah").

# BUG NOTE
# The program doesn't consistently write to the Excel spreadsheet (to signify that a result has been successfully uploaded)
#	unless your computer has no other Excel files open. It's possible that the Excel spreadsheet must only be open
#	on this computer as well and no others (i.e. you must be careful with OneDrive files).
# I think the real issue is that you must be SIGNED IN to MS Office, which isn't always obvious (it can halfway "sign you out"
#	if you sign in on another laptop). And you must have auto-save clicked on (do this by opening any Excel sheet, not
#	while the program's running, and set the auto-save slide to "ON").

# PROGRAM OUTLINE
# See the sample spreadsheet to understand the goals of this program. For every "x" that an employee has in a cell,
#	a RedCap form must be submitted. Note that column "L" of the spreadsheet contains data, written by the program,
#	that tells you if this employee's data has already been submitted or attempted to be submitted. The program 
#	will only attempt to enter new data that hasn't been done before. 
# Using auto_browser(), the necessary forms are submitted. Each time a form is submitted, column "L" for that employee's 
#	row will be updated as proof of success. The "L" cell will turn green when all forms are done for that person.
#	If there's a problem and all forms don't get in successfully, that employee's "L" cell will stay yellow.


import os # to install libraries with pip properly

try:
	import pandas # for reading Excel data
except:
	os.system("pip install pandas")
	import pandas

try:
	import xlwings # for writing to Excel sheet (change one cell)
except:
	os.system("pip install xlwings")
	import xlwings

try:
	from selenium import webdriver # to launch an automated web browser
except:
	os.system("pip3 install -U selenium") # As of March 2024, this doesn't work with normal "pip"; only "pip3" works
	from selenium import webdriver 

from selenium.webdriver.chrome.service import Service as ChromeService # so that Chrome specifically can be used as the browser

try:
	import openpyxl #This sublibrary is needed to run the program. I believe that pandas is using it.
except:
	os.system("pip install openpyxl")
	import openpyxl


import datetime # For converting Excel dates into strings and getting today's date
import time # For checking time elapsed with time.time() and a possible a pause with time.sleep()
import getpass # For a hidden password input through getpass.getpass()

# Basically this is used to tell us when HTML buttons aren't clickable
from selenium.common.exceptions import WebDriverException

# Seems like I don't need this because pandas will have numpy type support builtin
#	import numpy

# This takes the indices for a matrix (zero-based), and converts it to the corresponding Excel cell index (ie 3,8 --> "I4")
def indices_to_cell(row,col):
	return(chr(65+col) + str(row+1))

# If an Excel file is open (GUI), then pandas may not be able to read from it. This function makes sure it's closed.
def close_Excel(Excel_file_path):
	my_wb = xlwings.Book(Excel_file_path)
	my_wb.close()

# This will do the login portion of the program
def login(driver,username,password):
	driver.find_element("id","username").send_keys(username)
	driver.find_element("id","password").send_keys(password)

	# Press the login button
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

# This function opens up a Chrome tab and attempts to submit a survey with the parameters as the survey options
# VARIABLES EXPLAINED: 
#	"mask_or_failure_number" can be equal to 4,5,6,7,8. Refers to the 1860S, 1860R, 1870+, Halyard S, Halyard R. Can be also equal to 9 or 10. Refers to Fail-By-Face-Shape and Fail-By-Facial-Hair. 
# 	"submitting_for_real" is set to 0 if you're just debugging and don't wanna actually submit RedCap forms. Set it to 1 if you want the forms to actually submit.
#	Make sure the numeric arguments are vals and not strings.
#	The function auto_browser will always return one of four strings: "NOT SUBMITTED YET", "SUBMISSION FAILED", "SUBMISSION WORKED", "SUBMISSION WORKED (PRETEND)".
#	The first string is the default, the second is a failure to load or submit the form, the third is a real successful submission, and the fourth reveals that the form loaded correctly and was almost sent,
#	but the function was called with "submitting_for_real" equal to 0, so the submit button wasn't actually pressed at the end.
def auto_browser(username, password, employee_id, fit_test_date, mask_or_failure_number, submitting_for_real):
	my_return_value = "NOT SUBMITTED YET" # This will be returned at the end of the auto_browser() function to tell you if the form submitted or not
	
	# Make sure that the employee ID number is a 9-digit number as expected
	employee_id_as_string = str(employee_id)
	if len(employee_id_as_string)!=9:
		print("Error, employee ID is not 9 characters long.",employee_id_as_string,"\n")
		print(len(employee_id_as_string))
		return(my_return_value)
	if employee_id_as_string.isdigit == False:
		print("Error, employee ID is not all digits.",employee_id_as_string,"\n")
		for a_char in employee_id_as_string:
			print(a_char)
			print("Order:",ord(a_char))
		return(my_return_value)
	
	# Make the browser headless and set the URL of the Redcap survey:
	options = webdriver.ChromeOptions() ; options.headless = True ; url = "https://redcap.partners.org/redcap/plugins/survey_token/survey_token_login.php?pid=18168&hash=988968e7-e9d0-4581-9c1e-0ddd3f5b8036"

	# Initialise the driver used by Selenium
	driver = webdriver.Chrome()
	
	# Load the webpage
	driver.get(url)

	# Run the login function (defined in this program)
	login(driver, username, password)

	# Print the homepage title and make sure you've reached the homepage
	time_of_login_attempt = time.time()
	while 1:
		print(driver.title)
		print("\n")
		if (driver.title == "REDCap"):
			break
		if time.time() - time_of_login_attempt > 7:		
			print("\nLogin has not succeeded after 7 seconds. Maybe you have the incorrect password?\n\n")
			driver.close()
			return(my_return_value)

	# Wait until the text input field loads, then type in the employee's ID number:
	try:
		# driver.find_element("id","search_field").send_keys("100634497")
		driver.find_element("id","search_field").send_keys(employee_id_as_string)
	except WebDriverException:
		time.sleep(1)

	# I think if I take a closer look at the Selenium tutorial, I'll figure out how to select from
	#	the dropdown list and print the options.

	# I can't figure out how to select the option from the drop-down list, so let's click this one
	#	weird "leftdiv" on the page, which isn't a button but is some clickable region I
	#	guess (but the click is designed to do nothing). Anyway, this makes the website
	# 	respond by automatically selecting the (hopefully) only employee in the dropdown list.
	# NOTE: actually, using "respirator_fit_testing-left" only works on my desktop but throws an
	#	error on my laptop for some reason, so let's click on the text box ("Enter Employee") 
	#	to the left of the text entry box as a backup.
	try:
		driver.find_element("id","respirator_fit_testing-left").click()
	except:
		driver.find_element("xpath",r"""//*[@id="inputfield"]/label""").click()

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
	
	# Type in the fit test date (I don't think there's a need to wait for 
	#	loading at all, but I made the program do it just in case.) Delete the autopopulated response first.
	try:
		driver.find_element("name","fittestdate").clear()
		# driver.find_element("name","fittestdate").send_keys(input("Fit Test Date (MM-DD-YYYY): "))
		# driver.find_element("name","fittestdate").send_keys("01-29-2024")	
		driver.find_element("name","fittestdate").send_keys(fit_test_date.replace("/","-")) #make sure to replace Excel's native date slashes with date hyphens	
	except WebDriverException:
		time.sleep(1)

	# The next step is loading up all the xpath variables. These may change over time but are unlikely to:
	Small_xpath = """//*[@id="fitresult-tr"]/td[2]/i/i/span/div/div[1]/label/span"""
	Failed_xpath = """//*[@id="fitresult-tr"]/td[2]/i/i/span/div/div[2]/label/span"""
	MedReg_xpath = """//*[@id="fitresult-tr"]/td[2]/i/i/span/div/div[3]/label/span"""
	Yes_xpath = """//*[@id="fittestpassed-tr"]/td[2]/i/span/div/div[1]/label/span"""
	No_xpath = """//*[@id="fittestpassed-tr"]/td[2]/i/span/div/div[2]/label/span"""
	FacialHair_xpath = """//*[@id="fittestfailreason-tr"]/td[2]/i/span/div/div[1]/label/span"""
	FacialSize_xpath = """//*[@id="fittestfailreason-tr"]/td[2]/i/span/div/div[2]/label/span"""
	Mask1860S_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[3]/label/span"""
	Mask1860R_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[5]/label/span"""
	Mask1870Plus_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[9]/label/span"""
	HalyardSmall_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[20]/label/span"""
	HalyardRegular_xpath = """//*[@id="model-tr"]/td[2]/i/i/span/div/div[22]/label/span"""
	N95_xpath = """//*[@id="respclass-tr"]/td[2]/i/i/span/div/div[1]/label/span"""
	PAPR_xpath = """//*[@id="respclass-tr"]/td[2]/i/i/span/div/div[2]/label/span"""
	Submit_xpath = """//*[@id="questiontable"]/tbody/tr[22]/td/table/tbody/tr/td/button"""


	# Click the correct boxes depending on what fit test result is being input right now:

	# A passed 1860S Fit test:
	if mask_or_failure_number == 4:
		spam_by_x(driver,Small_xpath)
		spam_by_x(driver,Yes_xpath)
		spam_by_x(driver,Mask1860S_xpath)
		spam_by_x(driver,N95_xpath)
	
	# A passed 1860R Fit test:
	elif mask_or_failure_number == 5:	
		spam_by_x(driver,MedReg_xpath)
		spam_by_x(driver,Yes_xpath)
		spam_by_x(driver,Mask1860R_xpath)
		spam_by_x(driver,N95_xpath)

	# A passed 1870+ Fit test:
	elif mask_or_failure_number == 6:	
		spam_by_x(driver,MedReg_xpath)
		spam_by_x(driver,Yes_xpath)
		spam_by_x(driver,Mask1870Plus_xpath)
		spam_by_x(driver,N95_xpath)

	# A passed Halyard Small Fit test:
	elif mask_or_failure_number == 7:	
		spam_by_x(driver,Small_xpath)
		spam_by_x(driver,Yes_xpath)
		spam_by_x(driver,HalyardSmall_xpath)
		spam_by_x(driver,N95_xpath)

	# A passed Halyard Regular Fit test:
	elif mask_or_failure_number == 8:	
		spam_by_x(driver,MedReg_xpath)
		spam_by_x(driver,Yes_xpath)
		spam_by_x(driver,HalyardRegular_xpath)
		spam_by_x(driver,N95_xpath)		

	# A completely failed Fit Test due to facial size or shape ("Failed N95 options")
	elif mask_or_failure_number == 9:
		spam_by_x(driver,Failed_xpath)
		spam_by_x(driver,No_xpath)		
		spam_by_x(driver,FacialSize_xpath)
		spam_by_x(driver,PAPR_xpath)

	# A completely failed Fit Test due to facial hair
	elif mask_or_failure_number == 10:
		spam_by_x(driver,Failed_xpath)
		spam_by_x(driver,No_xpath)		
		spam_by_x(driver,FacialHair_xpath)
		spam_by_x(driver,PAPR_xpath)
	else:
		print("The function auto_browser() did not get a valid option passed as argument mask_or_failure_number.\n")
		driver.close()
		return(my_return_value)

	# Wait a few seconds before submitting, just in case the user wants to quickly see what the program's doing on the survey	
	time.sleep(5)

	# Submit the completed form (make sure this is OFF if you're bug testing)
	if submitting_for_real: # set this to 1 if you're running the program for real
		spam_by_x(driver,Submit_xpath)
		# Confirm that the form was submitted by checking that you're back to the homepage now
		time_of_submission = time.time()
		while 1:
			if driver.title == "REDCap":
				print("Form submittal confirmed (for real).\n")
				my_return_value = "SUBMISSION WORKED"
				break
			# Check if 45 seconds have passed since submittal but the homepage still hasn't loaded
			if time.time() - time_of_submission > 45:
				my_return_value = "SUBMISSION FAILED"
				print("Error: cannot confirm that form was submitted") 
	else:
		print("Form submittal confirmed (pretend).\n")
		my_return_value = "SUBMISSION WORKED (PRETEND)"

	# End program
	driver.close()
	# unused_var = input("Program done.\n")
	print("Closing...\n")
	return(my_return_value)

# Make a list of the settings from the settings.txt file
settings_file = open("files/settings.txt","r")
settings = []
for line in settings_file:
	if line[-1] == "\n": # the lines (but not necessarily the final line) will end in an unneeded newline
		line = line[0:-1]
	settings = settings + [line]

# Take only the data items and not the field names
settings[0] = settings[0][17:] # "Spreadsheet_path:" is a 17 character-long phrase
settings[1] = settings[1][9:] # This is for Username
settings[2] = settings[2][55:] # This is for Password

# Leave out the space at the start of the item if one exists
for count in range(0,len(settings)):
	if settings[count] != "": # Check if it's empty first, so you don't accessing the first (number 0) character of an empty string
		if settings[count][0] == " ":
			settings[count] = settings[count][1:]
SS_path = settings[0]
my_username = settings[1]
my_password = settings[2]
#print(my_password)
#for my_char in my_password:
#	print(ord(my_char))

# Ask for the password if it's not in the settings (left blank in the settings.txt file)
if my_password == "":
	my_password = getpass.getpass("\nYour password is not stored in your settings. Please type in your password now (it will not be stored): ")

# If the Excel file is already open (and it's stored in OneDrive), pandas will fail to open it. This will close the Excel file if it is currently open.
close_Excel(SS_path)

# Move the contents of the spreadsheet into the my_data NumPy array
try:
	my_data = pandas.read_excel(SS_path, header = None) # "No header" means that the first row is still used. We don't actually need the first row of the spreadsheet, but counting it makes things less confusing for me
except OSError:
	exit("\nFile not found. It may be unsynced with OneDrive right now due to someone editing it. Make sure the file has a green check mark in the File Explorer instead of the two blue arrows.\n")
my_data = my_data.to_numpy()
SS_rows = my_data.shape[0]
SS_cols = my_data.shape[1]

# Convert all the dates into strings
for count in range(5,SS_rows):
	date_time_example = datetime.datetime(2012, 2, 23, 0, 0) # This is an example variable, which we know is of the type "datetime". We'll use it to compare with other variables to see if those are also variables of the "datetime" type.
	if type(my_data[count][0]) == type(date_time_example): # Check if the current item (from col 0) is a date object
		my_data[count][0] = my_data[count][0].strftime("%m/%d/%Y") # Change it into a string

# Check that the NumPy array is working. NOTE: the array naturally uses zero-based numbering which I haven't adjusted. For example, cell 7B's (or colloquially "B7"'s ) value on the spreadsheet can be found at my_data[6,1].
#print("\nSpreadsheet has",SS_cols,"cols and",SS_rows,"rows.\n")
#print(my_data)
#print(my_data[eval(input("Get the cell at row: "))][eval(input("and column: "))],"\n")

# Print relevant spreadsheet contents (the actual employee data)
# for count in range(5,SS_rows):
#	print(my_data[count])

# COLUMNS GUIDE FOR my_data:
# Col 0 is Date 			(program will READ this)
# Col 2 is ID number 			(program will READ this)
# Col 4-8 are Masks 			(program will READ this)
# Col 9-10 are Failure 			(program will READ this)
# Col 11 is RSF Completion Info 	(program will READ AND WRITE this)
# Col 12 is RSF Last Time Accessed 	(program will WRITE this)

# Drake, now for each row of my_data[5:SS_rows], use the item of my_data[5:SS_rows][11] to determine if it needs operation,
#	and if so use the items of my_data[5:SS_rows][0:11] to send the survey with auto_browser(); 
#	then write into items my_data[5:SS_rows][11:13] detailling what happened.


# Go down column 11 and if any items are empty, change them to "Not yet entered"
for my_row in range(5,SS_rows):
	try: # if the cell's item gives an error at accessing its first character, then it must be a "nan" object (empty cell)
		unused_var = my_data[my_row][11][0]
	except:
		my_data[my_row][11] = "Not yet entered" #NOTE: this did not overwrite the spreadsheet's empty column 11 cell yet. It just changed our own item in my_data[]

# Pandas operations are done now (it just had to read the Excel file once). Next are xlwings operations (basically just WRITING 
#	to the Excel file). Let's open the file with Excel file with xlwings (note this will actually open it in the GUI too).
my_wb = xlwings.Book(SS_path)
my_ws = my_wb.sheets("Sheet1")

# Go through all the employee data in the spreadsheet. For each employee not put in yet: send their surveys and update their statuts on the spreadsheet.
for my_row in range(5,SS_rows):
	if my_data[my_row][11][0:3] == "Not":
		# Operations on this employee are done now. EVERY TIME a survey is sent, the real Excel file is updated with their status. That way if the program quits halfway through unexpectedly, the SS will still be updated (ie will say "Only 1/3 surveys sent")
		fit_test_date = my_data[my_row][0]
		employee_id = my_data[my_row][2]

		# First of all, count how many x's this employee has. AKA how many surveys must be sent for them total. That way, with column 11, we can track the progress of how many are surveys done and how many remain.
		x_total = 0
		for current_col in range(4,11): #go through each mask/failure cell to see if there's an x.
			if my_data[my_row][current_col] == "x" or my_data[my_row][current_col] == "X":
				x_total = x_total + 1

		# We'll also have to count how many x's have been successfully entered into survey so far. Note that this will hopefully come to equal the same value as x_total. Anyway, let's initialize that variable:
		x_done = 0
	
		# Same thing as previously, but instead of just counting the x's, we're gonna send a survey for each x and note the number of surveys sent successfully in column 11
		for current_col in range(4,11): # Note that when their is an x, we pass the column number as the auto_browser() argument for the mask/failure option to send the appropriate survey response
			if my_data[my_row][current_col] == "x" or my_data[my_row][current_col] == "X":

				# I overwrite whatever's in Col 12 (access date) in the REAL Excel spreadsheet's cell, with the current date. This is just for bug checking.
				my_ws.range(indices_to_cell( my_row , 12 )).value = datetime.datetime.today().strftime("%m/%d/%Y")

				# Attempt to submit the survey for the employee's result at the current column
				my_return_value = auto_browser(my_username, my_password, employee_id, fit_test_date, current_col, submitting_for_real )		
				
				# If the submission worked without error, update the status in column 11
				if my_return_value == "SUBMISSION WORKED": # if you told auto_browser to actually submit
					x_done = x_done + 1 
					if x_done == x_total: # Write an "All" phrase to the cell
						my_ws.range(indices_to_cell( my_row , 11 )).value = "All " + str(x_done) + "/" + str(x_total) + " results entered on " + datetime.datetime.today().strftime("%m/%d/%Y")			
					else: # Write an "Only" phrase to the cell
						my_ws.range(indices_to_cell( my_row , 11 )).value = "Only " + str(x_done) + "/" + str(x_total) + " results entered on " + datetime.datetime.today().strftime("%m/%d/%Y")				
						
				if my_return_value == "SUBMISSION WORKED (PRETEND)": # if you told auto_browser to run but not hit submit at the end. (for debug purposes)
					x_done = x_done + 1
					if x_done == x_total: # Write an "All" phrase to the cell
						my_ws.range(indices_to_cell( my_row , 11 )).value = "All " + str(x_done) + "/" + str(x_total) + " results entered on " + datetime.datetime.today().strftime("%m/%d/%Y")	+ "   (PRETEND)"
					else: # Write an "Only" phrase to the cell
						my_ws.range(indices_to_cell( my_row , 11 )).value = "Only " + str(x_done) + "/" + str(x_total) + " results entered on " + datetime.datetime.today().strftime("%m/%d/%Y") + "   (PRETEND)"
					
	elif my_data[my_row][11][0:3] == "Onl": # For this employee, there was some error before that prevented the program from finishing this person's surveys, so I'm gonna leave it undone by the program
		pass
	elif my_data[my_row][11][0:3] == "All": # For this employee, all their surveys were put in successfully so nothing needs to be done now.
		pass
	else: # An unexpected value was found in this employee's column 11
		exit("\nError: an item in the \"Redcap Status\" column is not blank nor was written in (at least, correctly) by \nthe program. Make sure that no one's manually written something in there.\n")

# We've edited the Excel spreadsheet with xlwings, and it was open in the GUI if the user wanted to see live changes happening.
#	But we're done now, so we can close the Excel file which will also close the GUI window.
time.sleep(5)
my_wb.close()



# Ways the program can improve:
# 	- What happens when the user ID put in is incomplete or has no results?
# 	- If you have multiple masks or multiple employee forms to put in, have the Chrome tab open only once and not close till done.
#		till it's done.  (This is not that pressing)

