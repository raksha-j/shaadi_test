import openpyxl
import os
import platform
import random
import subprocess
import sys
import time
from selenium import webdriver
from selenium.webdriver import ActionChains

# Identifies the Platfrom 
def platfrm():
    if sys.platform.startswith('linux'):
        platform = 'linux'
    elif sys.platform == 'darwin':
        platform = 'mac'
    elif sys.platform.startswith('win'):
        platform = 'win'
    else:
        raise RuntimeError('Could not determine this platform.')
    return platform
print('Platfrom is identified as '+platfrm())
print('_______________________________________')

#Identifes the installed Chrome Version
def chrome_version():
	osname = platform.system()
	if osname == 'Linux':
		with subprocess.Popen(['chromium-browser', '--version'], stdout=subprocess.PIPE) as proc:
			version = proc.stdout.read().decode('utf-8').replace('Chromium', '').strip()
			version = version.replace('Google Chrome', '').strip()
			version = 'ver' + version[0:2]
	elif osname == 'Darwin':
		process = subprocess.Popen(['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'], stdout=subprocess.PIPE)
		version = process.communicate()[0].decode('UTF-8').replace('Google Chrome', '').strip()
		version = 'ver' + version[0:2]
	elif osname == 'Windows':
		process = subprocess.Popen(['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', '/v', 'version'],stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
		version = process.communicate()[0].decode('UTF-8').strip().split()[-1]
		version = 'ver' + version[0:2]
	else:
		return
	return version
    
print('Chrome Version is identified as '+ chrome_version())
print('________________________________________________')

# Get the drivr based on Platfrom and Chrome version
def getchrdriver():
	if platfrm() == 'mac':
		path = os.getcwd()+'/chdriver/'+ chrome_version()+'/chromedriver'+ platfrm()
		os.system('xattr -d com.apple.quarantine ' + path)
	elif platfrm()  == 'win':
		path =  os.getcwd() +'\chdriver/'+chrome_version()+'\chromedriver.exe'
	elif platfrm()  == 'linux':
		path =  os.getcwd() +'/chdriver/'+chrome_version()+'/chromedriver'+ platfrm() 
	else:
		print('Unoknow OS')
	return path
print('Chrome Driver path is '+getchrdriver())
print('_____________________________________________________________________')

#Load the excel for parameters 
wb = openpyxl.load_workbook(os.getcwd() + '/data/MT_data.xlsx')
sheet = wb['Sheet']

# read vlaues from the cell 
def copyrange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
#Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
#Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
#Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
    return rangeSelected

# loop to count number of MT's
for cell in sheet["A"]:
    if cell.value is None:
        break
else:
    cell.row + 1
# removing one row as it is column names 
a = cell.row - 1

driver = webdriver.Chrome(executable_path= getchrdriver())

# List with community domain for refrence
# List_community = ['Hindi', 'Marathi','Punjabi', 'Bengali', 'Gujarati', 'Tamil', 'Odia', 'Kannada', 'Telugu', 'Urdu']

print('Begin of the Test case')
print('______________________')
print('Result: ')

#Loop for number of itterations for number of request to be made
#Data to select
for i in range(2,a):
	rangeSelected = copyrange(1,i,1,a,sheet)
	list_community = rangeSelected
	for i in range(len(list_community[0])):
		driver.get("https://www." + list_community[0][i] + "shaadi.com")
		driver.maximize_window()
		time.sleep(3)
# Clicks on  Lets begin
		login_lnk = driver.find_element_by_xpath("//button[@data-testid='lets_begin']")
		login_lnk.click()
# Enter the email id
		login_mail = driver.find_element_by_xpath("//input[@data-testid='email']")
		login_mail.send_keys('test' + str(random.randint(0, 100000000)) + '@tesst.com')
# Enter the password
		login_passwd = driver.find_element_by_xpath("//input[@data-testid='password1']")
		login_passwd.send_keys('sh@@ditim3')
# Clicks on Drop down
		login_dropdown = driver.find_element_by_xpath("//div[@class ='Dropdown-control postedby_selector' ]")
		login_dropdown.click()
# Selects an Option
		login_selectdropdown = driver.find_element_by_xpath("// div[@class ='Dropdown-option'][1]")
		login_selectdropdown.click()
		time.sleep(3)
# Selects a radiobutton
		login_gender = driver.find_element_by_id("gender_female")
		login_gender.click()
# Click on Next button
		login_nextbttn = driver.find_element_by_xpath("//button[@data-testid='next_button']")
		login_nextbttn.click()
		time.sleep(5)
# Gets text of Default Mother Tongue
		mt = driver.find_element_by_xpath("//div[@class='Dropdown-control mother_tongue_selector Dropdown-disabled']/div").text
# Gets the current url
		urlt = driver.current_url
# Compares the Mother Tongue and the current Url
		if mt == list_community[0][i] and mt[0:3].lower() == urlt[12:15]:
			print(list_community[0][i] + ' mother tongue is selected by default on ' + urlt[12:].strip('/'))
		else:
			print('Default selection is not working')
print('_____________________________')  
print('Test case execution completed')
print('_____________________________')      
# Close the browser
driver.quit()

