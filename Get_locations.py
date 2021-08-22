from time import sleep
from selenium import webdriver # selenium is the package, webdriver is the module
from selenium.common.exceptions import NoSuchElementException as NE
import pandas
import openpyxl
import win32com.client as win32

#search_query = input("What nearby place you want to see?")

search_query = input("What is your search wish?")
search_text = search_query + " " + "near me"

Driver = webdriver.Chrome(executable_path="C:\drivers\chromedriver")
Driver.maximize_window()
Driver.get("https://www.google.com/")
#searching the string
Driver.find_element_by_xpath("//input[@title='Search']").send_keys(search_text)
Driver.find_element_by_xpath("//input[@title='Search']").send_keys(u'\ue007')

#click on View all button
try:
    Driver.find_element_by_xpath("//div[@class='MXl0lf mtqGb']").click()
except NE:
    print("No proper search results found.")
    sleep(2)
    user_response = input("Try again? ")
    user_response.capitalize()
    while user_response == "YES":
        search_query = input("What is your search wish?")
        search_text = search_query + " " + "near me"
        Driver.back()
        Driver.find_element_by_xpath("//input[@title='Search']").send_keys(search_text)
        Driver.find_element_by_xpath("//input[@title='Search']").send_keys(u'\ue007')
        try:
            Driver.find_element_by_xpath("//div[@class='MXl0lf mtqGb']").click()
        except NE:
            print("No proper search results found.")
            user_response = input("Try again? ")
            user_response.capitalize()
        if user_response != "YES":
            print("Response not understood.")
            user_response = input("Try again? ")
            user_response.capitalize()
    Driver.close()


#Locate the results in the pages:
Location_title = []
Avg_Rating = []
Number_of_ratings_received = []

def creating_the_name_list():
    elements_found = Driver.find_elements_by_xpath("//div[@class='dbg0pd']/div")
    for element in elements_found:
        print(element.text)
        Location_title.append(element.text)

def creating_the_rating_list():
    overallratings = Driver.find_elements_by_xpath("//div[@class='rllt__details']/div[1]")
    Number_of_ratings = Driver.find_elements_by_xpath("//div[@class='rllt__details']/div[1]/span[2]")
    for rate in overallratings:
        if "No reviews" in rate.text:
            Avg_Rating.append("No reviews")
            continue
        Avg_Rating.append(rate.text[0:3])
    for number in Number_of_ratings:
        Number_of_ratings_received.append(number.text)

try:
    NextPage_link = Driver.find_element_by_xpath("//a[@id='pnnext']")
    Next_page_present = NextPage_link.is_displayed()
except NE:
    creating_the_name_list()
    creating_the_rating_list()

while Next_page_present:
    creating_the_name_list()
    creating_the_rating_list()
    try:
        NextPage_link = Driver.find_element_by_xpath("//a[@id='pnnext']")
        Next_page_present = NextPage_link.is_displayed()
        NextPage_link.click()
        sleep(5)
    except NE:
        print('No next link found')
        break

print("Last page reached")
Location_title_Final_list = sorted(Location_title)
print(set(Location_title_Final_list))
print(Avg_Rating)
#print(Number_of_ratings_received)
Final_rating_list = list(map(lambda x, y: x + ' out of ' + y, Avg_Rating, Number_of_ratings_received))
print(len(Location_title_Final_list)==len(Final_rating_list))
print(Final_rating_list)
Driver.close()

#loading the data collected to an excel file
data = pandas.DataFrame()
data['Shop Name'] = Location_title
data['Average Rating'] = Avg_Rating
#Name_of_file = 'GoogleData.xlsx'
data.to_excel('GoogleData.xlsx', index = False)

#adjusting the coloumn widths to fit the data
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r'C:\Users\rupsh\OneDrive\Desktop\exploring_python\Google search project\GoogleData.xlsx')
ws = wb.Worksheets("Sheet1")
ws.Columns.AutoFit()
wb.Save()
excel.Application.Quit()
