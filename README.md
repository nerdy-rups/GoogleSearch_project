# GoogleSearch_project
This project is a small utility created to help with common places available in Google search and listed as a Google business. 

Enter a search query and the program will try to search the same with "near me" appended and give out the results(place title and average rating) in an excel file. 

Progamming Language: Python

Packages used: Selenium, time, openpyxl, pandas, pywin32, os

Some known issues: 
1. Capturing the total number of ratings available against the average rating is not appropriate. It does not have a full exact match with the place name or average rating.
2. Inputs that do not retrieve any search result in the required format go into a loop that isn't properly tested. 
