# convert date string into datatime using datetime module
from datetime import datetime

# sample string
date_string1 = "04/24/2023"
date_string2 = "04/24/2022"

# convert string to date format
date_object1 = datetime.strptime(date_string1, "%m/%d/%Y")
date_object2 = datetime.strptime(date_string2, "%m/%d/%Y")

# format the date object into mm/dd/yyyy string
formatted_date1 = date_object1.strftime("%m/%d/%Y")
formatted_date2 = date_object2.strftime("%m/%d/%Y")

# print the formatted date string
print(formatted_date1>formatted_date2)
print(date_object1>date_object2)
