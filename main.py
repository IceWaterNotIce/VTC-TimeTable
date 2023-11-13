import time
import re
import uuid
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from datetime import datetime, timedelta
from webdriver_manager.chrome import ChromeDriverManager


options = webdriver.ChromeOptions()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

driver.get("https://myportal.vtc.edu.hk/wps/portal")
wait = WebDriverWait(driver, 10)  # Maximum wait time of 10 seconds
wait.until(EC.presence_of_element_located((By.TAG_NAME, "html")))

Your_CNA = input("Your CNA:") # @param {type:"string"}
Your_Password = input("Your Password:") # @param {type:"string"}
#region 登入 VTC Portal

input_text_userid = driver.find_element(By.NAME, "userid")
input_text_password = driver.find_element(By.NAME, "password")
input_btn_login = driver.find_element(By.ID, "submitBtn")


input_text_userid.send_keys(Your_CNA)
input_text_password.send_keys(Your_Password)
input_btn_login.click()
#endregion

#region 開啟時間表


wait = WebDriverWait(driver, 10) 

schoolname = re.split(r':',driver.find_element(By.XPATH, "//*[@id='banner']/div[2]/div[2]/div[1]/div[1]").text)[-1]

a_Langzhtw = driver.find_element(By.XPATH, "//a[@href='#' and contains(@onclick, 'ChangeLang')]")
# Execute JavaScript to click the element
driver.execute_script("arguments[0].click();", a_Langzhtw)

a_TimeTable = driver.find_element(By.XPATH, "//*[@id='hkvtcsp_menu']/span[3]/a")
driver.execute_script("arguments[0].click();", a_TimeTable)

img_TimeTable = driver.find_element(By.XPATH, '//*[@id="theme-content"]/div[2]/div/div[2]/div[1]/div/center/table/tbody/tr/td[1]/table/tbody/tr[1]/td/img')
driver.execute_script("arguments[0].click();", img_TimeTable)

# Wait for the page to finish loading
time.sleep(2)

select_EndWeek = driver.find_element(By.ID, 'j_id_7:beanDateTo')

# Create a Select object
select = Select(select_EndWeek)
# Select the last option
options = select.options
last_option = options[-1]
last_option.click()

start_date = datetime.strptime(options[0].text[4:].split(" - ")[0], "%d-%b-%Y")

#print(start_date)
input_btn_search = driver.find_element(By.ID, 'j_id_7:search')
input_btn_search.click()
#endregion
#region 取得時間表內容
time.sleep(2)

table_TimeTableResult = driver.find_element(By.ID, 'tableResult')


def get_formatted_date(start_date, week_number, day_of_week):

    # Define the week number and day of the week
    # Monday is 1, Tuesday is 2, and so on

    # Calculate the specific date
    specific_date = start_date + timedelta(weeks=week_number-1, days=day_of_week)

    # Format the date as "DD-MMM-YYYY"
    formatted_date = specific_date.strftime("%Y%m%d")
    #print(formatted_date)
    return formatted_date

def get_lst_modified_weeks(range_notation):

    result = []
    ranges = range_notation.split(",")
    for r in ranges:
        if "-" in r:
            start, end = map(int, r.split("-"))
            result.extend(range(start, end + 1))
        else:
            result.append(int(r))

    result = sorted(result)
    return result

td_elements = table_TimeTableResult.find_elements(By.TAG_NAME, 'td')

row_index = 0
column_index = 0
column_length = 7
blank_td_list = []


file_path = '.\\YourTimeTable.ics'
with open(file_path, 'w') as file:
    file.write("")
with open(file_path, 'a') as file:
    file.write("BEGIN:VCALENDAR"\
                + "\n" +"PRODID:-//Google Inc//Google Calendar 70.9054//EN"\
                + "\n" +"VERSION:2.0"\
                + "\n" +"CALSCALE:GREGORIAN"\
                + "\n" +"METHOD:PUBLISH"\
                + "\n" +"X-WR-CALNAME:VTC-TimeTable"\
                + "\n" +"X-WR-TIMEZONE:Asia/Hong_Kong"+ "\n")

for td_element in td_elements:

    if column_index % column_length == 0 :
        row_index += 1
        column_index = 0
        print("Row : " + str(row_index))

    column_index += 1

    while [row_index, column_index] in blank_td_list:
        column_index += 1

    rowspan = td_element.get_attribute('rowspan')
    if rowspan:
        for i in range(row_index + 1, row_index + int(rowspan)):
            blank_td_list.append([i,column_index])
        #print(blank_td_list)
    #print(str(column_index) + " : " + td_element.text)
    
    day = ""
    if column_index == 1:day = "Monday"
    elif column_index == 2:day = "Tuesday"
    elif column_index == 3:day = "Wednesday"
    elif column_index == 4:day = "Thursday"
    elif column_index == 5:day = "Friday"
    elif column_index == 6:day = "Saturday"
    
    if td_element.text.strip() != "" and column_index != 1:

        # Use regular expressions to extract the information
        lines = td_element.text.split("\n")
        lst_lesson_details = lines[:5]
        if len(lst_lesson_details) < 5:
            continue
        lst_lesson_details.append(column_index-1)
        """lst_lesson_details eg
        [   index   content
            0       'ITE3107 (Tutorial (T) )',      # Course Name
            1       '(09:30 - 11:30)',              # Time 
            2       'DL-CS-LW217',                  # Location
            3       'TANG KING SHING',              # Instructor
            4       'Wk:2-5,7,8,10-17',             # Weeks
            5       '1'                             # Day of the week
        ]
        """
        print(lst_lesson_details)

        lst_WeekNums =  get_lst_modified_weeks(lst_lesson_details[4][3:])

        time_range = lst_lesson_details[1].replace("(", "").replace(")", "")
        start_time, end_time = time_range.split(" - ")

        # Remove the colon ":" from the time strings
        start_time = start_time.replace(":", "")
        end_time = end_time.replace(":", "")

        # Append "00" to represent seconds
        start_time += "00"
        end_time += "00"

        for int_WeekNum in lst_WeekNums:
            #print(int_WeekNum)
            str_VEVENT = "BEGIN:VEVENT"\
                + "\n" + "DTSTART:" + str(get_formatted_date(start_date, int(int_WeekNum), int(lst_lesson_details[5]))) + "T" + start_time + "Z"\
                + "\n" + "DTEND:"   + str(get_formatted_date(start_date, int(int_WeekNum), int(lst_lesson_details[5]))) + "T" + end_time   + "Z"\
                + "\n" + "DTSTAMP:" + datetime.now().strftime("%Y%m%dT%H%M%SZ")\
                + "\n" + "UID:"     + f"event-{str(uuid.uuid4())}@example.com"\
                + "\n" + "CREATED:" + datetime.now().strftime("%Y%m%dT%H%M%SZ")\
                + "\n" + "DESCRIPTION:"\
                + "\n" + "LAST-MODIFIED:" + datetime.now().strftime("%Y%m%dT%H%M%SZ")\
                + "\n" + "LOCATION:"+ schoolname + " " +lst_lesson_details[2]\
                + "\n" + "SEQUENCE:0"\
                + "\n" + "STATUS:CONFIRMED"\
                + "\n" + "SUMMARY:" + lst_lesson_details[0]\
                + "\n" + "TRANSP:OPAQUE"\
                + "\n" + "END:VEVENT"\
                + "\n" 
            #print(str_VEVENT)

            # 指定文件路径和要写入的内容
            
            file_path = '.\\YourTimeTable.ics'
            with open(file_path, 'a') as file:
                file.write(str_VEVENT)
                
file_path = '.\\YourTimeTable.ics'
with open(file_path, 'a') as file:
    file.write("END:VCALENDAR")
#endregion
#region 製作日歷

#endregion

# Add a pause to keep the browser window open
#input("Press Enter to continue...")
