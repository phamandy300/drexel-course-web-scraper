import time as sleepytime
import webbrowser
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from dateutil.parser import parse
from datetime import datetime, time, date, timedelta
import pytz
import csv
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def login_to_drexel(driver):
    """Handle Microsoft authentication and MFA for Drexel"""
    # Get credentials from environment variables
    username = os.getenv('DREXEL_USERNAME')
    password = os.getenv('DREXEL_PASSWORD')
    
    if not username or not password:
        raise ValueError("Please set DREXEL_USERNAME and DREXEL_PASSWORD in your .env file")
    
    try:
        signin_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "_eventId_proceed"))
        )
        signin_button.click()

        # Wait for and fill in email field
        email_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "loginfmt"))
        )
        email_field.send_keys(username)
        email_field.send_keys(Keys.RETURN)
        
        # Wait for and fill in password field
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "passwd"))
        )
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)

        # Handle MFA verification
        print("\nWaiting for you to complete MFA verification...")
        print("Please check your phone and enter the code when prompted.")
        print("The script will continue automatically after successful verification.")
        
        # Wait for successful redirect back to TMS after MFA
        WebDriverWait(driver, 300).until(  # 5 minute timeout for MFA
            lambda driver: "termmasterschedule.drexel.edu" in driver.current_url
        )
        
        print("Successfully logged in with MFA!")
        return True
        
    except Exception as e:
        print(f"Login failed: {str(e)}")
        return False

def sleep():
    sleepytime.sleep(0.1)

def write_to_csv(event_data, csv_writer):
    """Write event data to CSV file"""
    if event_data == 'async':
        return
        
    # Extract all the information from the tuple
    event_name, location, description, start_time, end_time, last_date, days, timestamp = event_data
    
    # Parse description into individual fields
    description_parts = description.split('\n')
    description_dict = {}
    for part in description_parts:
        if ': ' in part:
            key, value = part.split(': ', 1)
            description_dict[key] = value
    
    # Get the days as a comma-separated string
    days_str = ','.join(days) if isinstance(days, list) else days
    
    # Format start and end times
    start_time = datetime.strptime(start_time, '%Y%m%dT%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
    end_time = datetime.strptime(end_time, '%Y%m%dT%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
    last_date = datetime.strptime(last_date, '%Y%m%dT%H%M%S').strftime('%Y-%m-%d')
    
    # Prepare row data
    row_data = {
        'Course': event_name,
        'Location': location,
        'Start Time': start_time,
        'End Time': end_time,
        'End Date': last_date,
        'Days': days_str,
        'Section': description_dict.get('Section', ''),
        'CRN': description_dict.get('CRN', ''),
        'Credits': description_dict.get('Credits', ''),
        'Instructor(s)': description_dict.get('Instructor(s)', ''),
        'Type': description_dict.get('Type', ''),
        'Max Enroll': description_dict.get('Max Enroll', ''),
        'Current Enrollment': description_dict.get('Enroll', ''),
        'College': description_dict.get('College', ''),
        'Department': description_dict.get('Department', ''),
        'Course Description': description_dict.get('Course Description', ''),
        'Section Comments': description_dict.get('Section Comments', ''),
        'Timestamp': timestamp
    }
    
    csv_writer.writerow(row_data)

def readx(xpath_string, wait):
    xpath_search = wait.until(EC.presence_of_element_located((By.XPATH, xpath_string)))
    extracted_text = xpath_search.text
    return extracted_text

def read_page(driver):
    sleep()
    # Wait for the element to be present
    wait = WebDriverWait(driver, 3)
    
    # first check if no time availble so program moved on to next item if so
    Times = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[3]', wait)
    if Times == 'TBD':
        return 'async'

    # Get title/EventName info
    SubjectCode = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]', wait)
    CourseNumber = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]', wait)
    Title = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[6]/td[2]', wait)
    EventName = SubjectCode+' '+CourseNumber+' - '+Title

    # Get location info
    Building = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[5]', wait)
    Room = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[6]', wait)
    Campus = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]', wait)
    if Campus == 'Online' or Campus == 'Remote':
        Location = 'Online'
    else:
        Location = Building+', '+Room+' ('+Campus+')'

    # Get description info (keeping the same format as original)
    Section = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[4]/td[2]', wait)
    Section = 'Section: '+Section
    CRN = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]', wait)
    CRN = 'CRN: '+CRN
    Credits = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]', wait)
    Credits = 'Credits: '+Credits
    Instructors = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[8]/td[2]', wait)
    Instructors = 'Instructor(s): '+Instructors
    InstructionType = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[9]/td[2]', wait)
    InstructionType = 'Type: '+InstructionType
    MaxEnroll = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[11]/td[2]', wait)
    MaxEnroll = 'Max Enroll: '+MaxEnroll
    Enroll = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[12]/td[2]', wait)
    Enroll = 'Enroll: '+Enroll
    SectionComments = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[13]/td[2]/table/tbody/tr/td', wait)
    SectionComments = 'Section Comments: '+SectionComments
    CourseDescription = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div[1]', wait)
    CourseDescription = CourseDescription[20:]
    CourseDescription = 'Course Description: '+CourseDescription
    College = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div[3]', wait)
    College = College[9:]
    College = 'College: '+College
    Department = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div[4]', wait)
    Department = Department[12:]
    Department = 'Department: '+Department
    
    if SectionComments=='Section Comments: None':
        Description = Section+'\n'+Instructors+'\n'+InstructionType+'\n'+Credits+'\n\n'+MaxEnroll+'\n'+Enroll+'\n'+CRN+'\n'+College+'\n'+Department+'\n\n'+CourseDescription
    else:
        Description = Section+'\n'+Instructors+'\n'+InstructionType+'\n'+Credits+'\n\n'+MaxEnroll+'\n'+Enroll+'\n'+CRN+'\n'+College+'\n'+Department+'\n'+SectionComments+'\n\n'+CourseDescription

    # Get time info
    time_format = '%I:%M %p'
    StartTime = datetime.strptime(Times.split(' - ')[0], time_format).time()
    EndTime = datetime.strptime(Times.split(' - ')[1], time_format).time()

    # Get days of the week recurrence info
    Days1L = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[4]', wait)
    Days = []
    if 'TBD' in Days1L: Days.append('TBD')
    else:
        if 'M' in Days1L: Days.append('MO')
        if 'T' in Days1L: Days.append('TU')
        if 'W' in Days1L: Days.append('WE')
        if 'R' in Days1L: Days.append('TH')
        if 'F' in Days1L: Days.append('FR')
        if 'S' in Days1L: Days.append('SA')
        if 'U' in Days1L: Days.append('SU')

    # Get date info (of first occurrence)
    FirstDate = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[1]', wait)
    FirstDate = parse(FirstDate).date()
    if str(FirstDate) == '2023-01-09':
        if 'MO' in Days:
            delta = timedelta(days=0)
        elif 'TU' in Days:
            delta = timedelta(days=1)
        elif 'WE' in Days:
            delta = timedelta(days=2)
        elif 'TH' in Days:
            delta = timedelta(days=3)
        elif 'FR' in Days:
            delta = timedelta(days=4)
        elif 'SA' in Days:
            delta = timedelta(days=5)
        elif 'SU' in Days:
            delta = timedelta(days=6)
        FirstDate += delta
     
    # Combine date and time info
    FirstDateStartTime = datetime.combine(FirstDate, StartTime)
    FirstDateEndTime = datetime.combine(FirstDate, EndTime)
    FirstDateStartTime = FirstDateStartTime.strftime('%Y%m%dT%H%M%S')
    FirstDateEndTime = FirstDateEndTime.strftime('%Y%m%dT%H%M%S')

    # Get date info (of end date)
    LastDate0 = readx('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]', wait)
    LastDate0 = parse(LastDate0).date()
    LastDate = LastDate0.strftime('%Y%m%d')
    LastDate = LastDate+'T000000'
    
    # Get current time
    utc_time = datetime.utcnow()
    est_time = utc_time.astimezone(pytz.timezone('US/Eastern'))
    timestamp = str(est_time)
    
    return EventName, Location, Description, FirstDateStartTime, FirstDateEndTime, LastDate, Days, timestamp

def main():
    # Initialize the CSV file
    csv_filename = 'drexel_courses.csv'
    fieldnames = [
        'Course', 'Location', 'Start Time', 'End Time', 'End Date', 'Days',
        'Section', 'CRN', 'Credits', 'Instructor(s)', 'Type', 'Max Enroll',
        'Current Enrollment', 'College', 'Department', 'Course Description',
        'Section Comments', 'Timestamp'
    ]
    
    # Initialize the driver
    driver = webdriver.Chrome()
    driver.get('https://termmasterschedule.drexel.edu/')
    
    # Handle login and MFA
    if not login_to_drexel(driver):
        print("Failed to log in. Exiting...")
        driver.quit()
        return
    
    # Continue with scraping after successful login
    driver.get('https://termmasterschedule.drexel.edu/webtms_du/collegesSubjects/202415?collCode=')
    
    # Open CSV file and initialize writer
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        csv_writer.writeheader()
        
        # Main scraping loop
        for i in range(1, 100):  # n1
            sleep()
            xpath1 = f'//*[@id="sideLeft"]/a[{i}]'
            wait = WebDriverWait(driver, 5)
            try:
                element = wait.until(EC.presence_of_element_located((By.XPATH, xpath1)))
                element.click()
            except TimeoutException:
                break

            for j in range(1, 50):  # n2
                sleep()
                xpath2 = f'/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[3]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div[{j}]/a'
                try:
                    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath2)))
                    element.click()
                except TimeoutException:
                    break

                for k in range(1, 500):  # n3
                    sleep()
                    xpath3 = f'//*[@id="sortableTable"]/tbody[1]/tr[{k}]/td[6]/span/a'
                    try:
                        element = wait.until(EC.presence_of_element_located((By.XPATH, xpath3)))
                        element.click()
                    except TimeoutException:
                        driver.back()
                        break
                        
                    # Read page data and write to CSV
                    course_data = read_page(driver)
                    write_to_csv(course_data, csv_writer)  # Remove the list wrapping
                    driver.back()

    driver.quit()
    print(f"Scraping complete! Data has been saved to {csv_filename}")

if __name__ == "__main__":
    main()
