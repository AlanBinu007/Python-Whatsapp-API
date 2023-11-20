from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import pyodbc
from time import sleep
import time
import sys


# file = open('Whatsapp Message Status.txt', 'a')
# sys.stdout = file


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))


# Microsoft Access Database Connection Parameters
db_path = r"D:\PERSONAL\Alan\Dev Work\Python Whatsapp API\UserDetails.accdb"  # Replace with your database path
conn_str = (
    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
    f'DBQ={db_path};'
)

# Function to fetch user information from the database
def get_user_info():
    users = []
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT UserID, UserName, PhoneNo, Status, Message FROM WhatsappData')
            rows = cursor.fetchall()
            for row in rows:
                user_info = {'id': row.UserID, 'name': row.UserName, 'phone': row.PhoneNo, 'message': row.Message, 'status': row.Status}
                users.append(user_info)

    except pyodbc.Error as e:
        print(f"Error: {e}")

    return users

# Function to update the status column in the database
def update_status(user_id, status, current_datetime):
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute('UPDATE WhatsappData SET Status = ?, LastMesgUpdateDate = ? WHERE UserID = ?', status, current_datetime, user_id)
            conn.commit()
            print(f"{current_datetime} : Database Updated")
    except pyodbc.Error as e:
        print(f"Error updating status: {e}")

# Function to Start Sending Whatsapp Message
def send_whatsapp_messages(users):

    link = f'https://web.whatsapp.com/'
    driver.get(link)
    driver.maximize_window()
    input("Press ENTER after login into Whatsapp Web and your chats are visiable.")
     
    for user in users:
        user_id = user['id']
        name = user['name']
        phone = user['phone']
        status = user['status']
        message = user['message']

        custom_format = "%d-%m-%Y %H:%M:%S"
        current_datetime = datetime.now().strftime(custom_format)

        if status != 'Sent':  
            message = message
            # driver.get(link2)
            # #Wait to 1 min to fully load the page
            # time.sleep(60)
            # actions = ActionChains(driver)
            # actions.send_keys(Keys.ENTER)
            # actions.perform()
            # # print(f"{current_datetime} :+91{phone} Sending Message")
            # #Press Enter and wait for 5 Second
            # time.sleep(5)
            max_attempts = 3
            attempt_count = 1

            while attempt_count <= max_attempts:
                try:
                    link2 = f'https://web.whatsapp.com/send/?phone=+91{phone}&text={message}'
                    driver.get(link2)
                    print(f"{current_datetime} : +91{phone} Start to send Message")
                    print(f"{current_datetime} : +91{phone} Trying {attempt_count} out of {max_attempts} to Send Message")
                    click_btn = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, '_3XKXx')))
                except Exception as e:
                    print(f"{current_datetime} : +91{phone} Try {attempt_count} failed")
                    # print(f"{current_datetime} : +91{phone} Failed to Send Message")
                    # update_status(user_id, 'Failed', current_datetime)
                    attempt_count += 1
                else:
                    sleep(2)
                    click_btn.click()
                    Flag = 1
                    sleep(5)
                    print(f"{current_datetime} : +91{phone} Message Send Successfully")
                    update_status(user_id, 'Send', current_datetime)
                    break
            else:
                print(f"{current_datetime} : +91{phone} Failed to Send Message")
                update_status(user_id, 'Failed', current_datetime)


if __name__ == "__main__":
    current_date = datetime.now().date().strftime("%d-%m-%Y")
    print()
    print()
    print()
    print()
    print(f"Starting the Whatsapp Message Sending Application for {current_date}")
    print("---------------------------------------------------------------------")
    users = get_user_info()
    if users:
        send_whatsapp_messages(users)
        driver.quit()
        print("Application Closed")
    else:
        print("No users found in the database.")
        print("Application Closed")

    # file.close()
