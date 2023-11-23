from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
import pyodbc
from time import sleep
import sys
import os

current_date = datetime.now().date().strftime("%d-%m-%Y")

custom_folder = 'Whatsapp Message Status/'

if not os.path.exists(custom_folder):
    os.makedirs(custom_folder)

file_name = f"Status of {current_date}.txt"
file_path = os.path.join(custom_folder, file_name)

file = open(file_path, 'a')
sys.stdout = file

total = 0
success = 0
failed = 0

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))


# Microsoft Access Database Connection Parameters
db_path = r"D:\PERSONAL\Alan\Dev Work\Python Whatsapp API\TEST.mdb"  # Replace with your database path
conn_str = (
    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
    f'DBQ={db_path};'
)

# Function to fetch user information from the database
def get_user_info():

    global total

    users = []
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT UserID, PhoneNo, Message FROM WhatsappData WHERE Status <> 'Send' OR Status IS NULL")
            rows = cursor.fetchall()
            total = len(rows)
            for row in rows:
                user_info = {'id': row.UserID, 'phone': row.PhoneNo, 'message': row.Message}
                users.append(user_info)

    except pyodbc.Error as e:
        print(f"Error: {e}")

    cursor.close()
    conn.close()

    return users

# Function to update the status column in the database
def update_status(user_id, status, current_datetime, phone):
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute('UPDATE WhatsappData SET Status = ?, LastMesgUpdateDate = ? WHERE UserID = ?', status, current_datetime, user_id)
            conn.commit()
            print(f"{current_datetime} : +91{phone} Database Updated as {status}")
    except pyodbc.Error as e:
        print(f"{current_datetime} : +91{phone} Error updating status: {e}")
    cursor.close()
    conn.close()


# Function to Start Sending Whatsapp Message
def send_whatsapp_messages(users):

    global success
    global failed
     
    for user in users:
        user_id = user['id']
        phone = user['phone']
        message = user['message']

        custom_format = "%d-%m-%Y %H:%M:%S"
        current_datetime = datetime.now().strftime(custom_format)
  
        message = message

        max_attempts = 3
        attempt_count = 1
        print(f"{current_datetime} : +91{phone} Start to send Message")
        while attempt_count <= max_attempts:
            try:
                link2 = f'https://web.whatsapp.com/send/?phone=+91{phone}&text={message}'
                driver.get(link2)
                print(f"{current_datetime} : +91{phone} Trying {attempt_count}/{max_attempts} to Send Message")
                click_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.CLASS_NAME, '_3XKXx')))
            except Exception as e:
                # print(f"{current_datetime} : +91{phone} Try {attempt_count} failed")
                attempt_count += 1
            else:
                sleep(2)
                click_btn.click()
                sleep(5)
                print(f"{current_datetime} : +91{phone} Message Send Successfully")
                update_status(user_id, 'Send', current_datetime,phone)
                success +=1
                break
        else:
            print(f"{current_datetime} : +91{phone} Failed to Send Message")
            failed += 1
            update_status(user_id, 'Failed', current_datetime, phone)

def is_login():
    link = f'https://web.whatsapp.com/'
    driver.get(link)
    driver.maximize_window()
    try:
        WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'g9p5wyxn')))  
        # _19vUU class name for the whatsapp QR Code return false and then return true
        return True
    except Exception as e:
        print(f"Login failed: {e}")
        return False


if __name__ == "__main__":
    # current_date = datetime.now().date().strftime("%d-%m-%Y")
    print()
    print(f"Starting the Whatsapp Message Sending Application for {current_date}")
    print("---------------------------------------------------------------------")
    print()
    users = get_user_info()
    
    if is_login():
        print("Login Successfull")
        sleep(60)
        driver.minimize_window()
        if users:
            send_whatsapp_messages(users)
            driver.quit()
            print()
            print("-----------------")
            print(f"Total   : {total}")
            print(f"Success : {success}")
            print(f"Failed  : {failed}")
            print("-----------------")
            print()
            print("Application Closed")
        else:
            print("No users found in the database.")
            print("Application Closed")
    else:
        print("Login failed, Please try to start the application again...")


    file.close()
    sys.exit()
