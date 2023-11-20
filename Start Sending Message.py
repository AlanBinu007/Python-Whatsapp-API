from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import pyodbc
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
            print(f"{current_datetime} :Database Updated")
    except pyodbc.Error as e:
        print(f"Error updating status: {e}")

# Function to send WhatsApp messages using pywhatkit
def send_whatsapp_messages(users):

    link = f'https://web.whatsapp.com/'
    driver.get(link)
    driver.maximize_window()
    time.sleep(90)
     
    for user in users:
        user_id = user['id']
        name = user['name']
        phone = user['phone']
        status = user['status']
        message = user['message']

        current_datetime = datetime.now()

        if status != 'Sent':  
            message = message
            link2 = f'https://web.whatsapp.com/send/?phone=+91{phone}&text={message}'
            print("--------------------------------------------------------------------------")
            # print(f"{current_datetime} :+91{phone} Opening Chat")
            driver.get(link2)
            #Wait to 1 min to fully load the page
            time.sleep(30)
            actions = ActionChains(driver)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            # print(f"{current_datetime} :+91{phone} Sending Message")
            #Press Enter and wwait for 5 Second
            time.sleep(5)
            print(f"{current_datetime} :+91{phone} Message Successfully sent to {phone}: {message}")
            update_status(user_id, 'Send', current_datetime)


if __name__ == "__main__":
    current_date = datetime.now().date().strftime("%d-%m-%Y")
    print()
    print()
    print()
    print()
    print(f"Starting the Whatsapp Message Sending Application for {current_date}")
    users = get_user_info()
    if users:
        send_whatsapp_messages(users)
        driver.quit()
        print("Application Closed")
    else:
        print("No users found in the database.")
        print("Application Closed")

    # file.close()
