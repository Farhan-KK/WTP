from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

# Path to your Excel file
xlsx_file_path = r'path to file'

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:/Users/farha/AppData/Local/Google/Chrome/User Data")
chrome_options.add_argument("--profile-directory=Default")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-extensions")

# Initialize WebDriver with WebDriver Manager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Navigate to WhatsApp Web
driver.get('https://web.whatsapp.com')

# Wait for the user to scan the QR code (if necessary)
print("Please scan the QR code to log in to WhatsApp Web.")
time.sleep(30)  # Increase time if needed to ensure you are logged in

# Function to send WhatsApp message
def send_invites(xlsx_file):
    # Read the Excel file using pandas
    df = pd.read_excel(xlsx_file)
    
    # Check if the 'Number' column exists
    if 'Number' not in df.columns:
        print("Excel file must contain a 'Number' column.")
        return

    for _, row in df.iterrows():
        number = row['Number']
        try:
            # Open the chat with the number
            driver.get(f'https://web.whatsapp.com/send?phone={number}')
            
            # Wait for the message input box to be present and interactable
            message_box = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Type a message" and @contenteditable="true"]'))
            )
            print(f"Message box found for {number}")

            # Type and send the invite link
            invite_link = (
                "invite link"
            )  # Replace 'your_actual_group_invite_link_here' with your actual group invite link
            message_box.clear()  # Clear the box before typing
            message_box.send_keys(invite_link)
            message_box.send_keys('\n')
            
            # Indicate that the invite was sent successfully
            print(f"Sent invite to {number} successfully.")
            time.sleep(5)  # Wait before sending the next invite
        except Exception as e:
            # Handle cases where the invite could not be sent
            print(f"Could not send invite to {number}: {str(e)}")

# Call the function to send invites
send_invites(xlsx_file_path)

# Close the browser after completing the task
driver.quit()
