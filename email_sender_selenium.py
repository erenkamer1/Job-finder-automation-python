from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import random
import os
import json
from datetime import datetime, timedelta
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException

# Constants for email tracking system
DAILY_EMAIL_LIMIT = 60
EMAIL_TRACKING_FILE = 'email_tracking.json'

def get_current_date():
    """Returns current date in YYYY-MM-DD format"""
    return datetime.now().strftime('%Y-%m-%d')

def load_email_tracking():
    if os.path.exists(EMAIL_TRACKING_FILE):
        with open(EMAIL_TRACKING_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_email_tracking(tracking_data):
    with open(EMAIL_TRACKING_FILE, 'w') as f:
        json.dump(tracking_data, f, indent=4)

def get_today_sent_count(tracking_data, email):
    today = get_current_date()
    if email in tracking_data and today in tracking_data[email]:
        return tracking_data[email][today]
    return 0

def update_email_tracking(tracking_data, email):
    today = get_current_date()
    if email not in tracking_data:
        tracking_data[email] = {}
    if today not in tracking_data[email]:
        tracking_data[email][today] = 0
    tracking_data[email][today] += 1
    save_email_tracking(tracking_data)

def clean_old_tracking_data(tracking_data):
    today = datetime.strptime(get_current_date(), '%Y-%m-%d')
    for email in list(tracking_data.keys()):
        for date_str in list(tracking_data[email].keys()):
            date = datetime.strptime(date_str, '%Y-%m-%d')
            if (today - date).days > 7:  # Clean records older than 7 days
                del tracking_data[email][date_str]
    save_email_tracking(tracking_data)

def setup_driver(profile_name):
    options = Options()
    
    # Create separate profile directory for each account
    user_data_dir = os.path.join(os.getcwd(), "ChromeProfiles", profile_name)
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)
    
    # Disable security warnings
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-extensions")
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    
    # Set Chrome profile
    options.add_argument(f'--user-data-dir={user_data_dir}')
    options.add_argument('--start-maximized')
    
    # Hide Chrome's automation mode
    options.add_argument("--disable-blink-features")
    options.add_argument('--disable-blink-features=AutomationControlled')
    
    driver = webdriver.Chrome(options=options)
    
    # Remove automation flags using JavaScript
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def gmail_login(driver, email, password):
    try:
        print(f"[*] Checking Gmail for {email} account...")
        driver.get("https://gmail.com")
        time.sleep(3)
        
        # Check if already on Gmail main page
        if "mail.google.com" in driver.current_url:
            print("[+] Gmail session already active")
            return True
            
        # If not logged in, proceed with normal login
        print(f"[*] Logging into {email} account...")
        driver.get("https://accounts.google.com/signin")
        time.sleep(2)

        # Email input
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
        )
        email_input.send_keys(email)
        
        # Next button
        next_button = driver.find_element(By.CSS_SELECTOR, "button[jsname='LgbsSe']")
        next_button.click()
        time.sleep(2)

        # Password input
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
        )
        password_input.send_keys(password)
        
        # Sign in button
        sign_in_button = driver.find_element(By.CSS_SELECTOR, "button[jsname='LgbsSe']")
        sign_in_button.click()
        time.sleep(5)

        return True
    except Exception as e:
        print(f"[!] Error during login: {str(e)}")
        return False

def send_email(driver, recipient, subject, body, cv_path):
    try:
        # Go to Gmail and wait for full load
        if 'gmail.com' not in driver.current_url:
            driver.get('https://gmail.com')
            print("[*] Loading Gmail (3 seconds)...")
            time.sleep(3)

        print("[*] Clicking Compose button...")
        # Click Compose button
        try:
            compose_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[text()='Compose']"))
            )
        except:
            compose_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".T-I.T-I-KE.L3"))
            )

        compose_btn.click()
        print("[*] Opening compose window (2 seconds)...")
        time.sleep(2)

        print("[*] Writing recipient address...")
        # Fill recipient field
        try:
            to_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[role='combobox']"))
            )
        except:
            to_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "to"))
            )
        
        to_field.send_keys(recipient)
        print("[*] Recipient written (1 second)...")
        time.sleep(1)

        print("[*] Writing subject...")
        # Fill subject field
        try:
            subject_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "subjectbox"))
            )
        except:
            subject_field = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Subject']")
        
        subject_field.send_keys(subject)
        print("[*] Subject written (1 second)...")
        time.sleep(1)

        print("[*] Writing email body...")
        # Fill message body
        try:
            body_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='textbox']"))
            )
        except:
            body_field = driver.find_element(By.CSS_SELECTOR, ".Am.Al.editable")
        
        body_field.send_keys(body)
        print("[*] Email body written (1 second)...")
        time.sleep(1)

        print("[*] Attaching CV file...")
        # Attach file
        try:
            attach_btn = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[command='Files']"))
            )
        except:
            try:
                attach_btn = driver.find_element(By.CSS_SELECTOR, "div[aria-label='Add attachment']")
            except:
                attach_btn = driver.find_element(By.CSS_SELECTOR, "div[aria-label='Attach files']")

        attach_btn.click()
        print("[*] Opening file selection window (2 seconds)...")
        time.sleep(2)

        # Send file path
        import pyautogui
        print(f"[*] Writing CV file path: {cv_path}")
        pyautogui.write(cv_path)
        time.sleep(1)
        print("[*] Loading CV file...")
        pyautogui.press('enter')
        print("[*] Completing CV upload (3 seconds)...")
        time.sleep(3)

        print("[*] Sending email...")
        # Error checking during email sending
        try:
            # Click send button
            try:
                send_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div[role='button'][aria-label*='Send']"))
                )
            except:
                send_btn = driver.find_element(By.CSS_SELECTOR, ".T-I.J-J5-Ji.aoO.v7.T-I-atl.L3")

            send_btn.click()
            time.sleep(2)

            # Check for error pop-up
            try:
                error_dialog = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='alertdialog']"))
                )
                error_text = error_dialog.text
                print(f"[!] Error sending email: {error_text}")
                return False
            except TimeoutException:
                # No error dialog found, email sent successfully
                print("[+] Email sent successfully!")
                return True

        except Exception as e:
            print(f"[!] Error clicking send button: {str(e)}")
            return False

    except Exception as e:
        print(f"[!] Error sending email: {str(e)}")
        return False

def main():
    try:
        # Load email tracking data
        tracking_data = load_email_tracking()
        clean_old_tracking_data(tracking_data)

        # Load email list from Excel
        df = pd.read_excel('email_results.xlsx')
        
        # Get Gmail credentials
        email = input("Enter Gmail address: ")
        password = input("Enter Gmail password: ")
        
        # Get CV file path
        cv_path = input("Enter full path to CV file: ")
        if not os.path.exists(cv_path):
            print("[!] CV file not found!")
            return
        
        # Get email subject and body
        subject = input("Enter email subject: ")
        print("Enter email body (press Ctrl+Z and Enter when done):")
        body = ""
        try:
            while True:
                line = input()
                body += line + "\n"
        except EOFError:
            pass

        # Setup Chrome driver
        driver = setup_driver(email)

        # Login to Gmail
        if not gmail_login(driver, email, password):
            print("[!] Gmail login failed!")
            driver.quit()
            return

        # Check daily email limit
        today_sent = get_today_sent_count(tracking_data, email)
        if today_sent >= DAILY_EMAIL_LIMIT:
            print(f"[!] Daily email limit ({DAILY_EMAIL_LIMIT}) reached for {email}")
            driver.quit()
            return

        # Process each email in the list
        success_count = 0
        error_count = 0
        
        for index, row in df.iterrows():
            recipient = row['Email']
            company = row['Company']
            
            if recipient == '-' or pd.isna(recipient):
                continue
                
            print(f"\n[*] Processing {company} ({recipient})")
            
            # Check if we've hit the daily limit
            if get_today_sent_count(tracking_data, email) >= DAILY_EMAIL_LIMIT:
                print(f"[!] Daily email limit ({DAILY_EMAIL_LIMIT}) reached")
                break
                
            # Random delay between emails
            delay = random.uniform(30, 60)
            print(f"[*] Waiting {delay:.1f} seconds before sending next email...")
            time.sleep(delay)
            
            if send_email(driver, recipient, subject, body, cv_path):
                success_count += 1
                update_email_tracking(tracking_data, email)
            else:
                error_count += 1
                
            print(f"[*] Progress: {success_count} sent, {error_count} failed")
            
        print(f"\n[+] Email sending completed!")
        print(f"[+] Total sent: {success_count}")
        print(f"[+] Total failed: {error_count}")
        
    except Exception as e:
        print(f"[!] Critical error: {str(e)}")
    finally:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[!] Program interrupted by user")
    except Exception as e:
        print(f"\n[!] Critical error: {str(e)}") 