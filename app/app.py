import sys
import traceback
from random import randrange
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import numpy as np
import time
import argparse
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

class WhatsappMessage:
    """
    A class that encapsulates WhatsApp Message automation
    functions and attributes
    """

    def __init__(self, sheet_name, image_path=None, poll_sheet_name=None):
        self.sheet_name = sheet_name
        self.image_path = image_path
        self.poll_sheet_name = poll_sheet_name
        self.excel_data = None
        self.poll_data = None
        self.driver = None
        self.driver_wait = None

    def start_process(self):
        self.read_data()
        self.load_driver()
        input("Press Enter after scanning the QR code and WhatsApp has loaded completely...")
        self.process_messages()

    def read_data(self):
        # Read data from the provided Google Sheets URL
        print(f"Reading data from sheet: {self.sheet_name}")
        self.excel_data = pd.read_excel(
            'https://docs.google.com/spreadsheets/d/1QEl5TpOb_NX46yknlOM6w7N1zsQvNT2zj0-u_duU-4E/export?format=xlsx',
            sheet_name=self.sheet_name, dtype={'Contact': np.int64}, engine='openpyxl')
        print(self.excel_data)

        if self.poll_sheet_name:
            print(f"Reading poll data from sheet: {self.poll_sheet_name}")
            self.poll_data = pd.read_excel(
                'https://docs.google.com/spreadsheets/d/1QEl5TpOb_NX46yknlOM6w7N1zsQvNT2zj0-u_duU-4E/export?format=xlsx',
                sheet_name=self.poll_sheet_name, engine='openpyxl')
            print(self.poll_data)

    def load_driver(self):
        chrome_options = Options()
        chrome_options.add_argument('--user-data-dir=./User_Data')
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service)
        self.driver_wait = WebDriverWait(self.driver, 20)
        # Open WhatsApp URL in Chrome browser
        self.driver.get("https://web.whatsapp.com")

    def process_messages(self):
        success_count = 0
        fail_count = 0

        # Loop through each contact and its corresponding message
        for index, row in self.excel_data.iterrows():
            contact_number = str(row['Contact'])
            message = row['Message']

            # Ensure the contact number is in the correct format
            if len(contact_number) == 10 or not contact_number.startswith('+91'):
                contact_number = '+91' + contact_number

            # Construct the WhatsApp URL for the contact
            user_url = f'https://web.whatsapp.com/send?phone={contact_number}'
            self.driver.get(user_url)
            time.sleep(5)  # Wait for the page to load

            if self.is_valid_chat():
                print(f'Ready to send message to {contact_number}')
                self.attach_image()  # Attach an image if provided
                self.send_message(message)  # Send the correct message
                if self.poll_sheet_name:
                    self.attach_poll()  # Attach a poll if provided
                success_count += 1
                print(f"Message sent to {contact_number}")
            else:
                fail_count += 1
                print(f"Failed to send message to {contact_number}")

            # Wait before sending the next message to avoid spamming
            timelapse = randrange(3, 10)
            print(f'Sleeping for: {timelapse} seconds')
            time.sleep(timelapse)

        print(f"Messages sent successfully: {success_count}")
        print(f"Messages failed: {fail_count}")

    def is_valid_chat(self):
        try:
            # Wait for the element indicating an invalid chat
            self.driver_wait.until(EC.presence_of_element_located(
                (By.XPATH, "//div[contains(text(),'Phone number shared via url is invalid.')]")))
            return False
        except TimeoutException:
            # If the element is not found within the timeout, consider it a valid chat
            return True

    def attach_image(self):
        if self.image_path is not None:
            print('Attaching image...')
            attachment_button_path = "//div[@title='Attach']"
            image_input_xpath = "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']"

            try:
                # Wait for the attachment button to be clickable
                attachment_button = self.driver_wait.until(
                    EC.element_to_be_clickable((By.XPATH, attachment_button_path))
                )
                attachment_button.click()
                time.sleep(2)  # Wait for the options to appear

                # Find the image input element and send the file path to it
                image_input = self.driver.find_element(By.XPATH, image_input_xpath)
                image_input.send_keys(self.image_path)
                time.sleep(5)  # Wait for the image to upload

                # Click the send button
                send_button_xpath = "//span[@data-icon='send']"
                send_button = self.driver.find_element(By.XPATH, send_button_xpath)
                send_button.click()
                time.sleep(2)

            except TimeoutException as e:
                print("Failed to attach image within the given time.", e)
                traceback.print_exc()

    def attach_poll(self):
        if self.poll_data is not None:
            print('Attaching poll...')
            attachment_button_path = "//div[@title='Attach']"

            try:
                # Wait for the attachment button to be clickable
                attachment_button = self.driver_wait.until(
                    EC.element_to_be_clickable((By.XPATH, attachment_button_path))
                )
                attachment_button.click()
                time.sleep(2)  # Wait for the options to appear

                poll_button_path = "//span[normalize-space()='Poll']"
                poll_button = self.driver_wait.until(
                    EC.element_to_be_clickable((By.XPATH, poll_button_path))
                )
                poll_button.click()
                time.sleep(2)  # Wait for the poll dialog to appear

                # Enter the poll question
                question_input = self.driver.find_element(By.XPATH, "//div[contains(@class,'_13NKt copyable-text selectable-text')]")
                question_input.click()
                question_input.clear()
                question_input.send_keys(self.poll_data['Question'][0])

                # Enter the poll options
                for i, option in enumerate(self.poll_data.columns[1:], start=1):
                    option_input = self.driver.find_element(By.XPATH, f"(//div[contains(@class,'_13NKt copyable-text selectable-text')])[{i+1}]")
                    option_input.click()
                    option_input.clear()
                    option_input.send_keys(self.poll_data[option][0])

                time.sleep(2)  # Wait for a moment before sending the poll
                actions = ActionChains(self.driver)
                actions.send_keys(Keys.ENTER).perform()
                time.sleep(2)  # Wait for the poll to send

            except TimeoutException as e:
                print("Failed to attach the poll within the given time.", e)
                traceback.print_exc()

    def send_message(self, message):
        actions = ActionChains(self.driver)
        message = message.replace("\n", '__new_line__').replace("\r", '__new_line__')
        msg_lines = [msg for msg in message.split('__new_line__') if msg.strip()]

        for msg in msg_lines:
            actions.send_keys(msg)
            actions.key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER)
            actions.key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER)

        actions.send_keys(Keys.ENTER).perform()
        time.sleep(2)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='WhatsApp Bulk Message Automation with optional Attachment feature')
    parser.add_argument('sheet_name', help='Google Sheet name', type=str)
    parser.add_argument('--image-path', help='Full path of image attachment', type=str, dest='image_path')
    parser.add_argument('--attach-poll', help='Google Sheet name for poll data', type=str, dest='poll_sheet_name')
    args = parser.parse_args()

    whatsapp = WhatsappMessage(sheet_name=args.sheet_name, image_path=args.image_path, poll_sheet_name=args.poll_sheet_name)
    whatsapp.start_process()
