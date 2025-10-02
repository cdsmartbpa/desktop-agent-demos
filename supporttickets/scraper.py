import json
import time
import re
import requests
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import os
from datetime import datetime

# API Configuration - Update these values as needed
BASE_URL = "http://localhost:8080/"
TENANT = "Sandbox"
QUEUE = "P001 - Test"
API_TOKEN = "JQ1mbu9UqHdRnlnIaIgi8xAmiZYfmOhRKleUiERbox5mRLhyev3JDF4IfLpKFrhL"
STATE = "NEW"

# Construct full API URL with proper encoding
ENCODED_QUEUE = urllib.parse.quote(QUEUE, safe='')
API_URL = f"{BASE_URL}admin/tenants/{TENANT}/queues/{ENCODED_QUEUE}/tasks/"

class TicketScraper:
    def __init__(self, url="https://workdemos.z1.web.core.windows.net/supporttickets/", demo_speed=1.0):
        self.url = url
        self.demo_mode = True  # Always in demo mode
        self.demo_speed = demo_speed  # Speed multiplier: 0.5=slower, 1.0=normal, 2.0=faster
        self.demo_delay = 1.0 / demo_speed  # Base delay adjusted by speed
        self.typing_delay = 0.05 / demo_speed  # Typing speed adjusted
        self.highlight_duration = 800 / demo_speed  # Highlight duration adjusted
        self.driver = None
        self.tickets_data = []
        
    def setup_driver(self):
        """Setup Chrome driver with visible browser for demo"""
        chrome_options = Options()
        
        # Always visible browser for demo
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("detach", True)
        
        # Suppress Chrome logs and notifications for cleaner demo
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument("--disable-background-networking")
        chrome_options.add_argument("--disable-background-timer-throttling")
        chrome_options.add_argument("--disable-backgrounding-occluded-windows")
        chrome_options.add_argument("--disable-breakpad")
        chrome_options.add_argument("--disable-component-extensions-with-background-pages")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-features=TranslateUI")
        chrome_options.add_argument("--disable-ipc-flooding-protection")
        chrome_options.add_argument("--no-default-browser-check")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
            
        # Additional options for stability
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
    def login(self):
        """Login to the support system"""
        # Navigate to the page
        self.driver.get(self.url)
        time.sleep(self.demo_delay)
        
        # Wait for login form
        wait = WebDriverWait(self.driver, 10)
        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        
        # Highlight and fill username
        self.highlight_element(username_field)
        username_field.clear()
        self.type_slowly(username_field, "support")
        time.sleep(self.demo_delay * 0.5)
        
        # Highlight and fill password
        password_field = self.driver.find_element(By.ID, "password")
        self.highlight_element(password_field)
        password_field.clear()
        self.type_slowly(password_field, "helpdesk123")
        time.sleep(self.demo_delay * 0.5)
        
        # Click login button
        login_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Access Dashboard')]")
        self.highlight_element(login_btn)
        time.sleep(self.demo_delay * 0.5)
        login_btn.click()
        
        # Wait for dashboard to load
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ticket-card")))
        time.sleep(self.demo_delay)
        
    def highlight_element(self, element):
        """Highlight an element for demo purposes"""
        self.driver.execute_script(f"""
            arguments[0].style.border = '3px solid red';
            arguments[0].style.backgroundColor = 'yellow';
            setTimeout(function() {{
                arguments[0].style.border = '';
                arguments[0].style.backgroundColor = '';
            }}, {self.highlight_duration});
        """, element)
        time.sleep(self.highlight_duration / 2000)  # Wait for half the highlight duration
    
    def type_slowly(self, element, text):
        """Type text slowly for demo effect"""
        for char in text:
            element.send_keys(char)
            time.sleep(self.typing_delay)
    
    def extract_ticket_data(self):
        """Extract data from all visible tickets"""
        # Find all ticket cards
        ticket_cards = self.driver.find_elements(By.CLASS_NAME, "ticket-card")
        
        for i, card in enumerate(ticket_cards, 1):
            # Scroll to ticket and highlight it
            self.driver.execute_script("arguments[0].scrollIntoView(true);", card)
            time.sleep(self.demo_delay * 0.3)
            self.highlight_element(card)
            
            # Extract ticket data
            ticket_data = self.parse_ticket_card(card)
            
            if ticket_data:
                self.tickets_data.append(ticket_data)
            
            time.sleep(self.demo_delay * 0.5)
    
    def parse_ticket_card(self, card):
        """Parse individual ticket card and extract data"""
        try:
            # Extract ticket ID and headline
            ticket_id_elem = card.find_element(By.CLASS_NAME, "ticket-id")
            ticket_id = ticket_id_elem.text.strip()
            
            # Extract headline/title
            title_elem = card.find_element(By.CLASS_NAME, "ticket-title")
            headline = title_elem.text.strip()
            
            # Extract description/issue
            description_elem = card.find_element(By.CLASS_NAME, "ticket-description")
            description_text = description_elem.text.strip()
            
            # Parse description to separate issue text and error details
            issue_text, error_message, error_code = self.parse_description(description_text)
            
            # Create structured data
            ticket_data = {
                "TicketID": ticket_id,
                "Headline": headline,
                "Issue": issue_text,
                "Error message": error_message,
                "ErrorCode": error_code,
                "ExtractedAt": datetime.now().isoformat(),
                "RawDescription": description_text
            }
            
            return ticket_data
            
        except Exception as e:
            return None
    
    def parse_description(self, description):
        """Parse description to extract issue, error message, and error code"""
        # Initialize variables
        issue_text = description
        error_message = ""
        error_code = ""
        
        # Extract error code (look for patterns like "Error 806", "0x8004010F", etc.)
        error_code_patterns = [
            r'Error Code:\s*([A-Za-z0-9_]+)',
            r'Error (\d+)',
            r'(0x[A-Fa-f0-9]+)',
            r'STOP code:\s*([A-Za-z_]+)',
            r'Error:\s*"[^"]*Error (\d+)',
            r'([A-Z_]{3,})'  # Pattern for codes like DRIVER_IRQL_NOT_LESS_OR_EQUAL
        ]
        
        for pattern in error_code_patterns:
            match = re.search(pattern, description)
            if match:
                error_code = match.group(1)
                break
        
        # Extract error messages (text in quotes or after "Error:")
        error_msg_patterns = [
            r'Error[^:]*:\s*"([^"]+)"',
            r'Error message:\s*"([^"]+)"',
            r'showing error[^:]*:\s*"([^"]+)"',
            r'Error:\s*"([^"]+)"'
        ]
        
        for pattern in error_msg_patterns:
            match = re.search(pattern, description)
            if match:
                error_message = match.group(1)
                break
        
        # Clean up issue text by removing error code line if present
        if "Error Code:" in issue_text:
            issue_text = re.sub(r'\s*Error Code:.*', '', issue_text).strip()
        
        return issue_text, error_message, error_code
        
    def send_all_tickets(self):
        """Send all extracted tickets to API"""
        print(f"\n🚀 Sending {len(self.tickets_data)} tickets to API...")
        print(f"🎯 Target: {API_URL}")
        print("=" * 50)
        
        for ticket in self.tickets_data:
            self.send_to_api(ticket)
            time.sleep(self.demo_delay * 0.3)  # Small delay between API calls
        
        print("🏁 All tickets sent!")
    
    def send_to_api(self, ticket_data):
        """Send individual ticket data to API"""
        # Prepare the input data string
        input_data = f"Headline: {ticket_data['Headline']} Issue: {ticket_data['Issue']} Error message: {ticket_data['Error message']} ErrorCode: {ticket_data['ErrorCode']}"
        
        # Prepare the API payload
        payload = {
            "name": ticket_data['TicketID'],
            "state": STATE,
            "inputData": input_data,
            "workData": "",
            "enabled": True,
            "requireUniqueData": True,
            "tags": []
        }
        
        # Prepare headers
        headers = {
            'authorization': f'Bearer {API_TOKEN}',
            'content-type': 'application/json'
        }
        
        try:
            # Make the POST request
            response = requests.post(API_URL, json=payload, headers=headers, timeout=30)
            
            # Print status for debugging
            print(f"📤 Sending {ticket_data['TicketID']}")
            print(f"📍 URL: {API_URL}")
            print(f"📊 Status: {response.status_code}")
            if response.status_code not in [200, 201]:
                print(f"❌ Response: {response.text}")
            else:
                print(f"✅ Success!")
            print("-" * 50)
            
            return response.status_code == 200 or response.status_code == 201
        except Exception as e:
            print(f"❌ API Error for {ticket_data['TicketID']}: {str(e)}")
            print("-" * 50)
            return False
    
    def run_demo(self):
        """Run the complete scraping demo"""
        try:            
            self.setup_driver()
            time.sleep(self.demo_delay * 0.5)
            
            self.login()
            time.sleep(self.demo_delay * 0.5)
            
            self.extract_ticket_data()
            time.sleep(self.demo_delay * 0.5)
            
            self.send_all_tickets()
            
            # Keep browser open for review
            input("Press Enter to close the browser...")
            
        except Exception as e:
            pass
            
        finally:
            if self.driver:
                self.driver.quit()

def main():
    """Main function to run the scraper"""
    # Adjust this value to control demo speed:
    # 0.5 = slower demo, 1.0 = normal speed, 2.0 = faster demo, 3.0 = very fast
    DEMO_SPEED = 2.0
    
    # Create and run scraper
    scraper = TicketScraper(demo_speed=DEMO_SPEED)
    scraper.run_demo()

if __name__ == "__main__":
    main()