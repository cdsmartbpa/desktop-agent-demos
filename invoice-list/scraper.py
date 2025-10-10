import json
import time
import sys
import subprocess
import importlib

# Auto-install missing packages
required_packages = {
    'selenium': 'selenium',
    'requests': 'requests'
}

for module, package in required_packages.items():
    try:
        importlib.import_module(module)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package, "--break-system-packages"])

import requests
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import datetime

# API Configuration - Update these values as needed
BASE_URL = "http://localhost:8080/"
TENANT = "Sandbox"
QUEUE = "P062 - Incomming invoices"
API_TOKEN = "JQ1mbu9UqHdRnlnIaIgi8xAmiZYfmOhRKleUiERbox5mRLhyev3JDF4IfLpKFrhL"

# Construct full API URL with proper encoding
ENCODED_QUEUE = urllib.parse.quote(QUEUE, safe='')
API_URL = f"{BASE_URL}admin/tenants/{TENANT}/queues/{ENCODED_QUEUE}/tasks/"

class InvoiceScraper:
    def __init__(self, url="https://cdsmartbpa.github.io/desktop-agent-demos/invoice-list/", demo_speed=1.0):
        self.url = url
        self.demo_mode = True  # Always in demo mode
        self.demo_speed = demo_speed  # Speed multiplier: 0.5=slower, 1.0=normal, 2.0=faster
        self.demo_delay = 1.0 / demo_speed  # Base delay adjusted by speed
        self.typing_delay = 0.05 / demo_speed  # Typing speed adjusted
        self.highlight_duration = 800 / demo_speed  # Highlight duration adjusted
        self.driver = None
        self.invoices_data = []
        
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
        """Login to the invoice system"""
        # Navigate to the page
        self.driver.get(self.url)
        time.sleep(self.demo_delay)
        
        # Wait for login form
        wait = WebDriverWait(self.driver, 10)
        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        
        # Highlight and fill username
        self.highlight_element(username_field)
        username_field.clear()
        self.type_slowly(username_field, "admin")
        time.sleep(self.demo_delay * 0.5)
        
        # Highlight and fill password
        password_field = self.driver.find_element(By.ID, "password")
        self.highlight_element(password_field)
        password_field.clear()
        self.type_slowly(password_field, "password123")
        time.sleep(self.demo_delay * 0.5)
        
        # Click login button
        login_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]")
        self.highlight_element(login_btn)
        time.sleep(self.demo_delay * 0.5)
        login_btn.click()
        
        # Wait for dashboard to load
        wait.until(EC.presence_of_element_located((By.ID, "invoicesTable")))
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
    
    def extract_invoice_data(self):
        """Extract data from invoice table and detailed supplier information"""
        # Find the invoices table
        table = self.driver.find_element(By.ID, "invoicesTable")
        self.highlight_element(table)
        time.sleep(self.demo_delay)
        
        # Get all data rows (skip header)
        rows = table.find_elements(By.XPATH, ".//tbody/tr")
        total_rows = len(rows)
        
        print(f"\n📊 Found {total_rows} invoices to process")
        print("=" * 50)
        
        # Process each row by index to avoid stale element issues
        for i in range(total_rows):
            print(f"\n🔍 Processing invoice {i+1}/{total_rows}")
            
            # Re-query the table and rows each iteration
            table = self.driver.find_element(By.ID, "invoicesTable")
            rows = table.find_elements(By.XPATH, ".//tbody/tr")
            row = rows[i]
            
            # Scroll to row and highlight it
            self.driver.execute_script("arguments[0].scrollIntoView(true);", row)
            time.sleep(self.demo_delay * 0.3)
            self.highlight_element(row)
            
            # Extract basic row data first
            basic_data = self.parse_invoice_row(row)
            
            if basic_data:
                # Click the row to view details using JavaScript
                time.sleep(self.demo_delay * 0.3)
                
                # Use JavaScript click to ensure the event fires properly
                self.driver.execute_script("arguments[0].click();", row)
                
                # Wait for details section to be visible (not just present)
                wait = WebDriverWait(self.driver, 10)
                wait.until(lambda d: d.find_element(By.ID, "supplierDetailsSection").is_displayed())
                time.sleep(self.demo_delay * 0.5)
                
                # Extract detailed supplier information
                detailed_data = self.extract_supplier_details()
                
                # Combine basic and detailed data
                complete_data = {**basic_data, **detailed_data}
                self.invoices_data.append(complete_data)
                
                print(f"✅ Extracted details for {complete_data['invoiceId']}")
                
                # Go back to invoice list using JavaScript click
                back_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Back to Invoices')]")
                self.highlight_element(back_btn)
                time.sleep(self.demo_delay * 0.3)
                self.driver.execute_script("arguments[0].click();", back_btn)
                
                # Wait for dashboard to be visible again
                wait.until(lambda d: d.find_element(By.ID, "dashboardSection").is_displayed())
                time.sleep(self.demo_delay * 0.3)
            
            time.sleep(self.demo_delay * 0.3)
    
    def parse_invoice_row(self, row):
        """Parse individual invoice row and extract basic data"""
        try:
            # Get all cells in the row
            cells = row.find_elements(By.TAG_NAME, "td")
            
            if len(cells) >= 5:
                invoice_id = cells[0].text.strip()
                date = cells[1].text.strip()
                vendor = cells[2].text.strip()
                amount_text = cells[3].text.strip()
                due_date = cells[4].text.strip()
                
                # Parse amount to determine state
                # Remove $ and commas, handle negative amounts
                amount_clean = amount_text.replace('$', '').replace(',', '').strip()
                amount_value = float(amount_clean)
                
                # Determine state based on amount
                state = "New Invoice" if amount_value >= 0 else "Creditnote"
                
                return {
                    "invoiceId": invoice_id,
                    "date": date,
                    "vendor": vendor,
                    "amount": amount_text,
                    "amountValue": amount_value,
                    "dueDate": due_date,
                    "state": state,
                    "extractedAt": datetime.now().isoformat()
                }
            
            return None
            
        except Exception as e:
            print(f"❌ Error parsing row: {str(e)}")
            return None
    
    def extract_supplier_details(self):
        """Extract detailed supplier information from the details page"""
        try:
            details = {}
            
            # Extract header information
            try:
                details['supplierName'] = self.driver.find_element(By.ID, "supplierName").text.strip()
            except:
                details['supplierName'] = ""
            
            try:
                details['supplierTagline'] = self.driver.find_element(By.ID, "supplierTagline").text.strip()
            except:
                details['supplierTagline'] = ""
            
            try:
                details['invoiceReference'] = self.driver.find_element(By.ID, "invoiceRef").text.strip()
            except:
                details['invoiceReference'] = ""
            
            # Company Information
            details['companyInfo'] = {
                'legalName': self.get_element_text('legalName'),
                'businessAddress': self.get_element_text('businessAddress'),
                'registrationNumber': self.get_element_text('registrationNumber')
            }
            
            # Contact Information
            details['contactInfo'] = {
                'contactPerson': self.get_element_text('contactPerson'),
                'email': self.get_element_text('email'),
                'phone': self.get_element_text('phone')
            }
            
            # Tax Information
            details['taxInfo'] = {
                'vatNumber': self.get_element_text('vatNumber'),
                'taxId': self.get_element_text('taxId'),
                'gstNumber': self.get_element_text('gstNumber')
            }
            
            # Banking Information
            details['bankingInfo'] = {
                'bankName': self.get_element_text('bankName'),
                'accountNumber': self.get_element_text('accountNumber'),
                'swiftCode': self.get_element_text('swiftCode'),
                'bankAddress': self.get_element_text('bankAddress')
            }
            
            # Payment Terms
            details['paymentTerms'] = self.get_element_text('paymentTerms')
            
            return details
            
        except Exception as e:
            print(f"❌ Error extracting supplier details: {str(e)}")
            return {}
    
    def get_element_text(self, element_id):
        """Safely get text from an element by ID"""
        try:
            element = self.driver.find_element(By.ID, element_id)
            text = element.text.strip()
            return text if text and text != '-' else ""
        except:
            return ""
    
    def send_all_invoices(self):
        """Send all extracted invoices to API"""
        print(f"\n🚀 Sending {len(self.invoices_data)} invoices to API...")
        print(f"🎯 Target: {API_URL}")
        print("=" * 50)
        
        success_count = 0
        fail_count = 0
        
        for invoice in self.invoices_data:
            if self.send_to_api(invoice):
                success_count += 1
            else:
                fail_count += 1
            time.sleep(self.demo_delay * 0.3)  # Small delay between API calls
        
        print("\n" + "=" * 50)
        print(f"✅ Successfully sent: {success_count}")
        print(f"❌ Failed: {fail_count}")
        print(f"📊 Total: {len(self.invoices_data)}")
    
    def send_to_api(self, invoice_data):
        """Send individual invoice data to API"""
        
        # Create a clean copy for API payload
        api_invoice_data = {
            "invoiceId": invoice_data.get('invoiceId', ''),
            "date": invoice_data.get('date', ''),
            "vendor": invoice_data.get('vendor', ''),
            "amount": invoice_data.get('amount', ''),
            "amountValue": invoice_data.get('amountValue', 0),
            "dueDate": invoice_data.get('dueDate', ''),
            "state": invoice_data.get('state', ''),
            "extractedAt": invoice_data.get('extractedAt', ''),
            "supplierDetails": {
                "supplierName": invoice_data.get('supplierName', ''),
                "supplierTagline": invoice_data.get('supplierTagline', ''),
                "invoiceReference": invoice_data.get('invoiceReference', ''),
                "companyInfo": invoice_data.get('companyInfo', {}),
                "contactInfo": invoice_data.get('contactInfo', {}),
                "taxInfo": invoice_data.get('taxInfo', {}),
                "bankingInfo": invoice_data.get('bankingInfo', {}),
                "paymentTerms": invoice_data.get('paymentTerms', '')
            }
        }
        
        # Prepare the API payload
        payload = {
            "name": invoice_data.get('invoiceId', 'Unknown'),
            "state": invoice_data.get('state', 'New Invoice'),
            "inputData": json.dumps(api_invoice_data, indent=2),
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
            
            # Print status
            print(f"\n📤 Sending {invoice_data.get('invoiceId', 'Unknown')} - State: {invoice_data.get('state', 'Unknown')}")
            print(f"📊 Status: {response.status_code}")
            if response.status_code not in [200, 201]:
                print(f"❌ Response: {response.text}")
            else:
                print(f"✅ Success!")
            print("-" * 50)
            
            return response.status_code == 200 or response.status_code == 201
        except Exception as e:
            print(f"❌ API Error for {invoice_data.get('invoiceId', 'Unknown')}: {str(e)}")
            print("-" * 50)
            return False
    
    def save_to_json(self, filename="invoices_data.json"):
        """Save extracted data to a JSON file"""
        try:
            output_path = f"/mnt/user-data/outputs/{filename}"
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(self.invoices_data, f, indent=2, ensure_ascii=False)
            print(f"\n💾 Data saved to: {output_path}")
            return output_path
        except Exception as e:
            print(f"❌ Error saving to JSON: {str(e)}")
            return None
    
    def run_demo(self):
        """Run the complete scraping demo"""
        try:            
            print("🚀 Starting Invoice Scraper Demo")
            print("=" * 50)
            
            self.setup_driver()
            time.sleep(self.demo_delay * 0.5)
            
            print("🔐 Logging in...")
            self.login()
            time.sleep(self.demo_delay * 0.5)
            
            print("\n📋 Extracting invoice data...")
            self.extract_invoice_data()
            time.sleep(self.demo_delay * 0.5)
            
            print("\n💾 Saving data to JSON...")
            json_path = self.save_to_json()
            
            print("\n📤 Sending to API...")
            self.send_all_invoices()
            
            print("\n✅ Demo completed successfully!")
            print("=" * 50)
            
            # Keep browser open for review
            input("\n⏸️  Press Enter to close the browser...")
            
        except Exception as e:
            print(f"❌ Error during demo: {str(e)}")
            import traceback
            traceback.print_exc()
            
        finally:
            if self.driver:
                self.driver.quit()

def main():
    """Main function to run the scraper"""
    # Adjust this value to control demo speed:
    # 0.5 = slower demo, 1.0 = normal speed, 2.0 = faster demo, 3.0 = very fast
    DEMO_SPEED = 2.0
    
    print("=" * 50)
    print("🎬 Invoice Detail Scraper")
    print("=" * 50)
    print(f"⚡ Demo Speed: {DEMO_SPEED}x")
    print("=" * 50)
    
    # Create and run scraper
    scraper = InvoiceScraper(demo_speed=DEMO_SPEED)
    scraper.run_demo()

if __name__ == "__main__":
    main()
