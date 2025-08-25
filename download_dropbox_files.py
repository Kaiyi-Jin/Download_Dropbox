import pandas as pd
import time
import os
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import logging

class DropboxDownloader:
    def __init__(self, keyword, download_folder="downloads", delay_between_searches=3):
        """
        Initialize the Dropbox downloader for a single keyword
        
        Args:
            keyword (str): Single keyword to search for
            download_folder (str): Folder to save downloaded files
            delay_between_searches (int): Delay in seconds between searches to avoid rate limiting
        """
        self.keyword = keyword
        self.download_folder = os.path.abspath(download_folder)
        self.delay = delay_between_searches
        self.driver = None
        self.wait = None
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('dropbox_download.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # Create download folder if it doesn't exist
        os.makedirs(self.download_folder, exist_ok=True)
        
        self.logger.info(f"Initialized with keyword: {keyword}")
    
    def setup_driver(self):
        """Setup Chrome driver with download preferences"""
        chrome_options = Options()
        
        # Set download preferences
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Optional: Run in background (uncomment next line if you don't want to see browser)
        # chrome_options.add_argument("--headless")
        
        # Additional options for stability
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-notifications")  # Disable notifications to avoid DEPRECATED_ENDPOINT error
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 10)
            self.logger.info("Chrome driver initialized successfully")
        except Exception as e:
            self.logger.error(f"Error setting up Chrome driver: {e}")
            self.logger.error("Make sure ChromeDriver is installed and in PATH")
            raise
    
    def login_to_dropbox(self):
        """Navigate to Dropbox and wait for user to complete SSO login"""
        try:
            self.driver.get("https://www.dropbox.com/login")
            self.logger.info("Navigated to Dropbox login page")
            
            print("\n" + "="*60)
            print("MANUAL LOGIN REQUIRED")
            print("="*60)
            print("1. Complete the SSO login process in the browser window")
            print("2. Make sure you reach the main Dropbox interface")
            print("3. Press ENTER in this console when login is complete...")
            print("="*60)
            
            input("Press ENTER after completing login...")
            
            # Wait for main Dropbox interface to load
            try:
                self.wait.until(
                    EC.any_of(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='search-input']")),
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='Search']")),
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".search-input"))
                    )
                )
                self.logger.info("Successfully logged in to Dropbox")
                return True
            except TimeoutException:
                self.logger.error("Could not find search input. Please make sure you're on the main Dropbox page")
                return False
                
        except Exception as e:
            self.logger.error(f"Error during login process: {e}")
            return False
    
    def clear_search_context(self):
        """Clear the current search context"""
        try:
            # Try multiple methods to clear search context
            clear_selectors = [
                "[data-testid='search-clear-button']",
                ".search-clear-button",
                "button[title*='Clear']",
                "button[aria-label*='Clear']"
            ]
            
            for selector in clear_selectors:
                try:
                    clear_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    clear_button.click()
                    self.logger.info("Cleared search context")
                    time.sleep(1)
                    return True
                except:
                    continue
            
            # Fallback: try clearing through search input
            search_input = self.get_search_input()
            if search_input:
                search_input.clear()
                search_input.send_keys(Keys.ESCAPE)
                self.logger.info("Cleared search context via search input")
                time.sleep(1)
                return True
                
            self.logger.warning("Could not find clear button or search input to clear context")
            return False
            
        except Exception as e:
            self.logger.warning(f"Error clearing search context: {e}")
            return False
    
    def get_search_input(self):
        """Helper method to find search input"""
        search_selectors = [
            "[data-testid='search-input']",
            "input[placeholder*='Search']",
            ".search-input",
            "input[type='search']",
            "#search-input"
        ]
        
        for selector in search_selectors:
            try:
                search_input = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                return search_input
            except:
                continue
        return None
    
    def sanitize_filename(self, filename):
        """Sanitize filename and apply custom renaming for the target .cwa file"""
        if not filename:
            return None
        # Custom renaming for the specific file
        if filename.lower() == 'file, f1-00094_76399_0000000000_ssr.cwa':
            return 'f1-00094_76399_0000000000_ssr.cwa'
        # General sanitization for other files
        return re.sub(r'[<>:"/\\|?*,]', '_', filename).strip()
    
    def verify_download(self, filename, timeout=30):
        """Verify if the file was downloaded to the specified folder"""
        sanitized_filename = self.sanitize_filename(filename)
        if not sanitized_filename:
            return False
        
        expected_path = os.path.join(self.download_folder, sanitized_filename)
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            if os.path.exists(expected_path):
                self.logger.info(f"Download verified: {expected_path}")
                return True
            # Check for partial downloads (Chrome adds .crdownload extension)
            if os.path.exists(expected_path + '.crdownload'):
                self.logger.info(f"Download in progress: {expected_path}.crdownload")
            time.sleep(1)
        
        self.logger.warning(f"Download not found: {expected_path}")
        return False
    
    def attempt_download(self, file_element, file_name, index):
        """Attempt to download a file with retries"""
        max_attempts = 2
        sanitized_file_name = self.sanitize_filename(file_name)
        
        for attempt in range(1, max_attempts + 1):
            try:
                # Right-click to open context menu
                webdriver.ActionChains(self.driver).context_click(file_element).perform()
                time.sleep(1)
                
                # Look for download option in context menu
                download_selectors = [
                    "[data-testid='download-menu-item']",
                    "span:contains('Download')",
                    ".download-option",
                    "*[title*='Download']",
                    "*[aria-label*='Download']",
                    "button[class*='download']",
                    "[role='menuitem'][data-action*='download']"
                ]
                
                for download_selector in download_selectors:
                    try:
                        if "contains" in download_selector:
                            download_option = self.driver.find_element(By.XPATH, f"//span[contains(text(), 'Download')]")
                        else:
                            download_option = self.driver.find_element(By.CSS_SELECTOR, download_selector)
                        
                        # Try JavaScript click as fallback
                        try:
                            download_option.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", download_option)
                        
                        self.logger.info(f"Initiated download for .cwa file {index+1} ({sanitized_file_name}) on attempt {attempt}")
                        # Verify download
                        if self.verify_download(file_name):
                            return True
                        else:
                            self.logger.warning(f"Download verification failed for {sanitized_file_name} on attempt {attempt}")
                    except:
                        continue
                
                # Try alternative: double-click to open file, then look for download button
                try:
                    webdriver.ActionChains(self.driver).double_click(file_element).perform()
                    time.sleep(2)
                    
                    download_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Download') or @title='Download' or @aria-label='Download']")
                    try:
                        download_btn.click()
                    except:
                        self.driver.execute_script("arguments[0].click();", download_btn)
                    
                    self.logger.info(f"Downloaded .cwa file {index+1} ({sanitized_file_name}) via file view on attempt {attempt}")
                    if self.verify_download(file_name):
                        self.driver.back()
                        time.sleep(2)
                        return True
                    else:
                        self.logger.warning(f"Download verification failed for {sanitized_file_name} via file view on attempt {attempt}")
                    
                    self.driver.back()
                    time.sleep(2)
                except:
                    self.logger.warning(f"Could not download via file view on attempt {attempt}")
                
            except Exception as e:
                self.logger.warning(f"Download attempt {attempt} failed for file {index+1} ({sanitized_file_name}): {e}")
            
            # Close any context menus
            try:
                self.driver.find_element(By.TAG_NAME, "body").click()
                time.sleep(1)
            except:
                pass
            
            if attempt < max_attempts:
                self.logger.info(f"Retrying download for file {index+1} ({sanitized_file_name})...")
                time.sleep(2)
        
        self.logger.warning(f"Could not download .cwa file {index+1} ({sanitized_file_name}) after {max_attempts} attempts")
        return False
    
    def search_and_download(self):
        """Search for the keyword and download matching .cwa files"""
        try:
            # Clear previous search context
            self.clear_search_context()
            
            # Find and use search input
            search_input = self.get_search_input()
            if not search_input:
                self.logger.error("Could not find search input field")
                return False
            
            # Enter the keyword
            search_input.send_keys(self.keyword)
            search_input.send_keys(Keys.RETURN)
            
            self.logger.info(f"Searching for: {self.keyword}")
            
            # Wait for search results to load
            time.sleep(5)
            
            # Look for file results
            file_selectors = [
                "[data-testid='virtual-list-item']",
                ".file-row",
                ".brws-file-name-cell",
                "[role='row']",
                ".sl-react-shared-file-row"
            ]
            
            files_found = []
            for selector in file_selectors:
                try:
                    files_found = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if files_found:
                        self.logger.info(f"Found files using selector: {selector}")
                        break
                except:
                    continue
            
            if not files_found:
                self.logger.warning(f"No files found for keyword: {self.keyword}")
                return True
            
            self.logger.info(f"Found {len(files_found)} potential results for: {self.keyword}")
            
            # Process each file
            downloaded_count = 0
            for i, file_element in enumerate(files_found[:5]):  # Limit to first 5 results
                try:
                    # Try multiple selectors to find file name
                    file_name = None
                    name_selectors = [
                        ".brws-file-name-cell-filename",
                        ".file-name",
                        "span[data-testid='file-name']",
                        "[data-testid='file-name-text']",
                        ".mc-media-cell-text",
                        "span[class*='file-name']",
                        ".brws-file-name",
                        "[class*='filename']",
                        ".mc-media-row-main-content",
                        "[data-testid*='file-row'] span"
                    ]
                    
                    for name_selector in name_selectors:
                        try:
                            file_name_element = file_element.find_element(By.CSS_SELECTOR, name_selector)
                            file_name = file_name_element.text.strip().lower()
                            self.logger.info(f"Found file name '{file_name}' using selector: {name_selector}")
                            break
                        except:
                            continue
                    
                    # Fallback: Try getting attributes
                    if not file_name:
                        try:
                            for attr in ["title", "aria-label", "data-filename", "data-testid"]:
                                file_name = file_element.get_attribute(attr)
                                if file_name:
                                    file_name = file_name.strip().lower()
                                    self.logger.info(f"Found file name '{file_name}' using attribute: {attr}")
                                    break
                        except:
                            pass
                    
                    if not file_name:
                        self.logger.warning(f"Could not determine file name for file {i+1}")
                        try:
                            element_html = file_element.get_attribute("outerHTML")[:200]
                            self.logger.debug(f"File element HTML: {element_html}")
                        except:
                            self.logger.debug("Could not retrieve file element HTML")
                        continue
                    
                    # Check if file name ends with .cwa
                    if not file_name.endswith('.cwa'):
                        self.logger.info(f"Skipping file {file_name} - does not end with .cwa")
                        continue
                        
                    # Attempt to download the file
                    if self.attempt_download(file_element, file_name, i):
                        downloaded_count += 1
                
                except Exception as e:
                    self.logger.warning(f"Error processing file {i+1} for keyword '{self.keyword}': {e}")
                    try:
                        element_html = file_element.get_attribute("outerHTML")[:200]
                        self.logger.debug(f"File element HTML: {element_html}")
                    except:
                        self.logger.debug("Could not retrieve file element HTML")
                    continue
            
            self.logger.info(f"Completed search for '{self.keyword}': {downloaded_count} .cwa downloads initiated")
            return True
            
        except Exception as e:
            self.logger.error(f"Error searching for keyword '{self.keyword}': {e}")
            return False
def cleanup(self):
    """Cleanup resources"""
    if self.driver:
        self.driver.quit()
        self.logger.info("Browser closed and resources cleaned up")

def run(self):
    """Main execution method"""
    try:
        self.logger.info("Starting search for keyword: {self.keyword}")
        
        # Process the keyword
        if self.search_and_download():
            self.logger.info("Keyword processed successfully")
        else:
            self.logger.error("Keyword processing failed")
        
    except KeyboardInterrupt:
        self.logger.info("Process interrupted by user")
    except Exception as e:
        self.logger.error(f"Unexpected error: {e}")

# Modify the main execution block
if __name__ == "__main__":
    # Configuration
    DOWNLOAD_FOLDER = r"E:\dropbox"  # Where to save files
    DELAY_BETWEEN_SEARCHES = 3  # Seconds between searches
    
    try:
        # Read keywords from CSV file
        dropbox_files = pd.read_csv(r'C:\Users\kaiyijin\NUS Dropbox\Kaiyi Jin\COBRA_Kaiyi\ID_Mapping\dropbox_files.csv') 
        keywords = dropbox_files['filename'].tolist()
        
        # Initialize downloader with first keyword
        if keywords:
            first_downloader = DropboxDownloader(
                keyword=keywords[0],
                download_folder=DOWNLOAD_FOLDER,
                delay_between_searches=DELAY_BETWEEN_SEARCHES
            )
            
            try:
                # Setup browser and login once
                first_downloader.setup_driver()
                if not first_downloader.login_to_dropbox():
                    raise Exception("Failed to login to Dropbox")
                
                # Process all keywords using the same browser session
                for keyword in keywords:
                    try:
                        # Update keyword and reuse the driver
                        first_downloader.keyword = keyword
                        
                        # Process keyword
                        if first_downloader.search_and_download():
                            first_downloader.logger.info(f"Successfully processed keyword: {keyword}")
                        else:
                            first_downloader.logger.error(f"Failed to process keyword: {keyword}")
                        
                        # Wait between searches
                        time.sleep(DELAY_BETWEEN_SEARCHES)
                        
                    except Exception as e:
                        first_downloader.logger.error(f"Error processing keyword '{keyword}': {e}")
                        continue
                        
            finally:
                # Cleanup only after all keywords are processed
                first_downloader.cleanup()
                
    except Exception as e:
        print(f"Error: {e}")
        print("\nMake sure you have:")
        print("1. Installed the required packages: pip install selenium pandas")
        print("2. ChromeDriver is installed and in your PATH")
        print("3. Provided the correct CSV file path with a 'filename' column")