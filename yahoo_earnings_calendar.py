# This script uses Selenium to fetch data from Yahoo Finance.
# You need to have Google Chrome installed.
# You need to download ChromeDriver and ensure it's in your system's PATH or specify its executable path.
# Install selenium: pip install selenium
# If you encounter a WebDriverException, it likely means ChromeDriver is not correctly set up or found.

# See README.md for detailed setup and usage instructions.

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, TimeoutException
import sys
import json
import time

def get_earnings_calendar(ticker):
    """
    Fetches and prints the earnings calendar data for a given stock ticker using Selenium.
    """
    url = f"https://finance.yahoo.com/calendar/earnings?symbol={ticker}"
    
    driver = None # Initialize driver to None for the finally block
    try:
        print("Initializing WebDriver...")
        # Ensure chromedriver is in your PATH or specify executable_path
        # e.g., driver = webdriver.Chrome(service=webdriver.chrome.service.Service(executable_path='/path/to/chromedriver'))
        try:
            driver = webdriver.Chrome() 
        except WebDriverException as e:
            print(f"WebDriverException: Chromedriver executable may not be in PATH or configured correctly. Please check setup. Detail: {e}")
            return
        print("WebDriver initialized.")

        print(f"Fetching page: {url}")
        driver.get(url)
        
        # Wait for a specific element that indicates page data (especially scripts) has likely loaded.
        # Waiting for a script tag containing 'root.App.main' is a proxy for this.
        wait_time = 20 # seconds
        try:
            print(f"Waiting up to {wait_time} seconds for page data to fully load (signaled by presence of 'root.App.main' script)...")
            WebDriverWait(driver, wait_time).until(
                EC.presence_of_element_located((By.XPATH, "//script[contains(text(), 'root.App.main')]"))
            )
            print("Page data appears to be loaded (marker found).")
        except TimeoutException:
            print(f"Timed out waiting for page data marker after {wait_time} seconds. Proceeding with current page source.")
            # Optionally, handle this more gracefully, e.g., by exiting or trying a different strategy

        html_content = driver.page_source
        print("Page source retrieved.")
        print(f"First 1000 characters of page source:\n{html_content[:1000]}")

        # JSON extraction logic
        print("\nExtracting data...")
        DATA_START_MARKER = "root.App.main = "
        DATA_END_MARKER = "};" # For Method 1
        json_string_to_parse = None # Will hold the string to be parsed

        start_marker_index = html_content.find(DATA_START_MARKER)

        if start_marker_index != -1:
            json_start_index = start_marker_index + len(DATA_START_MARKER)
            end_marker_index = html_content.find(DATA_END_MARKER, json_start_index)

            if end_marker_index != -1:
                json_string_to_parse = html_content[json_start_index : end_marker_index + 1].strip()
                print(f"Method 1: Found '{DATA_START_MARKER}' and '{DATA_END_MARKER}'.")
            else:
                print(f"Method 1: Found '{DATA_START_MARKER}' but could not find subsequent '{DATA_END_MARKER}'.")
        else:
            print(f"Method 1: Could not find '{DATA_START_MARKER}' in the page source.")

        # JSON parsing logic (and Method 2 if Method 1 failed to extract)
        if json_string_to_parse:
            print("\nParsing JSON (from Method 1)...")
            try:
                parsed_json = json.loads(json_string_to_parse)
                print("Successfully parsed JSON string from Method 1.")
            except json.JSONDecodeError as e:
                print(f"Error: Failed to decode JSON from Method 1. Detail: {e}")
                print(f"Problematic JSON string snippet (first 300 chars): {json_string_to_parse[:300]}")
                parsed_json = None # Ensure it's None if parsing fails
        else:
            print("Skipping parsing for Method 1 as no JSON string was extracted.")

        # Method 2: Fallback to finding a generic <script type="application/json">
        if parsed_json is None:
            print("\nAttempting Method 2: Searching for the first generic <script type=\"application/json\">...")
            script_tag_marker_start = '<script type="application/json">' 
            script_tag_pos_start = html_content.find(script_tag_marker_start)

            if script_tag_pos_start != -1:
                json_script_start = script_tag_pos_start + len(script_tag_marker_start)
                script_tag_marker_end = '</script>'
                script_tag_pos_end = html_content.find(script_tag_marker_end, json_script_start)

                if script_tag_pos_end != -1:
                    json_string_to_parse = html_content[json_script_start:script_tag_pos_end].strip()
                    print("Method 2: Found potential JSON content in a generic <script> tag.")
                    print(f"Extracted content (first 500 chars for Method 2):\n{json_string_to_parse[:500]}")
                    print("\nParsing JSON (from Method 2)...")
                    try:
                        if json_string_to_parse.startswith('{') and json_string_to_parse.endswith('}'):
                            parsed_json = json.loads(json_string_to_parse)
                            print("Successfully parsed JSON string from Method 2.")
                        else:
                            print("Method 2: Content of <script type=\"application/json\"> does not appear to be a valid JSON object (doesn't start/end with {}).")
                            parsed_json = None
                    except json.JSONDecodeError as e:
                        print(f"Method 2: Error parsing JSON from <script type=\"application/json\">: {e}")
                        parsed_json = None
                else:
                    print("Method 2: Could not find closing </script> tag for generic <script type=\"application/json\">.")
            else:
                print("Method 2: Could not find any generic <script type=\"application/json\"> tag.")
        
        # Process and print the successfully parsed JSON
        if parsed_json:
            print("\n--- Parsed JSON Data ---")
            # Exploratory data access as path is uncertain
            earnings_data = parsed_json.get("context", {}).get("dispatcher", {}).get("stores", {}).get("QuoteSummaryStore", {}).get("calendarEvents", {}).get("earnings", {}).get("earningsDate", [])
            
            if not earnings_data:
                print("Primary earnings data path ('QuoteSummaryStore') not found or empty.")
                print("Trying alternative path: 'StreamDataStore.results'")
                stream_data_store = parsed_json.get("context", {}).get("dispatcher", {}).get("stores", {}).get("StreamDataStore", {})
                if isinstance(stream_data_store, dict):
                    earnings_data_alternative = stream_data_store.get("results", [])
                    if isinstance(earnings_data_alternative, list):
                        earnings_data = earnings_data_alternative 
                        print("Data found in 'StreamDataStore.results'.")
                    else:
                        print("Alternative path 'StreamDataStore.results' did not yield a list.")
                        earnings_data = []
                else:
                    print("Alternative path 'StreamDataStore' not found or not a dictionary.")
                    earnings_data = []

                if earnings_data:
                    print("\nEarnings Data Found:")
                    print(json.dumps(earnings_data, indent=4))
                else:
                    print("\nCould not find specific earnings data using common exploratory paths.")
                    print("Printing full JSON structure (first 2000 chars) for manual inspection:")
                    print(json.dumps(parsed_json, indent=4)[:2000])
            
            except json.JSONDecodeError as e:
                print(f"\nError: Failed to decode JSON. Detail: {e}")
                print(f"Problematic JSON string snippet (first 300 chars): {json_string[:300]}")
        else:
            print("\nNo JSON string was extracted, skipping parsing.")

    except WebDriverException as e:
        print(f"WebDriverException: Chromedriver executable may not be in PATH or configured correctly. Please check setup. Detail: {e}")
    except Exception as e: 
        print(f"An unexpected error occurred: {e}")

                # Inspect and Extract Earnings Data
                earnings_data = parsed_json.get("context", {}).get("dispatcher", {}).get("stores", {}).get("QuoteSummaryStore", {}).get("calendarEvents", {}).get("earnings", {}).get("earningsDate", [])
                
                if not earnings_data: # Try another common path if the first one fails
                    # This path is a general guess, actual structure might vary.
                    # Often, earnings might be part of a more generic event store or stream.
                    # For example, if results are a list of events, we might need to filter them.
                    print("First earnings data path not found, trying alternative: context.dispatcher.stores.StreamDataStore.results")
                    stream_data_store = parsed_json.get("context", {}).get("dispatcher", {}).get("stores", {}).get("StreamDataStore", {})
                    if isinstance(stream_data_store, dict): #Check if stream_data_store is a dictionary
                         # Attempt to access 'results' if stream_data_store is a dictionary
                        earnings_data_alternative = stream_data_store.get("results", [])
                        if isinstance(earnings_data_alternative, list): # Ensure it's a list
                            # Further filtering might be needed here if 'results' contains various event types
                            earnings_data = earnings_data_alternative 
                        else: # If results is not a list, it's not what we expect for earnings_data
                            print("Alternative path 'StreamDataStore.results' did not yield a list.")
                            earnings_data = [] # Reset to empty list
                    else: # If stream_data_store is not a dictionary
                        print("Alternative path 'StreamDataStore' not found or not a dictionary.")
                        earnings_data = []


                if earnings_data:
                    print("\nAttempting to print extracted earnings data:")
                    print(json.dumps(earnings_data, indent=4))
                else:
                    print("\nCould not find specific earnings data path using common patterns.")
                    print("Printing full JSON structure (first 2000 chars) for inspection:")
                    print(json.dumps(parsed_json, indent=4)[:2000])
            
            except json.JSONDecodeError as e:
                print(f"\nError: Failed to decode JSON. Detail: {e}")
                print(f"Problematic JSON string snippet (first 300 chars): {json_string[:300]}")
        else:
            print("\nNo JSON string was extracted, skipping parsing.")

    except Exception as e: # More general exception to catch WebDriver-related issues too
        print(f"An error occurred: {e}")

    finally:
        if driver:
            driver.quit()
            print("WebDriver closed.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python yahoo_earnings_calendar.py <TICKER>")
        sys.exit(1)

    ticker_symbol = sys.argv[1]
    get_earnings_calendar(ticker_symbol)
