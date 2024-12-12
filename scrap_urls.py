from playwright.sync_api import sync_playwright
import csv
import os
import json
import time
import random
import pandas as pd

class PlaywrightDriver:
    def __init__(self, cookies_file=None):
        self.cookies_file = cookies_file
        self.playwright = None
        self.browser = None

    def initialize_driver(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=False,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--no-sandbox',
                '--disable-infobars',
                '--disable-dev-shm-usage',
                '--disable-extensions',
                '--disable-gpu',
            ]
        )
        context = self.browser.new_context()

        if self.cookies_file:
            with open(self.cookies_file, 'r') as file:
                cookies = json.load(file)
                context.add_cookies(cookies)

        return context

    def close(self, context):
        context.close()
        self.browser.close()
        self.playwright.stop()


def scrape_profiles(page_url, context):
    page = context.new_page()
    profiles = []

    try:
        page.goto(page_url)
        time.sleep(random.uniform(3, 5))  # Random delay to mimic human behavior

        # Updated locator to target profile URLs directly
        results = page.locator('a[href*="linkedin.com/in/"]')
        count = results.count()

        print(f"Found {count} profiles on {page_url}")

        for i in range(count):
            try:
                profile_link = results.nth(i).get_attribute('href')
                if profile_link:
                    profiles.append(profile_link.split('?')[0])  # Extract base URL
            except Exception as e:
                print(f"Error extracting profile link: {e}")

    except Exception as e:
        print(f"Error while scraping {page_url}: {e}")

    finally:
        page.close()

    return profiles


def scrape_urls(link, output_file, cookies_file=None, country=None):
    driver = PlaywrightDriver(cookies_file=cookies_file)
    context = driver.initialize_driver()

    try:
        all_profiles = []

        for page in range(1, 102):  # Loop through pages
            page_url = f"{link}&page={page}"
            print(f"Scraping page {page}...")

            profiles = scrape_profiles(page_url, context)
            if not profiles:
                print(f"No profiles found on page {page}. Stopping.")
                break

            all_profiles.extend(profiles)

        if not all_profiles:
            print("No profiles found for the entire link.")
            return

        # Check if the output file exists and if it has the 'Country' column
        file_exists = os.path.exists(output_file)

        if not file_exists:
            # Create a new file with header
            with open(output_file, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                header = ['Profile URL']
                if country:
                    header.append('Country')
                writer.writerow(header)

        # Read the existing data
        existing_data = pd.read_csv(output_file)
        existing_urls = existing_data['Profile URL'].tolist()

        with open(output_file, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)

            for profile in all_profiles:
                if profile in existing_urls:
                    # If country is provided, update the 'Country' column for existing entries
                    if country and 'Country' in existing_data.columns:
                        existing_data.loc[existing_data['Profile URL'] == profile, 'Country'] = country
                else:
                    # Add new entries
                    row = [profile]
                    if country:
                        row.append(country)
                    writer.writerow(row)

        # Save any updated 'Country' values back to the CSV
        if country and 'Country' in existing_data.columns:
            existing_data.to_csv(output_file, index=False)

        print(f"Scraping complete for {link}. Data saved to {output_file}.")

    finally:
        driver.close(context)


# Usage Example
output_file = "event_attendees.csv"
cookies_file = "cookies.json"

#scrape_general(
#    link="https://www.linkedin.com/search/results/people/?eventAttending=%5B%227270037764836913153%22%5D&origin=EVENT_PAGE_CANNED_SEARCH",
#    output_file=output_file,
#    cookies_file=cookies_file
#)

scrape_urls(
    link="https://www.linkedin.com/search/results/people/?eventAttending=%5B%227270037764836913153%22%5D&geoUrn=%5B%22101165590%22%2C%22103644278%22%2C%22105646813%22%2C%22106155005%22%2C%22100565514%22%2C%22105490917%22%2C%22103619019%22%2C%22104170880%22%2C%22106057199%22%5D&origin=FACETED_SEARCH&sid=Fs%3A",
    output_file=output_file,
    cookies_file=cookies_file,
    #country="India"
)
