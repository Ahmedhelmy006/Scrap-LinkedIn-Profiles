from playwright.sync_api import sync_playwright
import json
import time
import csv

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


def scrape_likes(link, cookies_file, output_file):
    driver = PlaywrightDriver(cookies_file=cookies_file)
    context = driver.initialize_driver()

    try:
        page = context.new_page()
        page.goto(link)
        time.sleep(3)  

        # Click on the "Likes" button
        #like_button = page.locator('span.social-details-social-counts__reactions-count')
        like_button = page.locator('span.social-details-social-counts__social-proof-text')
        like_button.click()
        time.sleep(5)  # Wait for the likes modal to open

        # Scroll through the modal to load all profiles
        modal = page.locator('div.artdeco-modal__content')
        all_profile_urls = set()  # Use a set to store unique profile URLs
        previous_profile_count = -1
        scroll_attempts = 0
        max_scroll_attempts = 50  # Limit to 50 scrolls

        while scroll_attempts < max_scroll_attempts:
            # Extract all loaded profiles at this stage
            profiles = modal.locator('a[href*="linkedin.com/in/"]')
            current_profiles = profiles.count()

            # Add new profiles to the set
            for i in range(current_profiles):
                profile_url = profiles.nth(i).get_attribute('href').split('?')[0]
                all_profile_urls.add(profile_url)

            # Stop if no new profiles are loaded
            if current_profiles == previous_profile_count:
                print("All profiles loaded.")
                break
            previous_profile_count = current_profiles

            # Scroll down to load more profiles
            modal.evaluate("(el) => el.scrollBy(0, el.scrollHeight)")  # Scroll down fully
            time.sleep(2)  # Adjust this to control speed (increase for slower, decrease for faster)
            scroll_attempts += 1

        print(f"Scrolling complete. Total scroll attempts: {scroll_attempts}")

        # Save all unique profiles to the CSV file
        with open(output_file, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Profile URL'])  # Write the header if the file is new
            for profile in all_profile_urls:
                writer.writerow([profile])

        print(f"Scraped {len(all_profile_urls)} unique profiles from likes.")

    finally:
        driver.close(context)



# Usage Example
cookies_file = r"Input Files\cookies.json"
output_file = "likes_profiles_second_post.xlsx"
linkedin_post_url = "https://www.linkedin.com/events/7270037764836913153/comments/"  
scrape_likes(link=linkedin_post_url, cookies_file=cookies_file, output_file=output_file)
