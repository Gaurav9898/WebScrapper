import time
import re
import pandas as pd
import random
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ----------------------
# SCROLL FUNCTION
# ----------------------
def scroll_until_end(browser, timeout=60):
    start_time = time.time()
    try:
        result_div = browser.find_element(By.XPATH, '//div[@role="feed"]')
    except NoSuchElementException:
        print("⚠ No feed div found, skipping.")
        return 'no_feed'

    while time.time() - start_time < timeout:
        current_height = browser.execute_script("return arguments[0].scrollHeight", result_div)
        browser.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", result_div)
        time.sleep(random.uniform(10, 20))
        new_height = browser.execute_script("return arguments[0].scrollHeight", result_div)

        if new_height == current_height:
            return 'done'
    return 'timeout'


# ----------------------
# SCRAPE BUSINESS DETAILS
# ----------------------
def scrape_business_details(browser, url):
    browser.get(url)
    time.sleep(random.uniform(10, 20))  # wait for JS to load
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    result = {"Listing URL": url}

    try:
        company_el = soup.find(name='h1', class_=re.compile('DUwDvf'))
        result['Company Name'] = company_el.get_text(strip=True) if company_el else None
    except:
        result['Company Name'] = None

    try:
        tag_el = soup.find(name='button', class_=re.compile('DkEaL'))
        result['Tag'] = tag_el.get_text(strip=True) if tag_el else None
    except:
        result['Tag'] = None

    try:
        address_el = soup.find(name='button', attrs={'data-item-id': 'address'})
        result['Full Address'] = address_el.get_text(strip=True) if address_el else None
    except:
        result['Full Address'] = None

    try:
        phone_el = soup.find(name='button', attrs={'data-item-id': re.compile('phone')})
        result['Phone Number'] = phone_el.get_text(strip=True) if phone_el else None
    except:
        result['Phone Number'] = None

    try:
        website_el = soup.find(name='a', attrs={'data-item-id': re.compile('authority')})
        result['Website'] = website_el['href'] if website_el else None
    except:
        result['Website'] = None

    return result


# ----------------------
# CHECK FOR CAPTCHA
# ----------------------
def check_captcha(browser):
    page_source = browser.page_source.lower()
    if "unusual traffic" in page_source or "verify you are human" in page_source:
        print("⚠ CAPTCHA detected! Solve it in the browser window...")
        while True:
            if "searchboxinput" in browser.page_source:
                print("✅ Captcha solved. Continuing...")
                break
            time.sleep(random.uniform(10, 20))


# ----------------------
# EXPLORE CITY: ZOOM IN + PAN + ZOOM OUT
# ----------------------
def explore_city(browser, term, zoom_in_levels=2, moves_per_level=4):
    """
    Explore a city by zooming in, panning around, then zooming back out.
    zoom_in_levels: how many times to zoom in
    moves_per_level: how many pans per zoom level
    """
    collected_links = set()
    directions = [Keys.ARROW_RIGHT, Keys.ARROW_DOWN, Keys.ARROW_LEFT, Keys.ARROW_UP]

    # 🔹 Phase 1: Zoom In & Explore
    for zoom in range(zoom_in_levels):
        print(f"🔍 Zooming in {zoom+1}/{zoom_in_levels}")
        browser.find_element(By.TAG_NAME, "body").send_keys(Keys.ADD)  # '+' zoom in
        time.sleep(random.uniform(5, 8))

        for move in range(moves_per_level):
            print(f"   → Move {move+1}/{moves_per_level} at zoom {zoom+1}")
            scroll_until_end(browser, timeout=20)

            listings = browser.find_elements(By.XPATH, '//a[contains(@href, "/maps/place/")]')
            for anchor in listings:
                link = anchor.get_attribute('href')
                if link and '/maps/place/' in link:
                    collected_links.add((term, link))

            direction = directions[move % len(directions)]
            browser.find_element(By.TAG_NAME, "body").send_keys(direction)
            time.sleep(random.uniform(5, 8))

    # 🔹 Phase 2: Zoom Out & Explore
    for zoom in range(zoom_in_levels):
        print(f"🔍 Zooming out {zoom+1}/{zoom_in_levels}")
        browser.find_element(By.TAG_NAME, "body").send_keys(Keys.SUBTRACT)  # '-' zoom out
        time.sleep(random.uniform(5, 8))

        for move in range(moves_per_level):
            print(f"   → Move {move+1}/{moves_per_level} at zoom-out step {zoom+1}")
            scroll_until_end(browser, timeout=20)

            listings = browser.find_elements(By.XPATH, '//a[contains(@href, "/maps/place/")]')
            for anchor in listings:
                link = anchor.get_attribute('href')
                if link and '/maps/place/' in link:
                    collected_links.add((term, link))

            direction = directions[move % len(directions)]
            browser.find_element(By.TAG_NAME, "body").send_keys(direction)
            time.sleep(random.uniform(5, 8))

    return collected_links


# ----------------------
# LAYER 1: Collect URLs (with map exploration & zooming)
# ----------------------
def layer1_collect_urls(input_file='WRE Layer 1 Search String.xlsx', sheet='Sheet3', output_file='Layer1_Output_URLs.xlsx'):
    search_terms = pd.read_excel(input_file, sheet_name=sheet)['URL'].dropna().tolist()
    search_terms = [term + " Rajasthan" for term in search_terms]

    chrome_options = Options()
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    google_url = 'https://www.google.com/maps/'
    all_results = []

    for idx, term in enumerate(search_terms, 1):
        print(f"\n[L1 {idx}/{len(search_terms)}] Searching: {term}")
        browser.get(google_url)
        time.sleep(8)
        check_captcha(browser)

        try:
            text_box = WebDriverWait(browser, 15).until(
                EC.element_to_be_clickable((By.ID, "searchboxinput"))
            )
            text_box.clear()
            text_box.send_keys(term)
            text_box.send_keys(Keys.RETURN)
            time.sleep(random.uniform(10, 20))
            check_captcha(browser)

            # Explore city streets deeper with zooming
            city_links = explore_city(browser, term, zoom_in_levels=2, moves_per_level=4)
            for search_term, link in city_links:
                all_results.append({"Search Term": search_term, "Listing URL": link})

            # Auto-save every term
            pd.DataFrame(all_results).drop_duplicates().to_excel(output_file, index=False)
            print(f"   → {len(all_results)} total URLs saved so far.")

        except Exception as e:
            print(f"❌ Error with {term}: {e}")

    browser.quit()
    print(f"\n✅ LAYER 1 finished. Saved to {output_file}")


# ----------------------
# LAYER 2: Enrich Details
# ----------------------
def layer2_scrape_details(input_file='Layer1_Output_URLs.xlsx', output_file='Layer2_Output_Details.xlsx'):
    urls_df = pd.read_excel(input_file).drop_duplicates(subset=["Listing URL"])
    urls = urls_df['Listing URL'].tolist()

    chrome_options = Options()
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    all_details = []

    for idx, url in enumerate(urls, 1):
        print(f"[L2 {idx}/{len(urls)}] Scraping: {url}")
        try:
            details = scrape_business_details(browser, url)
            search_term = urls_df.loc[urls_df['Listing URL'] == url, 'Search Term'].values[0]
            details['Search Term'] = search_term
            all_details.append(details)

            if idx % 5 == 0:  # Save every 5
                pd.DataFrame(all_details).to_excel(output_file, index=False)

        except Exception as e:
            print(f"❌ Failed {url}: {e}")

    browser.quit()
    pd.DataFrame(all_details).to_excel(output_file, index=False)
    print(f"\n✅ LAYER 2 finished. Saved to {output_file}")


# ----------------------
# MAIN
# ----------------------
if __name__ == "__main__":
    # First run Layer 1
    layer1_collect_urls()

    # Then run Layer 2
    layer2_scrape_details()
