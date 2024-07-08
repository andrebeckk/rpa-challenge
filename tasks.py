import os
import re
import json
import time
import logging
import requests
from pathlib import Path
from urllib.parse import unquote
from datetime import datetime
from dateutil.relativedelta import relativedelta

from robocorp.tasks import task
from RPA.Robocorp.WorkItems import WorkItems
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files as Excel

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

OUTPUT_DIR = Path(os.getenv("ROBOT_ARTIFACTS", "output"))
OUTPUT_EXCEL_FILE = 'challenge.xlsx'


@task
def solve_challenge():
    logger.info("Starting the solve_challenge task.")
    browser = Selenium()
    work_item = load_work_item(OUTPUT_DIR / "work-items-in/test-input/work-item.json")

    try:
        # Open browser and navigate to "https://apnews.com"
        logger.info("Opening the browser and navigating to https://apnews.com")
        browser.open_chrome_browser(url="https://apnews.com", maximized=True, headless=False)
        wait = WebDriverWait(browser.driver, 10)

        # Locate and click the search button


        # Check if a ad banner appears and close it
        try:
            browser.wait_until_element_is_visible("css:.fancybox-item fancybox-close", timeout=10)
            browser.click_element("css:.fancybox-item fancybox-close", timeout=10)
        except:
            pass

        logger.info("Locating and clicking the search button.")
        browser.wait_until_element_is_visible("css:.SearchOverlay", timeout=10)
        browser.click_element("css:.SearchOverlay")

        # Fill search field with the search phrase variable
        logger.info(f"Filling the search field with the phrase: {work_item['search_phrase']}")
        browser.wait_until_element_is_visible("css:.SearchOverlay-search-input", timeout=10)
        search_input = browser.find_element("css:.SearchOverlay-search-input")
        search_input.send_keys(work_item["search_phrase"])
        search_input.send_keys(Keys.RETURN)

        try:
            # Filter by category variable
            logger.info(f"Filtering by category: {work_item['news_category']}")
            browser.click_element("css:.SearchFilter-heading")
            browser.wait_until_element_is_visible("css:.SearchFilter-items-wrapper", timeout=10)
            browser.click_element(
                f"//span[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{work_item['news_category'].lower()}')]"
            )
            time.sleep(5)
        except:
            logger.error(f"Category not found.")

        # Sort the results by Newest first
        logger.info("Sorting the results by Newest first.")
        browser.wait_until_element_is_visible("css:.Select-input", timeout=10)
        browser.click_element("css:.Select-input")
        browser.wait_until_element_is_visible("css:select", timeout=10)
        select_element = browser.find_element("css:select")
        Select(select_element).select_by_visible_text("Newest")
        time.sleep(3)

        # Iterate over news and store the results in a list
        logger.info("Iterating over news results.")
        data_list = []
        while_controller = True
        while while_controller:
            search_result_list = get_search_list_results(browser=browser)

            for element in search_result_list:
                date = get_news_date(element, work_item["month_range"])
                if date:
                    title = get_news_title(element)
                    description = get_news_description(element)
                    picture_file_name = get_news_picture(element)
                    search_phrase_matches = count_search_matches(
                        work_item["search_phrase"], title, description
                    )
                    contains_money = contains_money_amount(title, description)

                    data_list.append(
                        {
                            "title": title,
                            "date": date,
                            "description": description,
                            "picture_file_name": picture_file_name,
                            "search_phrase_matches": search_phrase_matches,
                            "contains_money": contains_money,
                        }
                    )
                else:
                    while_controller = False

            browser.click_element("css:.Pagination-nextPage")
            wait.until(lambda browser: browser.execute_script("return document.readyState") == "complete")

        # Fill excel file with the result data
        logger.info("Filling Excel file with the results.")
        fill_excel_file(data_list=data_list)

        browser.close_all_browsers()
        logger.info("Task completed successfully.")

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        browser.close_all_browsers()


def get_search_list_results(browser):
    try:
        wait = WebDriverWait(browser.driver, 10)
        wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "div.SearchResultsModule-results > bsp-list-loadmore > div.PageList-items")))
        search_result = browser.find_element("css:.SearchResultsModule-results .PageListStandardD .PageList-items")
        search_result_list = search_result.find_elements(By.CLASS_NAME, "PageList-items-item")
        return search_result_list
    except Exception as e:
            logger.error(f"Error while getting search result list: {e}")



def get_news_title(element):
    try:    
        news_title = element.find_elements(By.CLASS_NAME, "PagePromo-title")[0].text
        return news_title
    except Exception as e:
            logger.error(f"Error while getting news title: {e}")

def get_news_description(element):
    try:
        news_description = element.find_elements(By.CLASS_NAME, "PagePromo-description")[0].text
        return news_description
    except Exception as e:
            logger.error(f"Error while getting news description: {e}")


def get_news_date(element, month_range):
    try:    
        if month_range < 1:
            month_range = 0
        else:
            month_range -= 1

        date_element = element.find_elements(By.TAG_NAME, "bsp-timestamp")[0]
        news_timestamp = int(date_element.get_attribute("data-timestamp"))
        news_date = datetime.fromtimestamp(news_timestamp / 1000)
        timestamp_month = news_date.month
        timestamp_year = news_date.year

        date_filter = datetime.now()
        month_filter = date_filter.month - month_range
        year_filter = date_filter.year

        if month_filter < 0:
            year_filter -= 1
            month_filter = 12 + month_filter

        if timestamp_month >= month_filter and timestamp_year >= year_filter:
            return news_date.strftime("%Y-%m-%d")
        else:
            return False
    except Exception as e:
            logger.error(f"Error while getting news date: {e}")


def get_news_picture(element):
    try:
        image_elements = element.find_elements(By.CLASS_NAME, "Image")
        if image_elements:
            news_picture = image_elements[0].get_attribute("src")
            encoded_url = news_picture.split("url=")[1]
            decoded_url = unquote(encoded_url)
            file_name = decoded_url.split("/")[-1] + ".jpeg"

            download_file(decoded_url, OUTPUT_DIR / 'images', file_name)
            return file_name
        else:
            return None
    except Exception as e:
            logger.error(f"Error while getting news image: {e}")


def count_search_matches(search_phrase, title, description):
    try:
        title_matches = title.lower().count(search_phrase.lower())
        description_matches = description.lower().count(search_phrase.lower())
        return title_matches + description_matches
    except Exception as e:
            logger.error(f"Error while counting how many times the search phrase appears in title or description: {e}")


def contains_money_amount(title, description):
    try:
        money_pattern = re.compile(r"\$\d{1,3}(,\d{3})*(\.\d{2})?" r"|\b\d+\s?(dollars|USD)\b", re.IGNORECASE)
        title_match = bool(money_pattern.search(title))
        description_match = bool(money_pattern.search(description))
        return title_match or description_match
    except Exception as e:
            logger.error(f"Error while checking if title or description contains any amount of money: {e}")

def download_file(url, target_dir, target_filename):
    try:
        response = requests.get(url)
        response.raise_for_status()
        target_dir.mkdir(exist_ok=True)
        local_file = target_dir / target_filename
        local_file.write_bytes(response.content)
        return local_file
    except Exception as e:
        logger.error(f"Error while downloading file from url: {url} {e}")


def load_work_item(file_path=None):
    try:
        wi = WorkItems()
        input_work_item = wi.get_input_work_item()
        if input_work_item.payload:
            work_item = input_work_item.payload
        else:
            with open(file_path, "r") as file:
                work_item = json.load(file)
        logger.info(f"Loaded work item: {work_item}")
        return work_item
    except Exception as e:
        logger.error(f"Error while loading work item: {e}")

def fill_excel_file(data_list):
    try:
        excel = Excel()

        header = [
            "title",
            "date",
            "description",
            "picture_file_name",
            "search_phrase_matches",
            "contains_money",
        ]
        rows = [header]

        for item in data_list:
            row = [
                item["title"],
                item["date"],
                item["description"],
                item["picture_file_name"],
                item["search_phrase_matches"],
                item["contains_money"],
            ]
            rows.append(row)

        excel.create_workbook(OUTPUT_DIR / OUTPUT_EXCEL_FILE)
        excel.append_rows_to_worksheet(rows, header=False)
        excel.save_workbook()
        excel.close_workbook()

        logger.info(f"Excel file '{OUTPUT_EXCEL_FILE}' created and saved with {len(data_list)} records.")

    except Exception as e:
        logger.error(f"Error while filling excel File: {e}")
