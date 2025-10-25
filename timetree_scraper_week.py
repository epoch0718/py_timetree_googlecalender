import os
import json
import requests
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import re
from datetime import datetime
import pytz

# --- 設定値 (環境変数) ---
load_dotenv()
TIMETREE_EMAIL = os.getenv("TIMETREE_EMAIL")
TIMETREE_PASSWORD = os.getenv("TIMETREE_PASSWORD")
CALENDAR_URL = os.getenv("TIMETREE_CALENDAR_URL")
GAS_WEBAPP_URL = os.getenv("GAS_WEBAPP_URL")

# --- ヘルパー関数 ---

def parse_weekly_time(date_str, time_str):
    """
    ウィークリービューの時間形式をパースし、"HH:MM" (24時間形式)の文字列を返す。
    """
    if not time_str:
        return None
    
    match = re.search(r'(\d{1,2}:\d{2})', time_str)
    if match:
        return match.group(1) # "7:00" や "23:00" をそのまま返す
            
    print(f"Could not find a valid time pattern in '{time_str}'")
    return None

# --- データ収集関数 ---

def scrape_one_week_events(page):
    """
    現在表示されている1週間分のイベントをスクレイピングする関数。
    """
    events = []
    
    try:
        month_year_locator = page.locator('time.css-1b9gib')
        month_year_locator.wait_for(state='visible', timeout=10000)
        month_year_text = month_year_locator.inner_text()
    except Exception as e:
        print(f"Error: Could not find or read month/year text: {e}")
        return []

    match = re.search(r'(\d{4})年(\d{1,2})月', month_year_text)
    if not match:
        print(f"Error: Could not parse month/year from text: '{month_year_text}'")
        return []
    year, month = match.groups()
    month_year = f"{year}-{month.zfill(2)}"

    date_to_column = {}
    day_number_elements = page.locator('[data-test-id="weekly-day-number"]').all()
    for i, day_element in enumerate(day_number_elements):
        day = day_element.locator('div').inner_text()
        if day.isdigit():
            date_to_column[str(i + 2)] = day # 月曜日(i=0)がcolumn=2から

    for col_index_str, day in date_to_column.items():
        column_element = page.locator(f'div[column="{col_index_str}"]')
        if column_element.count() == 0:
            continue

        event_elements_in_col = column_element.locator('div[data-grid-item="true"]').all()
        date_str = f"{month_year}-{day.zfill(2)}"

        for event_element in event_elements_in_col:
            title_element = event_element.locator('h2.css-1j6im95')
            time_element = event_element.locator('time.css-1sbza0d')

            if title_element.count() == 0 or time_element.count() == 0:
                continue

            title = title_element.inner_text()
            time_raw = time_element.inner_text()
            time_parsed = parse_weekly_time(date_str, time_raw)
            if not time_parsed:
                continue

            memo = None
            try:
                event_element.click(timeout=1000, force=True)
                page.wait_for_timeout(500)
                memo_element = page.locator('p.exlc7u1.vjrcbi0')
                if memo_element.count() > 0:
                    memo = memo_element.inner_text()
                
                close_button = page.get_by_label("閉じる", exact=True)
                if close_button.count() > 0:
                    close_button.click(timeout=1000)
                    page.wait_for_timeout(500)
            except Exception as e:
                print(f"Could not get memo for '{title}': {e}")

            events.append({
                'date': date_str,
                'time': time_parsed,
                'title': title.strip(),
                'memo': memo
            })
            
    return events

# --- メインの司令塔 ---

def main():
    all_events = []
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--force-device-scale-factor=1.1"],
        )
        context = browser.new_context(
            timezone_id="Asia/Tokyo",
            locale="ja-JP",
            #viewport={'width': 1920, 'height': 1080}
        )
        page = context.new_page()

        print("Navigating to calendar and logging in...")
        page.goto(CALENDAR_URL, wait_until="networkidle")
        
        if "signin" in page.url:
            print("Login page detected. Logging in...")
            page.fill('input[type="email"]', TIMETREE_EMAIL)
            page.fill('input[type="password"]', TIMETREE_PASSWORD)
            page.click('button[type="submit"]')
        
        print("Waiting for initial calendar load...")
        page.wait_for_selector('[data-test-id="weekly-calendar-root"]', timeout=30000)
        page.wait_for_load_state('networkidle')
        print("Calendar loaded.")

        # --- 1. 先週のデータを収集 ---
        print("\n--- Scraping Previous Week ---")
        page.locator('[data-test-id="pervious-button"]').click()
        page.wait_for_load_state('networkidle') # ページ遷移・データ更新を待つ
        page.wait_for_timeout(2000) # 念のため追加待機
        previous_week_events = scrape_one_week_events(page)
        all_events.extend(previous_week_events)
        print(previous_week_events)
        print(f"Found {len(previous_week_events)} events in the previous week.")

        # --- 2. 今週（最初に表示された週）のデータを収集 ---
        print("\n--- Scraping Current Week ---")
        page.locator('[data-test-id="next-button"]').click()
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(2000)
        current_week_events = scrape_one_week_events(page)
        all_events.extend(current_week_events)
        print(current_week_events)
        print(f"Found {len(current_week_events)} events in the current week.")

        # --- 3. 来週のデータを収集 ---
        print("\n--- Scraping Next Week ---")
        page.locator('[data-test-id="next-button"]').click()
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(2000)
        next_week_events = scrape_one_week_events(page)
        all_events.extend(next_week_events)
        print(next_week_events)
        print(f"Found {len(next_week_events)} events in the next week.")

        # --- 4. 全データをまとめてGASに送信 ---
        print(f"\nTotal events scraped from 3 weeks: {len(all_events)}")
        
        # 日付と時間でソート
        sorted_events = sorted(all_events, key=lambda x: (x['date'], x['time'] or ''))

        print("\n--- Final Scraped Events (All Weeks) ---")
        print(json.dumps(sorted_events, indent=2, ensure_ascii=False))

        if not sorted_events:
            print("\nWarning: No events were scraped. Skipping sending data to GAS.")
        elif GAS_WEBAPP_URL:
            print("\n--- Sending data to Google Apps Script ---")
            try:
                headers = {'Content-Type': 'application/json'}
                response = requests.post(GAS_WEBAPP_URL, data=json.dumps(sorted_events), headers=headers)
                response.raise_for_status()
                print(f"Successfully sent data. Status: {response.status_code}")
                print(f"Response from GAS:")
                gas_response = response.json()
                print(json.dumps(gas_response, indent=2, ensure_ascii=False))
            except requests.exceptions.RequestException as e:
                print(f"Error sending data to GAS: {e}")
            except json.JSONDecodeError:
                print("Error: Response from GAS is not in JSON format. Raw text:")
                print(response.text)
        else:
            print("\nGAS_WEBAPP_URL is not set. Skipping sending data to GAS.")

        context.close()
        browser.close()

if __name__ == "__main__":
    main()