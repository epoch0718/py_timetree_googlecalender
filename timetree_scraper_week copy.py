import os
import json
import requests
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import re
from datetime import datetime
import pytz

load_dotenv()
TIMETREE_EMAIL = os.getenv("TIMETREE_EMAIL")
TIMETREE_PASSWORD = os.getenv("TIMETREE_PASSWORD")
CALENDAR_URL = os.getenv("TIMETREE_CALENDAR_URL")
GAS_WEBAPP_URL = os.getenv("GAS_WEBAPP_URL")


def parse_weekly_time(date_str, time_str):
    """
    ウィークリービューの多様な時間形式をパースし、"H:MM AM/PM"形式に変換する。
    例: "7:00 - 8:00" -> "7:00 AM"
    例: "10/14 23:00 - 10/15 0:00" -> "11:00 PM"
    """
    if not time_str:
        return None, date_str
    
    # 正規表現パターンを強化：先頭の時刻 (HH:MM) を確実に抽出
    # "10/14 23:00" や "7:00" など、どちらの形式でもマッチさせる
    match = re.search(r'(\d{1,2}:\d{2})', time_str)
    
    if match:
        time_24h = match.group(1) # "23:00" や "7:00" を取得
        try:
            # 24時間形式の文字列からdatetimeオブジェクトを生成
            dt_obj = datetime.strptime(time_24h, "%H:%M")
            
            # datetimeオブジェクトを "H:MM AM/PM" 形式の文字列にフォーマット
            # %-I は先頭に0がつかない12時間表記 (macOS/Linux)
            # Windowsの場合は #I を使う必要があるかもしれないが、GitHub Actions (Linux) では %-I でOK
            time_12h = dt_obj.strftime("%#I:%M %p" if os.name == 'nt' else "%-I:%M %p").replace(" 0", " ")
            
            return time_12h, date_str
        except Exception as e:
            print(f"Time parsing error for '{time_str}': {e}")
            return None, date_str
            
    # どのパターンにもマッチしなかった場合
    print(f"Could not find a valid time pattern in '{time_str}'")
    return None, date_str

def get_events_by_bounding_box_weekly(page):
    """
    ウィークリービュー専用のスクレイピング関数（HTML構造ベース版）。
    """
    events = []
    #page.screenshot(path="screenshot.png")

    month_year_text = page.locator('time.css-1b9gib').inner_text()
    print(month_year_text)
    print('month_year_text\n')
    match = re.search(r'(\d{4})年(\d{1,2})月', month_year_text)
    if not match:
        print("Error: Could not find month/year.")
        return []
    year, month = match.groups()
    month_year = f"{year}-{month.zfill(2)}"

    # 1. 日付（"13", "14" ...）と、それに対応する列番号（column="X"）をマッピングする
    date_to_column = {}
    day_number_elements = page.locator('[data-test-id="weekly-day-number"]').all()
    # ウィークリービューは月曜始まり固定と仮定 (column="2" から "8")
    for i, day_element in enumerate(day_number_elements):
        day = day_element.locator('div').inner_text()
        if day.isdigit():
            column_index = i + 2 # 月曜日がcolumn=2から始まる
            date_to_column[str(column_index)] = day

    if not date_to_column:
        print("Error: Could not map dates to columns.")
        return []
    print(f"Date to column map: {date_to_column}")

    # 2. 各列（column）をループし、その中のイベントをすべて取得する
    for col_index_str, day in date_to_column.items():
        # `column="X"` という属性を持つdiv要素を取得
        column_element = page.locator(f'div[column="{col_index_str}"]')
        if column_element.count() == 0:
            continue

        # その列の中にある全てのイベント要素を取得
        event_elements_in_col = column_element.locator('div[data-grid-item="true"]').all()
        date_str = f"{month_year}-{day.zfill(2)}"
        
        print(f"Processing Day {day} (column {col_index_str}): Found {len(event_elements_in_col)} events.")

        for event_element in event_elements_in_col:
            title_element = event_element.locator('h2.css-1j6im95')
            time_element = event_element.locator('time.css-1sbza0d')

            if title_element.count() == 0 or time_element.count() == 0:
                continue

            title = title_element.inner_text()
            time_raw = time_element.inner_text()
            
            # 時間形式をパースして "H:MM AM/PM" 形式に統一
            time_parsed, _ = parse_weekly_time(date_str, time_raw)
            
            if not time_parsed:
                print(f"Skipping event with unparsable time: {title} - {time_raw}")
                continue

            # メモを取得
            memo = None
            try:
                event_element.click(timeout=1000)
                page.wait_for_timeout(500)
                memo_element = page.locator('ul.css-19zmclu + div p, ul.css-19zmclu + div a')
                if memo_element.count() > 0:
                    memo = memo_element.inner_text()
                
                close_button = page.get_by_label("閉じる", exact=True)
                if close_button.count() > 0:
                    close_button.click(timeout=1000)
                    page.wait_for_timeout(500)
            except Exception as e:
                print(f"Could not get memo for '{title}': {e}")


            event = {
                'date': date_str,
                'time': time_parsed,
                'title': title.strip(),
                'memo': memo
            }
            events.append(event)

    print(f"Total events scraped: {len(events)}")
    return events

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True , # 開発中はFalse推奨
            args=["--force-device-scale-factor=1.1"]
        )
        context = browser.new_context(
            timezone_id="Asia/Tokyo",
            locale="ja-JP"
            # ビューポートサイズを固定してレイアウトを安定させる
            #viewport={'width': 1400, 'height': 1080}
        )
        page = context.new_page()

        print("Navigating to calendar and logging in...")
        page.goto(CALENDAR_URL, wait_until="networkidle")
        
        if "signin" in page.url:
            print("Login page detected. Logging in...")
            page.fill('input[type="email"]', TIMETREE_EMAIL)
            page.fill('input[type="password"]', TIMETREE_PASSWORD)
            page.click('button[type="submit"]')
        
        print("Waiting for weekly calendar to load...")
        # 待機セレクタをウィークリービュー用に変更
        page.wait_for_selector('[data-test-id="weekly-calendar-root"]')
        page.wait_for_timeout(5000)
        print("Calendar loaded.")

        # ウィークリービュー用の新しい関数を呼び出す
        events = get_events_by_bounding_box_weekly(page)
        
        print("\n--- Scraped Events ---")
        sorted_events = sorted(events, key=lambda x: (x['date'], x['time'] or ''))
        print(json.dumps(sorted_events, indent=2, ensure_ascii=False))

        if not events:
            print("\nWarning: No events were scraped. Skipping sending data to GAS.")
        elif GAS_WEBAPP_URL:
            print("\n--- Sending data to Google Apps Script ---")
            try:
                headers = {'Content-Type': 'application/json'}
                # ソートした結果を送信
                response = requests.post(GAS_WEBAPP_URL, data=json.dumps(sorted_events), headers=headers)
                response.raise_for_status()
                print(f"Successfully sent data. Status: {response.status_code}")
                print(f"Response from GAS:")
                gas_response = response.json()
                print(json.dumps(gas_response, indent=2, ensure_ascii=False))
            except requests.exceptions.RequestException as e:
                print(f"Error sending data to GAS: {e}")
            except json.JSONDecodeError:
                # もしJSONでなかった場合は、テキストとしてそのまま表示
                print("Error Response is not in JSON format. Raw text:")
                print(response.text)
        else:
            print("\nGAS_WEBAPP_URL is not set. Skipping sending data to GAS.")



        context.close()
        browser.close()

if __name__ == "__main__":
    main()