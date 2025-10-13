import os
import json
import requests
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import re

load_dotenv()
TIMETREE_EMAIL = os.getenv("TIMETREE_EMAIL")
TIMETREE_PASSWORD = os.getenv("TIMETREE_PASSWORD")
CALENDAR_URL = os.getenv("TIMETREE_CALENDAR_URL")
GAS_WEBAPP_URL = os.getenv("GAS_WEBAPP_URL")

def get_events_by_bounding_box(page):
    """
    要素の画面上の座標（Bounding Box）を基準に、日付とイベントを紐付ける最終ロジック。
    """
    events = []
    month_year = page.locator('time').get_attribute('datetime')

    # 1. 全ての日付セルの日付と座標を取得
    date_boxes = []
    gridcells = page.locator('[role="gridcell"]').all()
    for cell in gridcells:
        # セル内の日付番号要素を取得 (div > div 構造を仮定)
        day_element = cell.locator('div > div')
        if day_element.count() > 0:
            day = day_element.inner_text()
            if day.isdigit():
                box = cell.bounding_box()
                if box:
                    date_boxes.append({'day': day, 'box': box})

    if not date_boxes:
        print("Error: Could not find or measure date cells.")
        return []
    
    print(f"Found and measured {len(date_boxes)} date cells.")

    # 2. 全てのイベントのタイトル、時間、座標を取得
    event_details = []
    event_elements = page.locator('.lndlxo5').all()


    for i,event_element in enumerate(event_elements):
        box = event_element.bounding_box()
        if not box:
            continue
        
        button = event_element.locator('button')
        if button.count() > 0:
            title_element = button.locator('.lndlxo9')
            time_element = button.locator('._1r1c5vla')
            
            title = ""
            if title_element.count() > 0:
                title = title_element.inner_text()
            else:
                full_text = button.inner_text()
                time_text = time_element.inner_text() if time_element.count() > 0 else ""
                title = full_text.replace(time_text, "").strip()
                print(full_text + " " + time_text + " " + title)
                print('full_text + " " + time_text + " " + title')

            time = time_element.inner_text() if time_element.count() > 0 else None

            print(time)
            print('time\n')

            button.click()
            page.wait_for_timeout(500)
            memo=None
            memo = page.locator('p.exlc7u1.vjrcbi0')
            if memo.count() >0:
                if memo.inner_text().find('Locator') == -1:
                    memo = memo.inner_text()
                    print(memo)
            else:
                memo = None

            if title != "" and time != None:
                event_details.append({
                    'title': title.strip(),
                    'time': time,
                    'box': box,
                    'memo': memo
                })
            closebotton = page.get_by_label("Close")
            if closebotton.count() > 0:
                closebotton.click()
                page.wait_for_timeout(500)


    if not event_details:
        print("Warning: Found date cells, but no event elements.")
        return []
        
    print(f"Found and measured {len(event_details)} events.")

    print(event_details)
    print('event_details_1')

    # 3. イベントの座標がどの日付セルの範囲内にあるかを判定して紐付け
    for event_detail in event_details:
        event_box = event_detail['box']
        # イベントの中心点がどのセルに入るかで判定
        event_center_x = event_box['x'] + event_box['width'] / 2
        event_center_y = event_box['y'] + event_box['height'] / 2

        for date_info in date_boxes:
            date_box = date_info['box']
            if (date_box['x'] <= event_center_x < date_box['x'] + date_box['width'] and
                date_box['y'] <= event_center_y < date_box['y'] + date_box['height']):
                
                event = {
                    'date': f"{month_year}-{str(date_info['day']).zfill(2)}",
                    'time': event_detail['time'],
                    'title': event_detail['title'],
                    'memo': event_detail['memo']
                }
                if event not in events:
                    events.append(event)
                break # マッチしたら次のイベントへ
    
    print(events)
    print('events_2')
    
    return events

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--force-device-scale-factor=1.1"]
            )
        
        # 日本のタイムゾーン（Asia/Tokyo）を指定して、新しいブラウザコンテキストを作成
        context = browser.new_context(
            timezone_id="Asia/Tokyo"
        )
        page = browser.new_page()

        print("Navigating to calendar and logging in...")
        page.goto(CALENDAR_URL, wait_until="networkidle")
        
        if "signin" in page.url:
            print("Login page detected. Logging in...")
            page.fill('input[type="email"]', TIMETREE_EMAIL)
            page.fill('input[type="password"]', TIMETREE_PASSWORD)
            page.click('button[type="submit"]')
        
        print("Waiting for calendar to load...")
        page.wait_for_selector('[data-test-id="calendar-main"]')
        page.wait_for_timeout(5000)
        print("Calendar loaded.")

        # 新しいスクレイピング関数を呼び出す
        events = get_events_by_bounding_box(page)
        

        print("\n--- Scraped Events ---")
        # 日付順にソートして表示
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

        browser.close()

if __name__ == "__main__":
    main()