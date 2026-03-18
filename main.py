import time
import smtplib
import os
import ssl
# import schedule
# import threading
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


def load_constants():
    """環境変数から設定値を読み込む"""
    required_keys = [
        'ID_first', 'ID_second', 'password', 'address1',
        'GMAIL_USER', 'GMAIL_APP_PASSWORD', 'NOTIFICATION_EMAIL',
        'SCHEDULE_TIME'
    ]
    
    constants = {}
    missing_keys = []
    
    for key in required_keys:
        value = os.environ.get(key)
        if value is None:
            missing_keys.append(key)
        else:
            constants[key] = value
    
    if missing_keys:
        raise ValueError(f"必須の環境変数が不足しています: {', '.join(missing_keys)}")
    
    return constants
CONFIG = load_constants()

ID_first = CONFIG['ID_first']
ID_second = CONFIG['ID_second']
password = CONFIG['password']
address1 = CONFIG['address1']
GMAIL_USER = CONFIG['GMAIL_USER']
GMAIL_APP_PASSWORD = CONFIG['GMAIL_APP_PASSWORD']
NOTIFICATION_EMAIL = CONFIG['NOTIFICATION_EMAIL']
SCHEDULE_TIME = CONFIG['SCHEDULE_TIME']


def send_email(subject, body_html):
    """メール送信関数（エラー時は警告のみで続行）"""
    try:
        sender = GMAIL_USER
        receiver = NOTIFICATION_EMAIL
        password = GMAIL_APP_PASSWORD.replace(" ", "").strip()

        if not password:
            print("警告: Gmailアプリパスワードが設定されていません。メール送信をスキップします。")
            return

        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = receiver
        msg["Subject"] = subject
        msg.attach(MIMEText(body_html, "html"))

        context = ssl.create_default_context()
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=30) as server:
                server.login(sender, password)
                server.send_message(msg)
            print(f"メール送信成功: {subject}")
        except (ssl.SSLEOFError, ssl.SSLError, smtplib.SMTPException, ConnectionError, OSError) as e:
            print(f"警告: メール送信に失敗しました（続行します）: {e}")
            return
    except Exception as e:
        print(f"警告: メール送信処理中にエラーが発生しました（続行します）: {e}")
        return

def main():
    """メインの自動化処理"""
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)
    
    try:
        print("ステップ1: ログインページを開いてログイン中...")
        driver.get("https://share.timescar.jp/")
        print("ページを読み込みました。JavaScriptの初期化を待機中...")
        time.sleep(2)
        print("ログインフォームを待機中...")
        
        print("ログインフォーム要素を待機中...")
        card_no1_found = False
        
        try:
            wait.until(EC.visibility_of_element_located((By.ID, 'cardNo1')))
            card_no1_found = True
            print("ログインフォームが見つかりました（表示中）")
        except TimeoutException:
            try:
                wait.until(EC.presence_of_element_located((By.ID, 'cardNo1')))
                card_no1_found = True
                print("ログインフォームが見つかりました（DOMに追加済み）")
            except TimeoutException:
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[id="cardNo1"]')))
                    card_no1_found = True
                    print("ログインフォームが見つかりました（代替セレクタ）")
                except Exception as e:
                    print(f"エラー: ログインフォーム要素が見つかりませんでした")
                    print(f"エラー詳細: {e}")
                    page_title = driver.title
                    page_url = driver.current_url
                    print(f"現在のページタイトル: {page_title}")
                    print(f"現在のページURL: {page_url}")
                    element_exists = driver.execute_script('return document.getElementById("cardNo1") !== null')
                    print(f"要素 #cardNo1 がDOMに存在: {element_exists}")
                    driver.save_screenshot('debug_login_page.png')
                    print("スクリーンショットを debug_login_page.png として保存しました")
                    raise Exception(f"ログインフォームが見つかりません。ページURL: {page_url}, タイトル: {page_title}")
        
        if not card_no1_found:
            raise Exception("複数の戦略を試してもログインフォームの位置を特定できませんでした")
        
        card_no1 = driver.find_element(By.ID, 'cardNo1')
        card_no1.click()
        card_no1.clear()
        card_no1.send_keys(ID_first)
        driver.execute_script('arguments[0].dispatchEvent(new Event("input", { bubbles: true })); arguments[0].dispatchEvent(new Event("change", { bubbles: true }));', card_no1)
        time.sleep(0.2)
        
        card_no2 = driver.find_element(By.ID, 'cardNo2')
        card_no2.click()
        card_no2.clear()
        card_no2.send_keys(ID_second)
        driver.execute_script('arguments[0].dispatchEvent(new Event("input", { bubbles: true })); arguments[0].dispatchEvent(new Event("change", { bubbles: true }));', card_no2)
        time.sleep(0.2)
        
        password_field = driver.find_element(By.ID, 'tpPassword')
        password_field.click()
        password_field.clear()
        password_field.send_keys(password)
        driver.execute_script('arguments[0].dispatchEvent(new Event("input", { bubbles: true })); arguments[0].dispatchEvent(new Event("change", { bubbles: true }));', password_field)
        time.sleep(0.2)
        
        driver.find_element(By.ID, 'tpLogin').click()
        time.sleep(2)
        wait.until(EC.presence_of_element_located((By.ID, 'yakkan_box')))
        send_email(subject="🎄🎄タイムズカー - ログイン成功🎄🎄", body_html="ログインに成功しました")

        display_value = driver.execute_script("""
            const el = document.getElementById('yakkan_box');
            if (!el) return null;
            return window.getComputedStyle(el).display;
        """)

        if display_value == "block":
            print("◎ yakkan_boxが表示されています（display: block）")
            
            xpaths = [
                '/html/body/div[2]/div[2]/div[6]/div[2]/p/a/img',
                '/html/body/div[2]/div[2]/div[5]/div[2]/p/a/img',
                '/html/body/div[2]/div[2]/div[4]/div[2]/p/a/img',
                '/html/body/div[2]/div[2]/div[3]/div[2]/p/a/img'
            ]
            
            for xpath in xpaths:
                try:
                    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
                    time.sleep(0.3)
                    wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    try:
                        element.click()
                    except:
                        driver.execute_script("arguments[0].click();", element)
                    time.sleep(0.2)
                except Exception as e:
                    print(f"警告: 要素のクリックに失敗しました（XPath: {xpath}）: {e}")
                    continue
        else:
            print("× yakkan_boxは表示されていません。クリックをスキップします。")

        wait.until(EC.visibility_of_element_located((By.ID, 'stationLink')))
        driver.find_element(By.ID, 'stationLink').click()
        time.sleep(2)
            
        print("ステップ2: ステーションを選択中...")
        wait.until(EC.visibility_of_element_located((By.ID, 'stationNm')))
        driver.find_element(By.ID, 'stationNm').click()
        wait.until(EC.visibility_of_element_located((By.ID, 'nameAdr-s')))
        driver.find_element(By.ID, 'nameAdr-s').clear()
        driver.find_element(By.ID, 'nameAdr-s').send_keys(address1)
        driver.find_element(By.ID, 'doNameAdrSearch').click()
        time.sleep(2)
        
        wait.until(EC.visibility_of_element_located((By.ID, 'isEnableToReserve')))
        driver.find_element(By.ID, 'isEnableToReserve').click()
        time.sleep(2)
        
        print("ステップ3: 異なる車両IDで予約を試行中...")            
        car_ids = [1212609, 1235450, 1238278, 1256218]
        reservation_successful = False
        
        for car_id in car_ids:
            print(f"車両ID {car_id} を試行中...")
            
            try:
                wait.until(EC.visibility_of_element_located((By.ID, 'carId')))
                
                car_id_select = Select(driver.find_element(By.ID, 'carId'))
                car_id_select.select_by_value(str(car_id))
                print("◎車両選択")

                wait.until(EC.visibility_of_element_located((By.ID, 'pack')))
                pack_select = Select(driver.find_element(By.ID, 'pack'))
                pack_options = pack_select.options
                if len(pack_options) >= 2:
                    second_last_index = len(pack_options) - 2
                    pack_select.select_by_index(second_last_index)
                    print("◎パック選択")

                wait.until(EC.visibility_of_element_located((By.ID, 'dateStart')))
                date_start_select = Select(driver.find_element(By.ID, 'dateStart'))
                date_start_options = date_start_select.options
                if len(date_start_options) > 0:
                    last_index = len(date_start_options) - 1
                    date_start_select.select_by_index(last_index)
                    print("◎開始日選択")
                    time.sleep(0.5)

                wait.until(EC.visibility_of_element_located((By.ID, 'dateEnd')))
                date_end_select = Select(driver.find_element(By.ID, 'dateEnd'))
                date_end_options = date_end_select.options
                if len(date_end_options) > 0:
                    second_last_index = len(date_end_options) - 2
                    date_end_select.select_by_index(second_last_index)
                    print("◎終了日選択")
                    time.sleep(0.5)

                wait.until(EC.visibility_of_element_located((By.ID, 'hourStart')))
                hour_start_elem = driver.find_element(By.ID, 'hourStart')
                
                driver.execute_script("""
                    const select = document.getElementById('hourStart');
                    if (select) {
                        const option = Array.from(select.options).find(opt => opt.value === '11');
                        if (option) {
                            select.value = '11';
                        } else {
                            const newOption = document.createElement('option');
                            newOption.value = '11';
                            newOption.textContent = '11';
                            select.appendChild(newOption);
                            select.value = '11';
                        }
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                """)
                time.sleep(0.3)
                
                current_value = hour_start_elem.get_attribute('value')
                print(f"現在の開始時間（時）: {current_value}")
                
                if current_value != "11":
                    try:
                        Select(hour_start_elem).select_by_value("11")
                    except:
                        pass
                    time.sleep(0.3)
                    current_value = hour_start_elem.get_attribute('value')
                    if current_value != "11":
                        driver.execute_script("""
                            const select = document.getElementById('hourStart');
                            if (select) {
                                select.value = '11';
                                select.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                        """)
                        time.sleep(0.3)
                        current_value = hour_start_elem.get_attribute('value')
                
                if current_value == "11":
                    print("✓ 開始時間（時）を11に設定しました")
                else:
                    print(f"⚠ 警告: 開始時間（時）が '{current_value}' です。期待値は '11' です。")

                wait.until(EC.visibility_of_element_located((By.ID, 'minuteStart')))
                minute_start_elem = driver.find_element(By.ID, 'minuteStart')
                
                driver.execute_script("""
                    const select = document.getElementById('minuteStart');
                    if (select) {
                        const option = Array.from(select.options).find(opt => opt.value === '00' || opt.value === '0');
                        if (option) {
                            select.value = option.value;
                        } else {
                            if (select.options.length > 0) {
                                select.selectedIndex = 0;
                            }
                        }
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                """)
                time.sleep(0.3)
                
                selected_minute = minute_start_elem.get_attribute('value')
                print(f"◎開始時間(分): {selected_minute}")
                
                if selected_minute not in ["00", "0"]:
                    try:
                        Select(minute_start_elem).select_by_value("00")
                    except:
                        try:
                            Select(minute_start_elem).select_by_value("0")
                        except:
                            Select(minute_start_elem).select_by_index(0)
                    time.sleep(0.3)
                    selected_minute = minute_start_elem.get_attribute('value')
                    print(f"◎開始時間(分) 再試行後: {selected_minute}")

                wait.until(EC.visibility_of_element_located((By.ID, 'hourEnd')))
                hour_end_elem = driver.find_element(By.ID, 'hourEnd')
                
                driver.execute_script("""
                    const select = document.getElementById('hourEnd');
                    if (select) {
                        const option = Array.from(select.options).find(opt => opt.value === '19');
                        if (option) {
                            select.value = '19';
                        } else {
                            const newOption = document.createElement('option');
                            newOption.value = '19';
                            newOption.textContent = '19';
                            select.appendChild(newOption);
                            select.value = '19';
                        }
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                """)
                time.sleep(0.3)
                
                selected_hour_end = hour_end_elem.get_attribute('value')
                print(f"◎終了時間(時): {selected_hour_end}")
                
                if selected_hour_end != "19":
                    try:
                        Select(hour_end_elem).select_by_value("19")
                    except:
                        pass
                    time.sleep(0.3)
                    selected_hour_end = hour_end_elem.get_attribute('value')
                    if selected_hour_end != "19":
                        driver.execute_script("""
                            const select = document.getElementById('hourEnd');
                            if (select) {
                                select.value = '19';
                                select.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                        """)
                        time.sleep(0.3)
                        selected_hour_end = hour_end_elem.get_attribute('value')
                
                if selected_hour_end == "19":
                    print("✓ 終了時間（時）を19に設定しました")
                else:
                    print(f"⚠ 警告: 終了時間（時）が '{selected_hour_end}' です。期待値は '19' です。")

                wait.until(EC.visibility_of_element_located((By.ID, 'minuteEnd')))
                minute_end_elem = driver.find_element(By.ID, 'minuteEnd')
                
                driver.execute_script("""
                    const select = document.getElementById('minuteEnd');
                    if (select) {
                        const option = Array.from(select.options).find(opt => opt.value === '00' || opt.value === '0');
                        if (option) {
                            select.value = option.value;
                        } else {
                            if (select.options.length > 0) {
                                select.selectedIndex = 0;
                            }
                        }
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                """)
                time.sleep(0.3)
                
                selected_minute_end = minute_end_elem.get_attribute('value')
                print(f"◎終了時間(分): {selected_minute_end}")
                
                if selected_minute_end not in ["00", "0"]:
                    try:
                        Select(minute_end_elem).select_by_value("00")
                    except:
                        try:
                            Select(minute_end_elem).select_by_value("0")
                        except:
                            Select(minute_end_elem).select_by_index(0)
                    time.sleep(0.3)
                    selected_minute_end = minute_end_elem.get_attribute('value')
                    print(f"◎終了時間(分) 再試行後: {selected_minute_end}")
                time.sleep(10)
                wait.until(EC.visibility_of_element_located((By.ID, 'exemptNocFlgYes')))
                driver.find_element(By.ID, 'exemptNocFlgYes').click()
                
                wait.until(EC.visibility_of_element_located((By.ID, 'doCheck')))
                driver.find_element(By.ID, 'doCheck').click()
                time.sleep(2)
                
                try:
                    error_title = driver.find_element(By.CSS_SELECTOR, 'p.errortitle')
                    error_exists = True
                except NoSuchElementException:
                    error_exists = False
                
                if error_exists:
                    print(f"車両ID {car_id} でエラーが見つかりました。エラーを記録中...")
                    try:
                        error_div = driver.find_element(By.ID, 'd_error')
                        error_html = error_div.get_attribute('innerHTML')
                    except NoSuchElementException:
                        error_html = "エラー詳細が見つかりませんでした"
                    
                    email_body = f"""
                    <html>
                    <body>
                    <h2>車両ID {car_id} の予約エラー</h2>
                    <div>
                    {error_html}
                    </div>
                    </body>
                    </html>
                    """
                    send_email(f"タイムズカー予約エラー - 車両ID {car_id}", email_body)
                    
                    continue
                else:
                    print(f"車両ID {car_id} の予約チェックが成功しました！")
                    reservation_successful = True
                    break
                    
            except Exception as e:
                print(f"車両ID {car_id} の処理中にエラーが発生しました: {e}")
                continue
        
        if not reservation_successful:
            print("すべての車両IDが失敗しました。通知メールを送信中...")
            email_body = """
            <html>
            <body>
            <h2>予約不可</h2>
            <p>すべての車両IDを試しましたが、予約を作成できませんでした。</p>
            </body>
            </html>
            """
            send_email("タイムズカー - 予約不可", email_body)
            return
        
        print("ステップ4: 最終登録を完了中...")
        wait.until(EC.visibility_of_element_located((By.ID, 'doOnceRegist')))
        driver.find_element(By.ID, 'doOnceRegist').click()
        time.sleep(2)
        
        print("予約処理が正常に完了しました！")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        email_body = f"""
        <html>
        <body>
        <h2>自動化エラー</h2>
        <p>自動化処理中にエラーが発生しました:</p>
        <pre>{str(e)}</pre>
        </body>
        </html>
        """
        send_email("タイムズカー自動化エラー", email_body)
    
    finally:
        time.sleep(3)
        driver.quit()


# def run_main():
#     """main関数を実行するラッパー関数"""
#     print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] スケジュールされたタスクを開始中...")
#     try:
#         main()
#         print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] タスクが正常に完了しました。")
#     except Exception as e:
#         print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] タスクがエラーで失敗しました: {e}")
#         import traceback
#         traceback.print_exc()


# def run_scheduler():
#     """別スレッドでスケジューラーを実行"""
#     while True:
#         schedule.run_pending()
#         time.sleep(60)


if __name__ == "__main__":
    print("=" * 60)
    print("タイムズカー予約自動化")
    print("=" * 60)
    print(f"開始時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    main()
