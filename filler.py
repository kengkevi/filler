import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, WebDriverException, NoSuchWindowException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# ============================================================
# デバッグモード設定
# ============================================================
DEBUG_MODE = True  # Trueにすると詳細なログをターミナルに出力

# ============================================================
# あなたの情報（初期値）
# ============================================================
# UIから編集可能です。ここでの値は起動時の初期値として使われます
MY_DATA = {
    "company": "ai-supporters",
    "company_furigana": "エーアイサポーターズ",
    "name_last": "橋本",
    "name_first": "健吾",
    "full_name": "橋本健吾",
    "furigana_last": "ハシモト",
    "furigana_first": "ケンゴ",
    "full_furigana": "ハシモトケンゴ",
    "email": "info@ai-supporters.jp",
    "email_confirm": "info@ai-supporters.jp",
    "tel": "080-3391-0114",
    "postal_code": "107-0052",
    "postal_code_1": "107",
    "postal_code_2": "0052",
    "prefecture": "東京都",
    "city": "港区",
    "address": "東京都港区赤坂8-12-4",
    "address_building": "",
    "url": "https://ai-supporters-lp1.vercel.app",
    "department": "営業部",
    "subject": "生成AIで実現する一歩先のスカウト戦略【ディスカッションのお願い】",
    "inquiry_body": """ご担当者様

突然のご連絡失礼いたします。 私は、人材紹介会社様の成長をAIでお手伝いするai-supportersの橋本と申します。
「候補者一人ひとりに合わせたスカウトメールを打ちたいが、時間が足りない」 「面談記録の作成や推薦状の準備に追われ、肝心の候補者様や企業様と向き合う時間が削られてしまう」
もし、このようなことでお悩みでしたら、AIがそのお悩みを解決する"新しい一手"になるかもしれません。
私どもがご提案するのは、単に作業が楽になる便利な道具ではありません。 スカウトメールの作成、面談記録の要約、日報の準備といった日々の業務をAIに任せ、そこで生まれた大切な時間で、候補者様との面談や企業様へのご提案といった、事業の核となる活動に集中する。 これは、かけた費用以上に会社の「成約数」を大きく伸ばす、成長のための仕組みづくりです。
私どものサービスや導入事例については、こちらのページで詳しくご紹介しております。 

▼ai-supporters サービス紹介ページ 
https://ai-supporters-lp1.vercel.app

つきましては、まず貴社のお悩みや「こうなったら嬉しい」というお話を、一度オンラインで30分ほどお聞かせいただけないでしょうか。 
情報収集の場としても歓迎ですので、ぜひお気軽にご活用ください。 もしご興味をお持ちいただけましたら、ご都合の良い日時を2〜3つお知らせいただけますと幸いです。
ご連絡を心よりお待ちしております。

ai-supporters 橋本"""
}

# ============================================================
# キーワード設定
# ============================================================
# フォームの項目名（label, placeholder等）からフィールドを特定するためのキーワード
KEYWORDS = {
    "company": ["会社名", "貴社名", "企業名", "法人名", "団体名", "御社名", "店舗名", "company"],
    "company_furigana": ["会社名（フリガナ）", "会社名（ふりがな）", "企業名（フリガナ）", "法人名（フリガナ）", "法人名（ふりがな）", "会社名フリガナ", "法人名フリガナ", "カイシャメイ", "ホウジンメイ"],
    "name_last": ["姓", "苗字", "氏", "lastName", "sei"],
    "name_first": ["名", "firstName", "mei"],
    "full_name": ["お名前", "氏名", "担当者名", "ご担当者", "御担当者", "名前", "fullname"],
    "furigana_last": ["姓（フリガナ）", "姓（ふりがな）", "セイ", "lastKana", "sei_kana"],
    "furigana_first": ["名（フリガナ）", "名（ふりがな）", "メイ", "firstKana", "mei_kana"],
    "full_furigana": ["フリガナ", "ふりがな", "カナ"],
    "email": ["メールアドレス", "Email", "mail", "email"],
    "email_confirm": ["メールアドレス（確認）", "確認のため", "確認用", "email_confirm", "confirm_email", "confirmemail"],
    "tel": ["電話番号", "tel", "電話"],
    "postal_code": ["郵便番号", "zipcode", "zip", "〒", "postal"],
    "postal_code_1": ["郵便番号（前半）", "郵便番号1", "zip1", "〒1", "-first", "postal-first"],
    "postal_code_2": ["郵便番号（後半）", "郵便番号2", "zip2", "〒2", "-last", "postal-last"],
    "prefecture": ["都道府県"],
    "city": ["市区町村"],
    "address": ["住所", "番地"],
    "address_building": ["建物名", "ビル名"],
    "url": ["ホームページ", "URL", "website"],
    "department": ["部署", "所属", "部署名", "所属部署"],
    "subject": ["件名", "題名", "subject", "ご用件", "用件", "タイトル"],
    "inquiry_body": ["お問い合わせ内容", "お問合せ内容", "お問合せ詳細", "ご依頼", "依頼内容", "内容", "詳細", "メッセージ", "inquiry", "message", "comment", "本文"],
    "agree": ["同意する", "個人情報の取り扱い", "プライバシーポリシー"]
}

# ============================================================
# name属性パターンマッピング
# ============================================================
# 様々なフォームで使われるname属性のパターンを登録
# 新しいフォームに対応する場合、ここに追加するだけでOK
NAME_ATTRIBUTE_MAPPING = {
    'company': ['ComName', 'company_name', '会社名', '企業名'],
    'department': ['DepName', 'department_name', '部署名'],
    'full_name': ['Name', 'inquiry_name', 'お名前', '氏名'],
    'name_last': ['姓', 'last_name', 'sei'],
    'name_first': ['名', 'first_name', 'mei'],
    'furigana_last': ['セイ', 'kana_last', 'sei_kana'],
    'furigana_first': ['メイ', 'kana_first', 'mei_kana'],
    'full_furigana': ['KanaName', 'kana_name', 'フリガナ'],
    'email': ['EMAIL', 'email', 'inquiry_email', 'メール'],
    'email_confirm': ['email2', 'confirm_email', 'メール確認'],
    'tel': ['telephone', 'tel', 'inquiry_tel', '電話', 'お電話'],
    'postal_code_1': ['zipcode1', 'zip1', '郵便番号1'],
    'postal_code_2': ['zipcode2', 'zip2', '郵便番号2'],
    'address': ['address', 'Prefecture', '住所'],
    'address_building': ['building', 'ビル名'],
    'inquiry_body': ['text', 'inquiry_content', 'message', 'お問い合わせ'],
}

# ============================================================
# 除外ルール設定
# ============================================================
# 特定のフィールドに誤って入力されないようにするルール
# 例: 「名」フィールドが「ビル名」に入力されないようにする
EXCLUSION_RULES = {
    'name_first': {
        # name属性にこれらが含まれていたら除外
        'name_contains': ['building', 'address', 'town', 'city', 'street', 'prefecture', 'zip'],
        # label/placeholderにこれらが含まれていたら除外
        'context_contains': ['ビル', '建物', '町', '番地', '住所', '都道府県', '市区町村', '郵便', '社', '店', '企業']
    },
    'name_last': {
        'name_contains': ['building', 'address', 'town', 'city', 'street', 'prefecture', 'zip'],
        'context_contains': ['ビル', '建物', '町', '番地', '住所', '都道府県', '市区町村', '郵便']
    },
    'subject': {
        # 件名が氏名欄に入らないようにする
        'context_contains': ['お名前', '氏名', '名前', '担当者名', 'ご担当者']
    },
    'full_name': {
        # 氏名が件名欄に入らないようにする
        'context_contains': ['件名', '題名', 'subject', 'タイトル', '用件']
    }
}

# ============================================================
# グローバル変数
# ============================================================
driver = None  # Seleniumのブラウザドライバー
url_input_timer = None  # URL入力のタイマー
browser_check_timer = None  # ブラウザ状態チェックのタイマー
status_label = None  # ステータス表示用のラベル

# ============================================================
# ヘルパー関数
# ============================================================

def is_browser_alive(driver_instance):
    """
    ブラウザが開いているかチェック
    手動でウィンドウを閉じた場合にFalseを返す
    """
    if not driver_instance:
        return False
    try:
        _ = driver_instance.title
        return True
    except Exception:
        return False

def check_name_attribute_match(key, field_name):
    """
    name属性がマッピングパターンに一致するかチェック
    """
    if key in NAME_ATTRIBUTE_MAPPING:
        return field_name in NAME_ATTRIBUTE_MAPPING[key]
    return False

def should_exclude_field(key, field, full_context):
    """
    除外ルールに基づいてフィールドをスキップすべきかチェック
    """
    if key not in EXCLUSION_RULES:
        return False
    
    rules = EXCLUSION_RULES[key]
    field_name_lower = field['name'].lower()
    
    if 'name_contains' in rules:
        for exclude_pattern in rules['name_contains']:
            if exclude_pattern in field_name_lower:
                return True
    
    if 'context_contains' in rules:
        for exclude_pattern in rules['context_contains']:
            if exclude_pattern in full_context:
                return True
    
    return False

def find_label_text_for_element(driver, element):
    """
    入力要素に対応するラベルテキストを見つける
    """
    try:
        element_id = element.get_attribute('id')
        if element_id:
            labels = driver.find_elements(By.XPATH, f"//label[@for='{element_id}']")
            if labels:
                text = driver.execute_script("return arguments[0].textContent;", labels[0])
                if text and text.strip():
                    return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()

        ancestor_labels = element.find_elements(By.XPATH, "./ancestor::label[1]")
        if ancestor_labels:
            text = driver.execute_script(
                "var e = arguments[0].cloneNode(true); var c = e.querySelector('input, textarea, select'); if(c){c.remove();} return e.textContent.trim();", 
                ancestor_labels[0]
            )
            if text and text.strip():
                return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()

        container = element.find_elements(By.XPATH, "./ancestor::div[contains(@class, 'form-group')][1]")
        if container:
            labels = container[0].find_elements(By.XPATH, ".//label")
            if labels:
                text = driver.execute_script("return arguments[0].textContent;", labels[0])
                if text and text.strip():
                    return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()

        ancestor_cell = element.find_elements(By.XPATH, "./ancestor::td[1]")
        if ancestor_cell:
            label_cells = ancestor_cell[0].find_elements(By.XPATH, "./preceding-sibling::th[1] | ./preceding-sibling::td[1]")
            if label_cells:
                text = driver.execute_script("return arguments[0].textContent;", label_cells[0])
                if text and text.strip():
                    return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()
        
        container = element.find_elements(By.XPATH, "./ancestor::*[self::p or self::div or self::li][1]")
        if container:
            all_text = driver.execute_script(
                "var e = arguments[0].cloneNode(true); "
                "var children = e.querySelectorAll('span, input, textarea, select, div, p');"
                "for (var i = 0; i < children.length; i++) { children[i].remove(); }"
                "return e.textContent.trim();",
                container[0]
            )
            if all_text and all_text.strip():
                return all_text.strip().split('\n')[0].replace("必須", "").replace("*", "").replace("※", "").strip()

    except (NoSuchElementException, StaleElementReferenceException):
        pass

    return ""

def show_toast(parent, message):
    """
    自動で消える通知ウィンドウを表示
    """
    toast = tk.Toplevel(parent)
    toast.overrideredirect(True)
    toast.attributes("-alpha", 0.9)
    label = ttk.Label(toast, text=message, padding="10", background="#333", foreground="white", font=("Yu Gothic UI", 12))
    label.pack()

    parent_x, parent_y = parent.winfo_x(), parent.winfo_y()
    parent_w, parent_h = parent.winfo_width(), parent.winfo_height()
    toast.update_idletasks()
    toast_w, toast_h = toast.winfo_reqwidth(), toast.winfo_reqheight()
    x = parent_x + (parent_w // 2) - (toast_w // 2)
    y = parent_y + (parent_h // 2) - (toast_h // 2)
    toast.geometry(f"+{x}+{y}")
    
    toast.after(2000, toast.destroy)

def check_browser_status():
    """
    定期的にブラウザの状態をチェック
    """
    global driver, browser_check_timer
    
    if driver is not None:
        if not is_browser_alive(driver):
            if DEBUG_MODE:
                print("DEBUG: Browser window was closed manually. Cleaning up driver.")
            try:
                driver.quit()
            except:
                pass
            driver = None
    
    if browser_check_timer is not None:
        browser_check_timer = None

# ============================================================
# フォーム検出・収集関数
# ============================================================

def get_all_form_contexts(driver):
    """
    メインページとすべてのiframe内のフォームコンテキストを取得
    """
    contexts = []
    contexts.append({'driver': driver, 'name': 'main page'})
    
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    if DEBUG_MODE:
        print(f"\nDEBUG: Found {len(iframes)} iframe(s)")
    
    for i, iframe in enumerate(iframes):
        try:
            driver.switch_to.frame(iframe)
            contexts.append({'driver': driver, 'name': f'iframe {i+1}'})
            driver.switch_to.default_content()
        except Exception as e:
            if DEBUG_MODE:
                print(f"DEBUG: Could not access iframe {i+1}: {e}")
            driver.switch_to.default_content()
    
    return contexts

def collect_form_fields_from_context(driver, context_name):
    """
    特定のコンテキストからフォームフィールドを収集
    """
    form_fields = []
    try:
        all_elements = driver.find_elements(By.XPATH, 
            "//input[@type='text' or @type='email' or @type='tel' or @type='url' or not(@type) or @type=''] | //textarea | //select")
        
        for element in all_elements:
            try:
                is_visible = driver.execute_script(
                    "var el = arguments[0];"
                    "var style = window.getComputedStyle(el);"
                    "return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;",
                    element
                )
                
                if is_visible and element.is_enabled():
                    form_fields.append({
                        'element': element,
                        'label': find_label_text_for_element(driver, element),
                        'placeholder': element.get_attribute('placeholder') or "",
                        'name': element.get_attribute('name') or "",
                        'id': element.get_attribute('id') or "",
                        'context': context_name
                    })
            except (StaleElementReferenceException, NoSuchElementException):
                continue
                
    except Exception as e:
        if DEBUG_MODE:
            print(f"DEBUG: Error collecting fields from {context_name}: {e}")
    
    return form_fields

# ============================================================
# メイン自動化処理
# ============================================================

def start_automation(driver, target_url, root_window):
    """
    フォーム自動入力のメイン処理
    """
    try:
        driver.get(target_url)
        
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//input | //textarea | //select"))
            )
        except:
            time.sleep(3)
        
        time.sleep(2)
        
        filled_count = 0
        filled_elements = set()
        subject_field_found = False
        
        contexts = get_all_form_contexts(driver)
        
        if DEBUG_MODE:
            print(f"\n--- DEBUG INFO: Scanning {len(contexts)} context(s) ---")
        
        all_form_fields = []
        for ctx in contexts:
            if ctx['name'] != 'main page':
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                iframe_index = int(ctx['name'].split()[-1]) - 1
                try:
                    driver.switch_to.frame(iframes[iframe_index])
                except:
                    continue
            
            fields = collect_form_fields_from_context(driver, ctx['name'])
            all_form_fields.extend(fields)
            
            driver.switch_to.default_content()
        
        if DEBUG_MODE:
            print("\n--- DEBUG INFO: Found Form Fields ---")
            for i, field in enumerate(all_form_fields):
                print(f"Field {i+1} [{field['context']}]: Label='{field['label']}', Placeholder='{field['placeholder']}', Name='{field['name']}', ID='{field['id']}'")
            print("------------------------------------\n")
        
        if len(all_form_fields) == 0:
            messagebox.showwarning("警告", "入力可能なフォーム要素が見つかりませんでした。")
            return
        
        priority_keys = ["company", "company_furigana", "department", "subject", "full_name", "name_last", "name_first", "furigana_last", "furigana_first", "full_furigana", "tel", "postal_code_1", "postal_code_2", "postal_code", "prefecture", "city", "address", "address_building", "email", "email_confirm", "url", "inquiry_body"]
        
        for key in priority_keys:
            if key not in MY_DATA:
                continue
            
            for field in all_form_fields:
                element = field['element']
                if element in filled_elements:
                    continue
                
                if field['context'] != 'main page':
                    iframe_index = int(field['context'].split()[-1]) - 1
                    iframes = driver.find_elements(By.TAG_NAME, "iframe")
                    try:
                        driver.switch_to.frame(iframes[iframe_index])
                    except:
                        driver.switch_to.default_content()
                        continue
                
                is_match = False
                full_context = field['label'] + field['placeholder'] + field['name'] + field['id']
                
                if check_name_attribute_match(key, field['name']):
                    is_match = True
                
                if not is_match:
                    for kw in KEYWORDS.get(key, []):
                        if kw in full_context:
                            if should_exclude_field(key, field, full_context): continue
                            if key == 'email' and any(k in full_context for k in ['確認', 'Confirm', 'confirm']): continue
                            if key == 'email_confirm' and not any(k in full_context for k in ['確認', 'Confirm', 'confirm']): continue
                            if key == 'company_furigana' and any(neg_kw in full_context for neg_kw in ["姓", "名", "氏名", "セイ", "メイ"]): continue
                            if key in ['furigana_last', 'furigana_first'] and any(neg_kw in full_context for neg_kw in ["会社", "企業", "法人", "カイシャ", "ホウジン"]): continue
                            if key in ['full_name', 'name_last', 'name_first'] and any(furi_kw in full_context for furi_kw in ["フリガナ", "ふりがな", "カナ", "セイ", "メイ", "Kana", "kana", "フリ", "ふり"]): continue
                            if key in ['furigana_last', 'furigana_first', 'full_furigana']:
                                if field['name'] in ['セイ', 'メイ']:
                                    is_match = True
                                    break
                                if not any(furi_kw in full_context for furi_kw in ["フリガナ", "ふりがな", "カナ", "セイ", "メイ", "Kana", "kana", "フリ"]): continue
                            if key == 'company_furigana' and not any(furi_kw in full_context for furi_kw in ["フリガナ", "ふりがな", "カナ", "カイシャメイ", "ホウジンメイ", "Kana", "kana", "フリ"]): continue
                            if key == 'postal_code_1' and ('-last' in full_context or 'last' in field['name']): continue
                            if key == 'postal_code_2' and ('-first' in full_context or 'first' in field['name']): continue
                            
                            is_match = True
                            break

                if is_match:
                    try:
                        text_to_send = MY_DATA[key]
                        
                        if key in ['tel', 'postal_code', 'postal_code_1', 'postal_code_2'] and 'ハイフンなし' in field['label']:
                            text_to_send = text_to_send.replace('-', '')
                        if key in ['postal_code_1', 'postal_code_2']:
                            text_to_send = text_to_send.replace('-', '')
                        
                        if key == 'inquiry_body' and not subject_field_found:
                            text_to_send = f"{MY_DATA.get('subject', '')}\n\n{MY_DATA[key]}"
                        
                        if element.is_displayed() and element.is_enabled():
                            if element.tag_name == 'select':
                                Select(element).select_by_visible_text(text_to_send)
                            else:
                                try:
                                    element.clear()
                                    element.send_keys(text_to_send)
                                except:
                                    driver.execute_script("arguments[0].value = arguments[1];", element, text_to_send)
                            
                            if key == 'subject': subject_field_found = True
                            filled_elements.add(element)
                            filled_count += 1
                            
                            if DEBUG_MODE:
                                print(f"DEBUG: Filled '{key}' in {field['context']}")
                        
                        driver.switch_to.default_content()
                        break
                    except Exception as e:
                        if DEBUG_MODE:
                            print(f"DEBUG: Error filling field '{key}': {e}")
                        driver.switch_to.default_content()
                        pass
                
                driver.switch_to.default_content()
        
        for ctx in contexts:
            if ctx['name'] != 'main page':
                iframe_index = int(ctx['name'].split()[-1]) - 1
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                try:
                    driver.switch_to.frame(iframes[iframe_index])
                except:
                    continue
            
            try:
                checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
                for checkbox in checkboxes:
                    try:
                        parent = checkbox.find_element(By.XPATH, "./ancestor::*[self::label or self::div or self::p][1]")
                        context_text = driver.execute_script("return arguments[0].textContent;", parent)
                        
                        if any(kw in context_text for kw in ["同意", "プライバシー", "個人情報", "利用規約", "承諾"]):
                            if checkbox.is_displayed() and checkbox.is_enabled() and not checkbox.is_selected():
                                driver.execute_script("arguments[0].click();", checkbox)
                                filled_count += 1
                                if DEBUG_MODE:
                                    print(f"DEBUG: Checked agreement checkbox")
                                break
                    except:
                        continue
            except Exception:
                pass
            
            driver.switch_to.default_content()

        if filled_count > 0:
            show_toast(root_window, f"{filled_count}個の項目を自動入力しました。")
        else:
            messagebox.showwarning("警告", "自動入力できる項目が見つかりませんでした。")

    except Exception as e:
        driver.switch_to.default_content()
        messagebox.showerror("エラー", f"処理中にエラーが発生しました。\n\n{e}")

# ============================================================
# GUI（ユーザーインターフェース）
# ============================================================

def main_gui():
    """
    メインのUIウィンドウを作成・表示
    """
    global browser_check_timer, status_label
    
    INITIAL_STATUS_TEXT = "準備完了"

    def trigger_automation_from_click(event=None):
        """クリックイベントで自動化処理を開始"""
        global driver
        
        url = ""
        try:
            url = root.clipboard_get()
            if not url.startswith(("http://", "https://")):
                status_label.config(text="クリップボードに有効なURLがありません。")
                root.after(3000, lambda: status_label.config(text=INITIAL_STATUS_TEXT))
                return
        except tk.TclError:
            status_label.config(text="クリップボードが空か、URLではありません。")
            root.after(3000, lambda: status_label.config(text=INITIAL_STATUS_TEXT))
            return

        if DEBUG_MODE: print(f"\n--- Automation Triggered for URL: {url} ---")
        
        status_label.config(text="処理を開始します...")
        
        if not is_browser_alive(driver):
            if DEBUG_MODE: print("DEBUG: Browser is not alive. Setting driver to None.")
            driver = None
        
        # 処理中はUIを無効化
        clickable_frame.unbind("<Button-1>")
        clickable_label.unbind("<Button-1>")
        edit_button.config(state="disabled")
        root.config(cursor="wait")
        root.update_idletasks()
        
        if driver is None:
            if DEBUG_MODE: print("DEBUG: Driver is None. Creating a new browser window...")
            try:
                options = Options()
                options.add_experimental_option("detach", True)
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=options)
                if DEBUG_MODE: print("DEBUG: New browser window created successfully.")
                start_browser_check()
            except Exception as e:
                messagebox.showerror("WebDriverエラー", f"Chrome Driverの起動に失敗しました。\n\n{e}")
                # UIを再度有効化
                clickable_frame.bind("<Button-1>", trigger_automation_from_click)
                clickable_label.bind("<Button-1>", trigger_automation_from_click)
                edit_button.config(state="normal")
                root.config(cursor="")
                status_label.config(text="✗ エラーが発生しました")
                root.after(3000, lambda: status_label.config(text=INITIAL_STATUS_TEXT))
                return

        # 自動化処理を実行
        try:
            status_label.config(text="入力中...")
            start_automation(driver, url, root)
            status_label.config(text="✓ 入力完了")
        except Exception as e:
            status_label.config(text=f"✗ エラー: {e}")
        finally:
            # UIを再度有効化
            clickable_frame.bind("<Button-1>", trigger_automation_from_click)
            clickable_label.bind("<Button-1>", trigger_automation_from_click)
            edit_button.config(state="normal")
            root.config(cursor="")
            root.after(4000, lambda: status_label.config(text=INITIAL_STATUS_TEXT))
    
    def start_browser_check():
        """ブラウザの状態チェックを開始（5秒ごと）"""
        global browser_check_timer
        check_browser_status()
        browser_check_timer = root.after(5000, start_browser_check)
    
    def stop_browser_check():
        """ブラウザチェックを停止"""
        global browser_check_timer
        if browser_check_timer is not None:
            root.after_cancel(browser_check_timer)
            browser_check_timer = None

    def on_closing():
        """ウィンドウを閉じる時の処理"""
        global driver
        stop_browser_check()
        if driver is not None:
            try: driver.quit()
            except: pass
            driver = None
        root.destroy()

    def open_settings_window():
        """設定ウィンドウを開く（入力情報の編集）"""
        settings_win = tk.Toplevel(root)
        settings_win.title("入力情報の編集")
        settings_win.geometry("600x700")

        x, y, w, h = root.winfo_x(), root.winfo_y(), root.winfo_width(), root.winfo_height()
        sw, sh = 600, 700
        settings_win.geometry(f"{sw}x{sh}+{x + (w - sw)//2}+{y + (h - sh)//2}")
        
        main_frame = ttk.Frame(settings_win, padding="10")
        main_frame.pack(fill="both", expand=True)

        entries = {}
        fields = [
            ("company", "会社名"), ("company_furigana", "会社名（フリガナ）"),
            ("name_last", "姓"), ("name_first", "名"),
            ("furigana_last", "姓（フリガナ）"), ("furigana_first", "名（フリガナ）"),
            ("email", "メールアドレス"), ("tel", "電話番号"), 
            ("postal_code", "郵便番号"), ("prefecture", "都道府県"),
            ("city", "市区町村"), ("address", "住所（全体）"),
            ("address_building", "建物名など"), ("url", "ホームページURL"),
            ("department", "部署"), ("subject", "件名")
        ]
        
        for i, (key, text) in enumerate(fields):
            ttk.Label(main_frame, text=text + ":").grid(row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(main_frame, width=60)
            entry.grid(row=i, column=1, sticky="ew", pady=2)
            entry.insert(0, MY_DATA.get(key, ""))
            entries[key] = entry

        ttk.Label(main_frame, text="お問い合わせ内容:").grid(row=len(fields), column=0, sticky="nw", pady=5)
        inquiry_text = tk.Text(main_frame, width=60, height=15, wrap="word", relief="solid", bd=1)
        inquiry_text.grid(row=len(fields), column=1, sticky="ew", pady=5)
        inquiry_text.insert("1.0", MY_DATA.get("inquiry_body", ""))
        entries["inquiry_body"] = inquiry_text

        main_frame.columnconfigure(1, weight=1)

        def save_settings():
            for key, widget in entries.items():
                MY_DATA[key] = widget.get("1.0", "end-1c") if isinstance(widget, tk.Text) else widget.get()
            
            MY_DATA["full_name"] = f"{MY_DATA.get('name_last','')}{MY_DATA.get('name_first','')}"
            MY_DATA["full_furigana"] = f"{MY_DATA.get('furigana_last','')}{MY_DATA.get('furigana_first','')}"
            MY_DATA["email_confirm"] = MY_DATA.get("email", "")
            
            if "postal_code" in MY_DATA and "-" in MY_DATA["postal_code"]:
                parts = MY_DATA["postal_code"].split("-")
                if len(parts) == 2:
                    MY_DATA["postal_code_1"], MY_DATA["postal_code_2"] = parts
            
            settings_win.destroy()
            show_toast(root, "入力情報を保存しました。")

        save_button = ttk.Button(main_frame, text="保存して閉じる", command=save_settings)
        save_button.grid(row=len(fields)+1, column=0, columnspan=2, pady=15)

    # --- メインウィンドウのセットアップ ---
    root = tk.Tk()
    root.title("フォーム自動入力ツール")
    root.geometry("450x220")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.resizable(False, False)

    # --- スタイルの設定 ---
    style = ttk.Style()
    style.theme_use('clam')
    BG_COLOR = "#f2f2f2"
    CLICK_BG_COLOR = "#ffffff"
    TEXT_COLOR = "#333333"
    
    root.configure(bg=BG_COLOR)
    
    style.configure("TFrame", background=BG_COLOR)
    style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR, font=("Yu Gothic UI", 10))
    style.configure("TButton", font=("Yu Gothic UI", 10))
    style.configure("Status.TLabel", font=("Yu Gothic UI", 11), padding=(5,0))
    style.configure("Click.TLabel", background=CLICK_BG_COLOR, foreground=TEXT_COLOR, font=("Yu Gothic UI", 12))
    
    # --- メインフレーム ---
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill="both", expand=True)

    # --- クリックエリア ---
    clickable_frame = tk.Frame(main_frame, bg=CLICK_BG_COLOR, relief="solid", bd=1, cursor="hand2")
    clickable_frame.pack(fill="both", expand=True, pady=(0, 15))
    
    clickable_label_text = "📋 URLをコピーして、ここをクリック 📋"
    clickable_label = ttk.Label(clickable_frame, text=clickable_label_text, style="Click.TLabel", anchor="center")
    clickable_label.place(relx=0.5, rely=0.5, anchor="center")
    
    clickable_frame.bind("<Button-1>", trigger_automation_from_click)
    clickable_label.bind("<Button-1>", trigger_automation_from_click)

    # --- 下部エリア (ステータスと設定ボタン) ---
    bottom_frame = ttk.Frame(main_frame)
    bottom_frame.pack(fill="x", side="bottom")
    bottom_frame.columnconfigure(0, weight=1)

    # ステータスラベル
    status_label = ttk.Label(bottom_frame, text=INITIAL_STATUS_TEXT, style="Status.TLabel", anchor="w")
    status_label.grid(row=0, column=0, sticky="ew")

    # 設定ボタン
    edit_button = ttk.Button(bottom_frame, text="⚙️ 設定", command=open_settings_window)
    edit_button.grid(row=0, column=1, sticky="e")
    
    root.mainloop()

# ============================================================
# プログラムのエントリーポイント
# ============================================================
if __name__ == '__main__':
    main_gui()