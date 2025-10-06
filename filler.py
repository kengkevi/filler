import sys
import time
import json
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
# グローバル変数
# ============================================================
driver = None
browser_check_timer = None
status_label = None
DEBUG_MODE = True

# ============================================================
# あなたの情報（初期値）
# ============================================================
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
私どもがご提案するのは、単に作業が楽になる便利な道具ではありません。 スウトメールの作成、面談記録の要約、日報の準備といった日々の業務をAIに任せ、そこで生まれた大切な時間で、候補者様との面談や企業様へのご提案といった、事業の核となる活動に集中する。 これは、かけた費用以上に会社の「成約数」を大きく伸ばす、成長のための仕組みづくりです。
私どものサービスや導入事例については、こちらのページで詳しくご紹介しております。 

▼ai-supporters サービス紹介ページ 
https://ai-supporters-lp1.vercel.app

つきましては、まず貴社のお悩みや「こうなったら嬉しい」というお話を、一度オンラインで30分ほどお聞かせいただけないでしょうか。 
情報収集の場としても歓迎ですので、ぜひお気軽にご活用ください。 もしご興味をお持ちいただけましたら、ご都合の良い日時を2〜3つお知らせいただけますと幸いです。
ご連絡を心よりお待ちしております。

ai-supporters 橋本"""
}

# ============================================================
# ヘルパー関数 (Selenium & System)
# ============================================================

def load_patterns(filename="form_patterns.json"):
    """設定ファイル(JSON)を読み込む"""
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        messagebox.showerror("エラー", f"設定ファイル '{filename}' が見つかりません。")
        sys.exit()
    except json.JSONDecodeError:
        messagebox.showerror("エラー", f"設定ファイル '{filename}' の形式が正しくありません。")
        sys.exit()

def is_browser_alive(driver_instance):
    """ブラウザが開いているかチェック"""
    if not driver_instance:
        return False
    try:
        _ = driver_instance.title
        return True
    except Exception:
        return False

def katakana_to_hiragana(katakana_string):
    """カタカナをひらがなに変換する"""
    kata_to_hira = str.maketrans(
        "ァィゥェォッャュョヮカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポヴ",
        "ぁぃぅぇぉっゃゅょゎかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんがぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽゔ"
    )
    return katakana_string.translate(kata_to_hira)

def show_toast(parent, message):
    """自動で消える通知ウィンドウを表示"""
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

def find_label_text_for_element(driver, element):
    """入力要素に対応するラベルテキストを見つける（優先順位を考慮）"""
    # 優先順位順に探索メソッドをリスト化
    search_methods = [
        find_label_by_for_attribute,
        find_label_by_ancestor_label,
        find_label_by_table_row,
        find_label_by_definition_list,
        find_label_by_form_group,
        find_label_by_generic_container,
        find_label_by_preceding_sibling
    ]
    
    for method in search_methods:
        try:
            text = method(driver, element)
            if text:
                return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    return ""

# --- ラベル探索の個別メソッド ---
def find_label_by_for_attribute(driver, element):
    element_id = element.get_attribute('id')
    if element_id:
        labels = driver.find_elements(By.XPATH, f"//label[@for='{element_id}']")
        if labels:
            return driver.execute_script("return arguments[0].textContent;", labels[0])
    return None

def find_label_by_ancestor_label(driver, element):
    ancestor_labels = element.find_elements(By.XPATH, "./ancestor::label[1]")
    if ancestor_labels:
        return driver.execute_script(
            "var e = arguments[0].cloneNode(true); var c = e.querySelector('input, textarea, select'); if(c){c.remove();} return e.textContent.trim();", 
            ancestor_labels[0])
    return None

def find_label_by_table_row(driver, element):
    ancestor_row = element.find_elements(By.XPATH, "./ancestor::tr[1]")
    if ancestor_row:
        label_header = ancestor_row[0].find_elements(By.XPATH, ".//th")
        if label_header:
            return driver.execute_script("return arguments[0].textContent;", label_header[0])
    return None

def find_label_by_definition_list(driver, element):
    ancestor_dd = element.find_elements(By.XPATH, "./ancestor::dd[1]")
    if ancestor_dd:
        label_dt = ancestor_dd[0].find_elements(By.XPATH, "./preceding-sibling::dt[1]")
        if label_dt:
            return driver.execute_script("return arguments[0].textContent;", label_dt[0])
    return None

def find_label_by_form_group(driver, element):
    container = element.find_elements(By.XPATH, "./ancestor::div[contains(@class, 'form-group')][1]")
    if container:
        labels = container[0].find_elements(By.XPATH, ".//label")
        if labels:
            return driver.execute_script("return arguments[0].textContent;", labels[0])
    return None

def find_label_by_generic_container(driver, element):
    container = element.find_elements(By.XPATH, "./ancestor::*[self::p or self::div or self::li][1]")
    if container:
        all_text = driver.execute_script(
            "var e = arguments[0].cloneNode(true); var c = e.querySelectorAll('span, input, textarea, select, div, p'); for(var i=0; i<c.length; i++){c[i].remove();} return e.textContent.trim();",
            container[0])
        if all_text and all_text.strip():
            return all_text.strip().split('\n')[0]
    return None

def find_label_by_preceding_sibling(driver, element):
    preceding_sibling = element.find_elements(By.XPATH, "./preceding-sibling::*[self::p or self::span or self::div or self::br][1]")
    if preceding_sibling:
        if preceding_sibling[0].tag_name == 'br':
             preceding_sibling = element.find_elements(By.XPATH, "./preceding-sibling::*[self::p or self::span or self::div][1]")
        if preceding_sibling:
            text = driver.execute_script("return arguments[0].textContent;", preceding_sibling[0])
            if text and text.strip() and any(kw in text for kw in ["確認", "もう一度", "confirm"]):
                 return text
    return None

# ============================================================
# スコアリングシステム & 自動化処理
# ============================================================

def calculate_score(field_info, key, patterns):
    """フィールドが指定されたキーにどれだけ一致するかをスコアリングする"""
    score = 0
    key_patterns = patterns.get(key, {})
    
    # 除外ルールの適用 (最優先)
    for rule in key_patterns.get("exclusions", []):
        attr_to_check = rule["attribute"]
        attr_value = field_info.get(attr_to_check, "").lower()
        for keyword in rule["contains"]:
            if keyword.lower() in attr_value:
                if DEBUG_MODE: print(f"DEBUG: Score reset to 0 for '{field_info['name']}' on key '{key}' due to exclusion rule '{keyword}'")
                return 0 # 除外ルールに一致したらスコアを0に
    
    # 属性ごとのチェック
    for attr, rules in key_patterns.get("attributes", {}).items():
        attr_value = field_info.get(attr, "").lower()
        if not attr_value:
            continue
        
        for keyword, points in rules.get("contains", {}).items():
            if keyword.lower() in attr_value:
                score += points
        for keyword, points in rules.get("exact", {}).items():
            if keyword.lower() == attr_value:
                score += points
    
    return score

def get_all_form_fields(driver):
    """ページとiframeから全てのフォームフィールド情報を収集する"""
    all_fields = []
    
    # メインコンテンツ
    if DEBUG_MODE: print("DEBUG: Collecting fields from main page.")
    all_fields.extend(collect_fields_from_current_context(driver))

    # IFrame
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    if DEBUG_MODE: print(f"DEBUG: Found {len(iframes)} iframe(s).")
    for i, iframe in enumerate(iframes):
        try:
            driver.switch_to.frame(iframe)
            if DEBUG_MODE: print(f"DEBUG: Switched to iframe {i+1}. Collecting fields.")
            all_fields.extend(collect_fields_from_current_context(driver))
            driver.switch_to.default_content()
        except Exception as e:
            if DEBUG_MODE: print(f"DEBUG: Could not access iframe {i+1}: {e}")
            driver.switch_to.default_content()
            
    return all_fields

def collect_fields_from_current_context(driver):
    """現在のコンテキスト（メインページ or iframe）からフィールドを収集"""
    fields = []
    elements = driver.find_elements(By.XPATH, "//input[@type='text' or @type='email' or @type='tel' or @type='url' or not(@type) or @type=''] | //textarea | //select")
    for el in elements:
        try:
            if el.is_displayed() and el.is_enabled():
                fields.append({
                    "element": el,
                    "label": find_label_text_for_element(driver, el),
                    "placeholder": el.get_attribute('placeholder') or "",
                    "name": el.get_attribute('name') or "",
                    "id": el.get_attribute('id') or "",
                    "type": el.get_attribute('type') or ""
                })
        except (StaleElementReferenceException, NoSuchElementException):
            continue
    return fields


def start_automation(driver, target_url, root_window):
    """フォーム自動入力のメイン処理"""
    patterns = load_patterns()
    if not patterns:
        return (False, "設定ファイルの読み込みに失敗")

    try:
        driver.get(target_url)
        WebDriverWait(driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')
        time.sleep(2) # 動的コンテンツの描画待ち

        all_form_fields = get_all_form_fields(driver)
        if DEBUG_MODE:
            print("\n--- DEBUG INFO: Found Form Fields ---")
            for i, field in enumerate(all_form_fields):
                print(f"Field {i+1}: Label='{field['label']}', Placeholder='{field['placeholder']}', Name='{field['name']}', ID='{field['id']}'")
            print("------------------------------------\n")

        if not all_form_fields:
            return (False, "入力可能なフォームが見つかりませんでした")

        filled_elements = set()
        filled_count = 0
        
        # 処理順序を定義
        priority_keys = patterns.get("priority", [])

        for key in priority_keys:
            if key not in MY_DATA and "hiragana" not in key:
                continue

            best_field = None
            highest_score = 0

            for field in all_form_fields:
                if field["element"] in filled_elements:
                    continue

                score = calculate_score(field, key, patterns)
                if score > highest_score:
                    highest_score = score
                    best_field = field

            if best_field and highest_score > patterns.get("score_threshold", 5):
                try:
                    element_to_fill = best_field["element"]
                    text_to_send = ""
                    
                    if "hiragana" in key:
                        original_key = key.replace("_hiragana", "_furigana")
                        if "_last" in original_key: original_key = "furigana_last"
                        elif "_first" in original_key: original_key = "furigana_first"
                        text_to_send = katakana_to_hiragana(MY_DATA.get(original_key, ""))
                    else:
                        text_to_send = MY_DATA.get(key, "")

                    full_context = best_field['label'] + best_field['placeholder'] + best_field['name']

                    if key in ['tel', 'postal_code']:
                        if not any(kw in full_context for kw in ["ハイフンあり", "ハイフンを入れて"]):
                            text_to_send = text_to_send.replace('-', '')
                    if key in ['postal_code_1', 'postal_code_2']:
                        text_to_send = text_to_send.replace('-', '')

                    if element_to_fill.tag_name == 'select':
                        Select(element_to_fill).select_by_visible_text(text_to_send)
                    else:
                        element_to_fill.clear()
                        element_to_fill.send_keys(text_to_send)

                    filled_elements.add(element_to_fill)
                    filled_count += 1
                    if DEBUG_MODE: print(f"DEBUG: Filled '{key}' with score {highest_score} in field (Name: {best_field['name']}, Label: {best_field['label']})")

                except Exception as e:
                    if DEBUG_MODE: print(f"DEBUG: Error filling field for '{key}': {e}")
        
        # 同意チェックボックスの処理
        try:
            checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
            for checkbox in checkboxes:
                try:
                    parent_container = checkbox.find_element(By.XPATH, "./ancestor::*[self::p or self::div or self::label or self::span][1]")
                    context_text = driver.execute_script("return arguments[0].textContent;", parent_container)
                    if any(kw in context_text for kw in ["同意", "プライバシー", "個人情報"]):
                        if checkbox.is_displayed() and checkbox.is_enabled() and not checkbox.is_selected():
                            driver.execute_script("arguments[0].click();", checkbox)
                            filled_count += 1
                            if DEBUG_MODE: print("DEBUG: Checked agreement checkbox.")
                            break
                except:
                    continue
        except Exception:
            pass

        if filled_count > 0:
            show_toast(root_window, f"{filled_count}個の項目を自動入力しました。")
            return (True, "入力完了")
        else:
            return (False, "自動入力できる項目がありませんでした")

    except Exception as e:
        if DEBUG_MODE:
            import traceback
            traceback.print_exc()
        return (False, "処理中にエラーが発生しました")


# ============================================================
# GUI（ユーザーインターフェース）
# ============================================================
def main_gui():
    global browser_check_timer, status_label
    
    INITIAL_STATUS_TEXT = "URLをコピーしてボタンをクリック"

    def trigger_automation_from_click(event=None):
        global driver
        
        try:
            url = root.clipboard_get()
            if not url.startswith(("http://", "https://")):
                status_label.config(text="✗ クリップボードにURLがありません", style="Error.TLabel")
                root.after(3000, lambda: status_label.config(text=INITIAL_STATUS_TEXT, style="Status.TLabel"))
                return
        except tk.TclError:
            status_label.config(text="✗ クリップボードが空です", style="Error.TLabel")
            root.after(3000, lambda: status_label.config(text=INITIAL_STATUS_TEXT, style="Status.TLabel"))
            return

        status_label.config(text="処理を開始します...", style="Status.TLabel")
        
        if not is_browser_alive(driver):
            driver = None
        
        canvas.unbind("<Enter>")
        canvas.unbind("<Leave>")
        canvas.unbind("<Button-1>")
        canvas.config(cursor="")
        edit_button.config(state="disabled")
        root.config(cursor="wait")
        root.update_idletasks()
        
        if driver is None:
            try:
                options = Options()
                options.add_experimental_option("detach", True)
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=options)
                start_browser_check()
            except Exception as e:
                status_label.config(text="✗ Chromeドライバの起動に失敗", style="Error.TLabel")
        
        if driver:
            try:
                status_label.config(text="入力中...", style="Status.TLabel")
                success, message = start_automation(driver, url, root)
                if success:
                    status_label.config(text=f"✓ {message}", style="Status.TLabel")
                else:
                    status_label.config(text=f"✗ {message}", style="Error.TLabel")
            except Exception as e:
                status_label.config(text="✗ 予期せぬエラーが発生", style="Error.TLabel")
        
        # UIを再有効化
        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)
        canvas.bind("<Button-1>", trigger_automation_from_click)
        canvas.config(cursor="hand2")
        edit_button.config(state="normal")
        root.config(cursor="")
        root.after(4000, lambda: status_label.config(text=INITIAL_STATUS_TEXT, style="Status.TLabel"))

    def start_browser_check():
        global browser_check_timer
        check_browser_status()
        browser_check_timer = root.after(5000, start_browser_check)
    
    def on_closing():
        global driver
        if browser_check_timer:
            root.after_cancel(browser_check_timer)
        if driver:
            try: driver.quit()
            except: pass
        root.destroy()
    
    def open_settings_window():
        settings_win = tk.Toplevel(root)
        settings_win.title("入力情報の編集")
        settings_win.geometry("600x700")
        
        main_frame = ttk.Frame(settings_win, padding="10")
        main_frame.pack(fill="both", expand=True)

        entries = {}
        fields = [("company", "会社名"), ("company_furigana", "会社名（フリガナ）"),
                  ("name_last", "姓"), ("name_first", "名"),
                  ("furigana_last", "姓（フリガナ）"), ("furigana_first", "名（フリガナ）"),
                  ("email", "メールアドレス"), ("tel", "電話番号"), 
                  ("postal_code", "郵便番号"), ("prefecture", "都道府県"),
                  ("city", "市区町村"), ("address", "住所（全体）"),
                  ("address_building", "建物名など"), ("url", "ホームページURL"),
                  ("department", "部署"), ("subject", "件名")]
        
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

    root = tk.Tk()
    root.title("フォーム自動入力ツール v2.3")
    root.geometry("320x320")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.resizable(False, False)

    style = ttk.Style()
    style.theme_use('clam')
    BG_COLOR, TEXT_COLOR = "#f2f2f2", "#333333"
    BUTTON_COLOR, BUTTON_HOVER_COLOR, BUTTON_TEXT_COLOR = "#4CAF50", "#45a049", "#FFFFFF"
    ERROR_COLOR = "#D8000C"
    
    root.configure(bg=BG_COLOR)
    style.configure("TFrame", background=BG_COLOR)
    style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR, font=("Yu Gothic UI", 10))
    style.configure("TButton", font=("Yu Gothic UI", 10))
    style.configure("Status.TLabel", font=("Yu Gothic UI", 11), padding=(5,0))
    style.configure("Error.TLabel", font=("Yu Gothic UI", 11), padding=(5,0), foreground=ERROR_COLOR)
    
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(main_frame, width=180, height=180, bg=BG_COLOR, bd=0, highlightthickness=0, cursor="hand2")
    canvas.pack(pady=20)
    button_oval = canvas.create_oval(5, 5, 175, 175, fill=BUTTON_COLOR, outline="")
    canvas.create_text(90, 90, text="自動入力\nスタート", fill=BUTTON_TEXT_COLOR, font=("Yu Gothic UI", 22, "bold"), justify="center")

    def on_enter(event): canvas.itemconfig(button_oval, fill=BUTTON_HOVER_COLOR)
    def on_leave(event): canvas.itemconfig(button_oval, fill=BUTTON_COLOR)
    canvas.bind("<Enter>", on_enter)
    canvas.bind("<Leave>", on_leave)
    canvas.bind("<Button-1>", trigger_automation_from_click)

    bottom_frame = ttk.Frame(main_frame)
    bottom_frame.pack(fill="x", side="bottom", pady=(10, 0))
    bottom_frame.columnconfigure(0, weight=1)

    status_label = ttk.Label(bottom_frame, text=INITIAL_STATUS_TEXT, style="Status.TLabel", anchor="w")
    status_label.grid(row=0, column=0, sticky="ew")

    edit_button = ttk.Button(bottom_frame, text="⚙️ 設定", command=open_settings_window)
    edit_button.grid(row=0, column=1, sticky="e")
    
    root.mainloop()

if __name__ == '__main__':
    main_gui()

