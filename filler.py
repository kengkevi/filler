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
    
    引数:
        key: データのキー（例: 'company', 'full_name'）
        field_name: フォームフィールドのname属性
    
    戻り値:
        一致すればTrue、しなければFalse
    """
    if key in NAME_ATTRIBUTE_MAPPING:
        return field_name in NAME_ATTRIBUTE_MAPPING[key]
    return False

def should_exclude_field(key, field, full_context):
    """
    除外ルールに基づいてフィールドをスキップすべきかチェック
    
    引数:
        key: データのキー
        field: フィールド情報（name属性含む）
        full_context: label + placeholder + name + id の結合文字列
    
    戻り値:
        除外すべきならTrue、入力してOKならFalse
    """
    if key not in EXCLUSION_RULES:
        return False
    
    rules = EXCLUSION_RULES[key]
    field_name_lower = field['name'].lower()
    
    # name属性チェック
    if 'name_contains' in rules:
        for exclude_pattern in rules['name_contains']:
            if exclude_pattern in field_name_lower:
                return True
    
    # context（label/placeholder等）チェック
    if 'context_contains' in rules:
        for exclude_pattern in rules['context_contains']:
            if exclude_pattern in full_context:
                return True
    
    return False

def find_label_text_for_element(driver, element):
    """
    入力要素に対応するラベルテキストを見つける
    様々なHTML構造に対応（label for, ancestor label, form-group, tableなど）
    """
    try:
        # パターン1: <label for="element_id">
        element_id = element.get_attribute('id')
        if element_id:
            labels = driver.find_elements(By.XPATH, f"//label[@for='{element_id}']")
            if labels:
                text = driver.execute_script("return arguments[0].textContent;", labels[0])
                if text and text.strip():
                    return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()

        # パターン2: <label>項目名<入力欄></label>
        ancestor_labels = element.find_elements(By.XPATH, "./ancestor::label[1]")
        if ancestor_labels:
            text = driver.execute_script(
                "var e = arguments[0].cloneNode(true); var c = e.querySelector('input, textarea, select'); if(c){c.remove();} return e.textContent.trim();", 
                ancestor_labels[0]
            )
            if text and text.strip():
                return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()

        # パターン3: Bootstrapのform-group構造
        container = element.find_elements(By.XPATH, "./ancestor::div[contains(@class, 'form-group')][1]")
        if container:
            labels = container[0].find_elements(By.XPATH, ".//label")
            if labels:
                text = driver.execute_script("return arguments[0].textContent;", labels[0])
                if text and text.strip():
                    return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()

        # パターン4: テーブル構造
        ancestor_cell = element.find_elements(By.XPATH, "./ancestor::td[1]")
        if ancestor_cell:
            label_cells = ancestor_cell[0].find_elements(By.XPATH, "./preceding-sibling::th[1] | ./preceding-sibling::td[1]")
            if label_cells:
                text = driver.execute_script("return arguments[0].textContent;", label_cells[0])
                if text and text.strip():
                    return text.strip().replace("必須", "").replace("*", "").replace("※", "").strip()
        
        # パターン5: 一般的なコンテナ構造
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
    2秒後に自動的に閉じます
    """
    toast = tk.Toplevel(parent)
    toast.overrideredirect(True)
    toast.attributes("-alpha", 0.9)
    label = ttk.Label(toast, text=message, padding="10", background="#333", foreground="white", font=("", 12))
    label.pack()

    # 親ウィンドウの中央に配置
    parent_x, parent_y = parent.winfo_x(), parent.winfo_y()
    parent_w, parent_h = parent.winfo_width(), parent.winfo_height()
    toast.update_idletasks()
    toast_w, toast_h = toast.winfo_reqwidth(), toast.winfo_reqheight()
    x = parent_x + (parent_w // 2) - (toast_w // 2)
    y = parent_y + (parent_h // 2) - (toast_h // 2)
    toast.geometry(f"+{x}+{y}")
    
    toast.after(2000, toast.destroy)

def animate_loading():
    """
    ローディングアニメーションを表示
    回転するスピナー文字で処理中であることを示す
    """
    global loading_animation_timer, loading_frame, status_label
    
    if status_label is None:
        return
    
    # 回転するスピナー文字
    spinner = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
    
    loading_frame = (loading_frame + 1) % len(spinner)
    status_label.config(text=f"{spinner[loading_frame]} 入力中...")
    
    # 100msごとに次のフレームを表示
    loading_animation_timer = status_label.after(100, animate_loading)

def stop_loading_animation():
    """ローディングアニメーションを停止"""
    global loading_animation_timer, status_label
    if loading_animation_timer is not None and status_label is not None:
        status_label.after_cancel(loading_animation_timer)
        loading_animation_timer = None

def check_browser_status():
    """
    定期的にブラウザの状態をチェック
    手動で閉じられた場合はクリーンアップを実行
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
    iframe内のフォームにも対応
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
    特定のコンテキスト（メインページまたはiframe）からフォームフィールドを収集
    表示されていて有効なフィールドのみを取得
    """
    form_fields = []
    try:
        all_elements = driver.find_elements(By.XPATH, 
            "//input[@type='text' or @type='email' or @type='tel' or @type='url' or not(@type) or @type=''] | //textarea | //select")
        
        for element in all_elements:
            try:
                # CSSで非表示になっているフィールドを除外
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
    
    処理の流れ:
    1. ページを開く
    2. フォームフィールドを検出
    3. 各フィールドに適切なデータを入力
    4. 同意チェックボックスをチェック
    """
    try:
        # ページを開く
        driver.get(target_url)
        
        # フォームの読み込みを待つ
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//input | //textarea | //select"))
            )
        except:
            time.sleep(3)  # iframeの読み込みを待つ
        
        time.sleep(2)  # 動的生成されるフォームを待つ
        
        filled_count = 0  # 入力したフィールド数
        filled_elements = set()  # 既に入力済みのフィールド（重複防止）
        subject_field_found = False  # 件名フィールドが見つかったか
        
        # すべてのコンテキスト（メインページ + iframe）を取得
        contexts = get_all_form_contexts(driver)
        
        if DEBUG_MODE:
            print(f"\n--- DEBUG INFO: Scanning {len(contexts)} context(s) ---")
        
        # 各コンテキストからフォームフィールドを収集
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
        
        # デバッグ情報を出力
        if DEBUG_MODE:
            print("\n--- DEBUG INFO: Found Form Fields ---")
            for i, field in enumerate(all_form_fields):
                print(f"Field {i+1} [{field['context']}]: Label='{field['label']}', Placeholder='{field['placeholder']}', Name='{field['name']}', ID='{field['id']}'")
            print("------------------------------------\n")
        
        if len(all_form_fields) == 0:
            messagebox.showwarning("警告", "入力可能なフォーム要素が見つかりませんでした。")
            return
        
        # 入力の優先順位（この順番で処理されます）
        priority_keys = ["company", "company_furigana", "department", "subject", "full_name", "name_last", "name_first", "furigana_last", "furigana_first", "full_furigana", "tel", "postal_code_1", "postal_code_2", "postal_code", "prefecture", "city", "address", "address_building", "email", "email_confirm", "url", "inquiry_body"]
        
        # 各データ項目について処理
        for key in priority_keys:
            if key not in MY_DATA: 
                continue
            
            # 各フォームフィールドをチェック
            for field in all_form_fields:
                element = field['element']
                if element in filled_elements:  # 既に入力済みならスキップ
                    continue
                
                # フィールドのコンテキストに切り替え（iframe内の場合）
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
                
                # ステップ1: name属性の完全一致チェック（最優先）
                if check_name_attribute_match(key, field['name']):
                    is_match = True
                
                # ステップ2: キーワードマッチング
                if not is_match:
                    for kw in KEYWORDS[key]:
                        if kw in full_context:
                            # 除外ルールチェック
                            if should_exclude_field(key, field, full_context):
                                continue
                            
                            # その他の除外条件（互換性のため残す）
                            if key == 'email' and any(k in full_context for k in ['確認', 'Confirm', 'confirm']):
                                continue
                            if key == 'email_confirm' and not any(k in full_context for k in ['確認', 'Confirm', 'confirm']):
                                continue
                            if key == 'company_furigana' and any(neg_kw in full_context for neg_kw in ["姓", "名", "氏名", "セイ", "メイ"]):
                                continue
                            if key in ['furigana_last', 'furigana_first'] and any(neg_kw in full_context for neg_kw in ["会社", "企業", "法人", "カイシャ", "ホウジン"]):
                                continue
                            
                            # 漢字データがフリガナ欄に入らないようにする
                            if key in ['full_name', 'name_last', 'name_first']:
                                if any(furi_kw in full_context for furi_kw in ["フリガナ", "ふりがな", "カナ", "セイ", "メイ", "Kana", "kana", "フリ", "ふり"]):
                                    continue
                            
                            # フリガナフィールドの検出
                            if key in ['furigana_last', 'furigana_first', 'full_furigana']:
                                if field['name'] in ['セイ', 'メイ']:
                                    is_match = True
                                    break
                                if not any(furi_kw in full_context for furi_kw in ["フリガナ", "ふりがな", "カナ", "セイ", "メイ", "Kana", "kana", "フリ"]):
                                    continue
                            if key == 'company_furigana':
                                if not any(furi_kw in full_context for furi_kw in ["フリガナ", "ふりがな", "カナ", "カイシャメイ", "ホウジンメイ", "Kana", "kana", "フリ"]):
                                    continue
                            
                            # 郵便番号分割フィールド
                            if key == 'postal_code_1':
                                if '-last' in full_context or 'last' in field['name']:
                                    continue
                            if key == 'postal_code_2':
                                if '-first' in full_context or 'first' in field['name']:
                                    continue
                            
                            is_match = True
                            break

                # マッチした場合、データを入力
                if is_match:
                    try:
                        text_to_send = MY_DATA[key]
                        
                        # ハイフンなし対応
                        if key in ['tel', 'postal_code', 'postal_code_1', 'postal_code_2'] and 'ハイフンなし' in field['label']:
                            text_to_send = text_to_send.replace('-', '')
                        # 郵便番号分割の場合、ハイフンを除去
                        if key in ['postal_code_1', 'postal_code_2']:
                            text_to_send = text_to_send.replace('-', '')
                        
                        # お問い合わせ内容に件名を含める（件名フィールドがない場合）
                        if key == 'inquiry_body' and not subject_field_found:
                            text_to_send = f"{MY_DATA['subject']}\n\n{MY_DATA['inquiry_body']}"
                        
                        # 要素が有効か再確認
                        if element.is_displayed() and element.is_enabled():
                            # セレクトボックスの場合
                            if element.tag_name == 'select':
                                Select(element).select_by_visible_text(text_to_send)
                            else:
                                # 通常の入力フィールドの場合
                                try:
                                    element.clear()
                                    element.send_keys(text_to_send)
                                except:
                                    # send_keysが失敗した場合はJavaScriptで直接値を設定
                                    driver.execute_script("arguments[0].value = arguments[1];", element, text_to_send)
                            
                            if key == 'subject': 
                                subject_field_found = True
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
        
        # 同意チェックボックスの処理
        for ctx in contexts:
            if ctx['name'] != 'main page':
                iframe_index = int(ctx['name'].split()[-1]) - 1
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                try:
                    driver.switch_to.frame(iframes[iframe_index])
                except:
                    continue
            
            try:
                # すべてのチェックボックスを取得
                checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
                for checkbox in checkboxes:
                    try:
                        # チェックボックスの近くのテキストを取得
                        parent = checkbox.find_element(By.XPATH, "./ancestor::*[self::label or self::div or self::p][1]")
                        context_text = driver.execute_script("return arguments[0].textContent;", parent)
                        
                        # 同意系のキーワードをチェック
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

        # 結果を表示
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
    ユーザーがURLを入力するとフォーム自動入力が開始されます
    """
    global browser_check_timer, status_label
    
    def run_automation_from_ui():
        """URL入力時に呼ばれる自動化処理のトリガー"""
        global driver
        
        if DEBUG_MODE: print("\n--- Automation Triggered ---")
        
        # ステータス更新（シンプルに）
        status_label.config(text="入力中...")
        
        if not is_browser_alive(driver):
            if DEBUG_MODE: print("DEBUG: Browser is not alive. Setting driver to None.")
            driver = None
        
        url = url_entry.get()
        if not url.startswith("http"): 
            status_label.config(text="URLを入力またはペーストしてください")
            return
        
        # ボタンを無効化（処理中は操作不可）
        edit_button.config(state="disabled")
        url_entry.config(state="disabled")
        root.update_idletasks()
        
        # ブラウザがまだ起動していない場合は起動
        if driver is None:
            if DEBUG_MODE: print("DEBUG: Driver is None. Creating a new browser window...")
            try:
                options = Options()
                options.add_experimental_option("detach", True)  # ブラウザを独立させる
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=options)
                if DEBUG_MODE: print("DEBUG: New browser window created successfully.")
                
                # ブラウザの状態チェックを開始
                start_browser_check()
            except Exception as e:
                messagebox.showerror("WebDriverエラー", f"Chrome Driverの起動に失敗しました。\n\n{e}")
                edit_button.config(state="normal")
                url_entry.config(state="normal")
                status_label.config(text="✗ エラーが発生しました")
                root.after(3000, lambda: status_label.config(text="URLを入力またはペーストしてください"))
                return

        # 自動化処理を実行
        try:
            start_automation(driver, url, root)
            status_label.config(text="✓ 入力完了")
        except Exception as e:
            status_label.config(text="✗ エラーが発生しました")
        finally:
            # ボタンを再度有効化
            edit_button.config(state="normal")
            url_entry.config(state="normal")
            url_entry.delete(0, tk.END)
            # 3秒後にステータスをリセット
            root.after(3000, lambda: status_label.config(text="URLを入力またはペーストしてください"))
    
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

    def schedule_run(event=None):
        """
        タイピングが止まったら自動実行をスケジュール
        ただし、長いURLがペーストされた場合は即座に実行
        """
        global url_input_timer
        
        # 現在のURL入力内容を取得
        current_url = url_entry.get()
        
        # URLが完全（http/httpsで始まり、それなりの長さがある）場合は即座に実行
        if current_url.startswith("http") and len(current_url) > 15:
            if url_input_timer:
                root.after_cancel(url_input_timer)
            # 即座に実行
            run_automation_from_ui()
        else:
            # 手入力の場合は0.5秒待機（タイピング中の誤動作防止）
            if url_input_timer:
                root.after_cancel(url_input_timer)
            url_input_timer = root.after(500, run_automation_from_ui)

    def on_closing():
        """ウィンドウを閉じる時の処理"""
        global driver, browser_check_timer
        
        # ブラウザチェックを停止
        stop_browser_check()
        
        # ブラウザを終了
        if driver is not None:
            try:
                driver.quit()
            except:
                pass
            driver = None
        
        root.destroy()

    def on_window_focus(event=None):
        """
        ウィンドウがフォーカスを得たときの処理
        クリップボードにURLがあれば自動ペースト
        """
        if DEBUG_MODE:
            print("DEBUG: on_window_focus called")
        
        try:
            # url_entryが存在するか確認
            if not hasattr(on_window_focus, 'url_entry_ready'):
                if DEBUG_MODE:
                    print("DEBUG: url_entry not ready yet")
                return
            
            # 入力欄が空の場合のみ実行
            current_value = url_entry.get().strip()
            if DEBUG_MODE:
                print(f"DEBUG: Current url_entry value: '{current_value}'")
            
            if current_value == "":
                # クリップボードの内容を取得
                clipboard_content = root.clipboard_get()
                if DEBUG_MODE:
                    print(f"DEBUG: Clipboard content: '{clipboard_content}'")
                
                # URLっぽい文字列か確認
                if clipboard_content.startswith(("http://", "https://")) and len(clipboard_content) > 10:
                    url_entry.delete(0, tk.END)
                    url_entry.insert(0, clipboard_content)
                    if DEBUG_MODE:
                        print(f"DEBUG: Auto-pasted URL from clipboard: {clipboard_content}")
                        print("DEBUG: About to call schedule_run()")
                    
                    # 自動ペースト後、処理を開始
                    try:
                        schedule_run()
                        if DEBUG_MODE:
                            print("DEBUG: schedule_run() called successfully")
                    except NameError as e:
                        if DEBUG_MODE:
                            print(f"DEBUG: NameError - schedule_run not found: {e}")
                    except Exception as e:
                        if DEBUG_MODE:
                            print(f"DEBUG: Error calling schedule_run: {e}")
                            import traceback
                            traceback.print_exc()
                else:
                    if DEBUG_MODE:
                        print("DEBUG: Clipboard content is not a valid URL")
            else:
                if DEBUG_MODE:
                    print("DEBUG: url_entry is not empty, skipping auto-paste")
        except tk.TclError as e:
            # クリップボードが空か、テキストでない場合
            if DEBUG_MODE:
                print(f"DEBUG: TclError accessing clipboard: {e}")
        except Exception as e:
            if DEBUG_MODE:
                print(f"DEBUG: Error in on_window_focus: {e}")
                import traceback
                print("DEBUG: Full traceback:")
                traceback.print_exc()

    def open_settings_window():
        """設定ウィンドウを開く（入力情報の編集）"""
        settings_win = tk.Toplevel(root)
        settings_win.title("入力情報の編集")
        settings_win.geometry("600x700")

        # 親ウィンドウの中央に配置
        x, y, w, h = root.winfo_x(), root.winfo_y(), root.winfo_width(), root.winfo_height()
        sw, sh = 600, 700
        settings_win.geometry(f"{sw}x{sh}+{x + (w - sw)//2}+{y + (h - sh)//2}")
        
        main_frame = ttk.Frame(settings_win, padding="10")
        main_frame.pack(fill="both", expand=True)

        entries = {}
        fields = [
            ("company", "会社名"), 
            ("company_furigana", "会社名（フリガナ）"),
            ("name_last", "姓"), 
            ("name_first", "名"),
            ("furigana_last", "姓（フリガナ）"), 
            ("furigana_first", "名（フリガナ）"),
            ("email", "メールアドレス"), 
            ("tel", "電話番号"), 
            ("postal_code", "郵便番号"),
            ("prefecture", "都道府県"),
            ("city", "市区町村"),
            ("address", "住所（全体）"),
            ("address_building", "建物名など"),
            ("url", "ホームページURL"),
            ("department", "部署"), 
            ("subject", "件名")
        ]
        
        # 各フィールドの入力欄を作成
        for i, (key, text) in enumerate(fields):
            label = ttk.Label(main_frame, text=text + ":")
            label.grid(row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(main_frame, width=60)
            entry.grid(row=i, column=1, sticky="ew", pady=2)
            entry.insert(0, MY_DATA.get(key, ""))
            entries[key] = entry

        # お問い合わせ内容（複数行テキスト）
        inquiry_label = ttk.Label(main_frame, text="お問い合わせ内容:")
        inquiry_label.grid(row=len(fields), column=0, sticky="nw", pady=5)
        inquiry_text = tk.Text(main_frame, width=60, height=15, wrap="word")
        inquiry_text.grid(row=len(fields), column=1, sticky="ew", pady=5)
        inquiry_text.insert("1.0", MY_DATA.get("inquiry_body", ""))
        entries["inquiry_body"] = inquiry_text

        main_frame.columnconfigure(1, weight=1)

        def save_settings():
            """設定を保存"""
            for key, widget in entries.items():
                MY_DATA[key] = widget.get("1.0", "end-1c") if isinstance(widget, tk.Text) else widget.get()
            
            # 自動生成される項目
            MY_DATA["full_name"] = f"{MY_DATA['name_last']}{MY_DATA['name_first']}"
            MY_DATA["full_furigana"] = f"{MY_DATA['furigana_last']}{MY_DATA['furigana_first']}"
            MY_DATA["email_confirm"] = MY_DATA["email"]
            
            # 郵便番号を分割
            if "postal_code" in MY_DATA and "-" in MY_DATA["postal_code"]:
                parts = MY_DATA["postal_code"].split("-")
                if len(parts) == 2:
                    MY_DATA["postal_code_1"] = parts[0]
                    MY_DATA["postal_code_2"] = parts[1]
            
            settings_win.destroy()
            show_toast(root, "入力情報を保存しました。")

        save_button = ttk.Button(main_frame, text="保存して閉じる", command=save_settings)
        save_button.grid(row=len(fields)+1, column=0, columnspan=2, pady=15)

    def on_window_focus(event=None):
        """
        ウィンドウがフォーカスを得たときの処理
        クリップボードにURLがあれば自動ペースト
        """
        try:
            # 入力欄が空の場合のみ実行
            if url_entry.get().strip() == "":
                # クリップボードの内容を取得
                clipboard_content = root.clipboard_get()
                
                # URLっぽい文字列か確認
                if clipboard_content.startswith(("http://", "https://")) and len(clipboard_content) > 10:
                    url_entry.delete(0, tk.END)
                    url_entry.insert(0, clipboard_content)
                    if DEBUG_MODE:
                        print(f"DEBUG: Auto-pasted URL from clipboard: {clipboard_content}")
        except tk.TclError:
            # クリップボードが空か、テキストでない場合
            pass
        except Exception as e:
            if DEBUG_MODE:
                print(f"DEBUG: Error accessing clipboard: {e}")
        """ウィンドウを閉じる時の処理"""
        global driver, browser_check_timer
        
        # ブラウザチェックを停止
        stop_browser_check()
        
        # ブラウザを終了
        if driver is not None:
            try:
                driver.quit()
            except:
                pass
            driver = None
        
        root.destroy()

    # メインウィンドウの作成
    root = tk.Tk()
    root.title("フォーム自動入力ツール")
    root.geometry("500x120")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # ウィンドウがフォーカスを得たときにクリップボードチェック
    root.bind("<FocusIn>", on_window_focus)

    frame = ttk.Frame(root, padding="10")
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(0, weight=1)

    # 設定ボタン
    edit_button = ttk.Button(frame, text="⚙️", command=open_settings_window, width=3)
    edit_button.grid(row=0, column=1, sticky="ne", padx=5, pady=5)

    # URLラベル
    url_label = ttk.Label(frame, text="対象URL:")
    url_label.grid(row=0, column=0, sticky="nw", padx=5, pady=5)

    # URL入力欄
    url_entry = ttk.Entry(frame)
    url_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
    url_entry.focus()
    url_entry.bind("<KeyRelease>", schedule_run)  # タイピングが止まったら自動実行
    
    # url_entryが準備できたことをマーク
    on_window_focus.url_entry_ready = True
    
    # ステータスラベル
    status_label = ttk.Label(frame, text="URLを入力またはペーストしてください")
    status_label.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=(10,5))
    
    root.mainloop()

# ============================================================
# プログラムのエントリーポイント
# ============================================================
if __name__ == '__main__':
    main_gui()