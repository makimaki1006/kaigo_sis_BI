import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
import threading
import time

def element_path(element):
    """要素のパス（タグ名とクラスまたはID）を生成します。"""
    if not isinstance(element, Tag):
        return ""
    path = [element.name]
    if element.get('id'):
        path.append(f"id_{element['id']}")
    elif element.get('class'):
        path.append(f"class_{'_'.join(element['class'])}")
    return "_".join(path)

def extract_elements(element, path=""):
    """要素内の全てのテキストを含む要素を再帰的に抽出します。"""
    if element.name and element.text.strip():
        current_path = f"{path}_{element_path(element)}".strip("_")
        yield current_path, element.text.strip()
    for child in element.children:
        yield from extract_elements(child, path=current_path)

def process_row(index, row, lock):
    print(f"作業中 {index + 1}/{total_rows} ({(index + 1) * 100 / total_rows:.2f}%)")
    
    try:
        html_contents = requests.get(row["URL"], headers={'User-agent': 'Mozilla/5.0'}).text
        html_soup = BeautifulSoup(html_contents, "html.parser")
        
        tab_content = html_soup.find("div", {"id": "tab-2"})
        if tab_content:
            data = dict(extract_elements(tab_content))
                    
            with lock:
                for key, value in data.items():
                    # 新しい列を動的に追加
                    if key not in df.columns:
                        df[key] = None
                    df.loc[index, key] = value
        else:
            print(f"タブ 'tab-2' がURL {row['URL']} に見つかりません。")

    except Exception as e:
        print(f"Error processing row {index + 1}: {e}")

    time.sleep(0.1)  # サーバーへの負荷を考慮して待機

# Excelファイルを読み込む
df = pd.read_excel("事業所.xlsx", sheet_name="Sheet1")

# URL列が存在することを確認
if "URL" not in df.columns:
    raise ValueError("Excelに'URL'列が見つかりません。")

total_rows = len(df)
threads = []
lock = threading.Lock()

for index, row in df.iterrows():
    thread = threading.Thread(target=process_row, args=(index, row, lock))
    thread.start()
    threads.append(thread)

for thread in threads:
    thread.join()

# スクレイピングの結果を新しいExcelファイルに保存
df.to_excel("完了_事業所.xlsx", index=False)
