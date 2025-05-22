import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def setup_webdriver():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.implicitly_wait(10)
    return driver

def scrape_website(driver, url):
    data = {}
    try:
        driver.get(url)
        # タブのXPathを更新してクリック
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[1]/a'))).click()
        
        # スクレイピング内容を更新
        selectors = [
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr/th'),
            ("id", "rowSpanId3"),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[2]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[3]/th'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[3]/td'),
            ("id", "rowSpanId4"),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[5]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[4]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[6]/th'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[6]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[7]/th[2]'),
            ("xpath", '//*[@id="tableGroup-1"]/table[1]/tbody/tr[7]/td'),
            ("id", "rowSpanId5"),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[8]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[9]/td'),
            ("id", "rowSpanId6"),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[10]/th[2]'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[10]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[11]/th'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[11]/td'),
            ("xpath", '//div[@id="tableGroup-1"]/table/tbody/tr[12]/th'),
            ("xpath", '//*[@id="tableGroup-1"]/table[1]/tbody/tr[12]/td[2]'),
             ("id", "rowSpanId7"),
    ("xpath", "//div[@id='tableGroup-1']/table/tbody/tr[13]/th[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table/tbody/tr[13]/td"),
    ("xpath", "//div[@id='tableGroup-1']/table/tbody/tr[14]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table/tbody/tr[14]/td"),
    ("xpath", "//div[@id='tableGroup-1']/table/tbody/tr[15]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table/tbody/tr[15]/td"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[3]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[4]/th[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[4]/th[3]"),
    ("xpath", "//img[@alt='なし']"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[4]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[4]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[4]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[5]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[5]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[5]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[5]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[5]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[6]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[6]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[6]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[7]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[7]/th[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[7]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[7]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[7]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[7]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[8]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[8]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[8]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[8]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[8]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[9]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[9]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[9]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[10]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[10]/th[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[10]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[10]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[10]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[10]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[11]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[11]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[11]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[11]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[11]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[12]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[12]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[12]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[13]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[13]/th[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[13]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[13]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[13]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[13]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[14]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[14]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[14]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[14]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[14]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[15]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[15]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[15]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[16]/th"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[16]/th[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[16]/td/img"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[16]/td[2]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[16]/td[3]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[16]/td[4]"),
    ("xpath", "//div[@id='tableGroup-1']/table[2]/tbody/tr[17]/th")


        ]
        
        for i, (sel_type, selector) in enumerate(selectors, start=1):
            if sel_type == "xpath":
                elements = driver.find_elements(By.XPATH, selector)
            elif sel_type == "id":
                elements = [driver.find_element(By.ID, selector)]
            elif sel_type == "link":
                elements = [driver.find_element(By.LINK_TEXT, selector)]
            
            # elementsが複数ある場合でも処理できるようにリストで扱う
            for element in elements:
                if element.tag_name == "img":
                    # 要素が画像の場合はsrc属性を取得
                    data[f"key_{i}"] = element.get_attribute('src')
                    break  # 1つのキーに対して1つのURLのみを保存
                else:
                    # 画像でなければテキストを取得
                    data[f"key_{i}"] = element.text
                    break  # テキストが見つかればそれ以上の検索を停止
    
    except Exception as e:
        print(f"Error accessing {url}: {e}")
    
    return data
def main():
    excel_path = '事業所.xlsx'
    urls_df = pd.read_excel(excel_path, usecols=["URL"])
    driver = setup_webdriver()

    scraped_data_list = []

    for url in urls_df['URL']:
        scraped_row = scrape_website(driver, url)
        scraped_data_list.append(scraped_row)

    driver.quit()

    # リストをDataFrameに変換してExcelに書き出し
    scraped_data = pd.DataFrame(scraped_data_list)
    scraped_data.to_excel("scraped_data_result.xlsx", index=False)

if __name__ == "__main__":
    main()
