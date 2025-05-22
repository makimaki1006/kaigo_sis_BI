from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

def setup_webdriver():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.implicitly_wait(10)
    return driver

def scrape_website(driver, url):
    data = {}
    try:
        driver.get(url)
        # タブのXPATHを更新してクリック
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[2]/a'))).click()
        
        data["事業所の名称"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[3]/td").text
        data["ふりがな"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[2]/td").text
        data["郵便番号"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[4]/td").text
        data["市町村"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[4]/td[2]").text
        data["住所(番地まで)"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[5]/td").text
        data["番地以降"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[6]/td").text
        data["電話番号"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[7]/td").text
        data["FAX番号"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[8]/td").text
        data["ホームページ"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[9]/td[2]/a").get_attribute('href')
        data["介護保険事業所番号"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[10]/td").text
        data["事業所の管理者の氏名"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[11]/td").text
        data["職名"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[12]/td").text
        data["事業の開始（予定）年月日"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[14]/td").text
        data["指定の年月日"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[15]/td").text
        data["指定の更新年月日（直近）"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[16]/td").text
        image_element = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[17]/td/img")
        data["生活保護法第54条の2に規定する介護機関（生活保護の介護扶助を行う機関）の指定"] = image_element.get_attribute('src') if image_element else None
        data["事業所までの主な利用交通手段"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[19]/td").text
        data["ケアプランデータ連携システム（国保中央会）の利用登録の有無"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[20]/td").text
        
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

    # データをDataFrameに変換してExcelに書き出し
    scraped_data = pd.DataFrame(scraped_data_list)
    scraped_data.to_excel("scraped_data_result.xlsx", index=False)

if __name__ == "__main__":
    main()
