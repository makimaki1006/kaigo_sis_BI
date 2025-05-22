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
        # 指定されたタブをクリック
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[5]/a'))).click()

        # スクレイピングする項目の取得
        data["利用者の選定により、通常の事業の実施地域以外で当該介護サービスを行う場合、それに要する交通費の額及びその算定方法"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-6']/table/tbody/tr[3]/td").text
        image_element = driver.find_element(By.XPATH, "//div[@id='tableGroup-6']/table/tbody/tr[4]/td/img")
        data["利用者の都合により介護サービスを提供できなかった場合に係る費用（キャンセル料）の徴収状況"] = image_element.get_attribute('src') if image_element else None
        data["（その額、算定方法等）"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-6']/table/tbody/tr[5]/td").text

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
