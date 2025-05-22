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
        
        data["key_1"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr/th").text
        data["key_2"] = driver.find_element(By.ID, "rowSpanId14").text
        data["key_3"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[3]/td").text
        data["key_4"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[2]/td").text
        data["key_5"] = driver.find_element(By.ID, "rowSpanId15").text
        data["key_6"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[4]/td").text
        data["key_7"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[4]/td[2]").text
        data["key_8"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[5]/td").text
        data["key_9"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[6]/td").text
        data["key_10"] = driver.find_element(By.ID, "rowSpanId16").text
        data["key_11"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[7]/th[2]").text
        data["key_12"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[7]/td").text
        data["key_13"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[8]/th").text
        data["key_14"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[8]/td").text
        data["key_15"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[9]/th").text
        data["key_16"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[9]/td[2]/a").get_attribute('href')
        data["key_17"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[10]/th").text
        data["key_18"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[10]/td").text
        data["key_19"] = driver.find_element(By.ID, "rowSpanId17").text
        data["key_20"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[11]/th[2]").text
        data["key_21"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[11]/td").text
        data["key_22"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[12]/th").text
        data["key_23"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[12]/td").text
        data["key_24"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[13]/th").text
        data["key_25"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[14]/th[2]").text
        data["key_26"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[14]/td").text
        data["key_27"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[15]/th").text
        data["key_28"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[15]/td").text
        data["key_29"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[16]/th").text
        data["key_30"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[16]/td").text
        data["key_31"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[17]/th").text
        # 画像のsrc属性を取得
        image_element = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[17]/td/img")
        data["key_32"] = image_element.get_attribute('src') if image_element else None
        data["key_33"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[18]/th").text
        data["key_34"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[19]/td").text
        data["key_35"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[20]/th").text
        data["key_36"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[20]/td").text
        
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
