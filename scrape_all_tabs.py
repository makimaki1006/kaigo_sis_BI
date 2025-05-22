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


def scrape_basic_info(driver):
    """Scrape the 法人情報 tab assuming the page is already loaded."""
    data = {}
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[1]/a'))
        ).click()
        # Only a subset of fields are extracted here. Extend as needed.
        selectors = {
            "法人名": (By.XPATH, "//div[@id='tableGroup-1']/table/tbody/tr[2]/td"),
            "住所": (By.XPATH, "//div[@id='tableGroup-1']/table/tbody/tr[4]/td"),
        }
        for key, sel in selectors.items():
            element = driver.find_element(*sel)
            data[key] = element.text
    except Exception as e:
        print(f"Error scraping basic info: {e}")
    return data


def scrape_location(driver):
    """Scrape 所在地 tab."""
    data = {}
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[2]/a'))
        ).click()
        data["事業所の名称"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[3]/td").text
        data["郵便番号"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-3']/table/tbody/tr[4]/td").text
    except Exception as e:
        print(f"Error scraping location: {e}")
    return data


def scrape_staff(driver):
    """Scrape 従業者等 tab (truncated)."""
    data = {}
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[3]/a'))
        ).click()
        data["key_1"] = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr/th"))
        ).text
    except Exception as e:
        print(f"Error scraping staff: {e}")
    return data


def scrape_services(driver):
    """Scrape サービス内容 tab (truncated)."""
    data = {}
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[4]/a'))
        ).click()
        data["事業所の運営に関する方針"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-5']/table/tbody/tr[2]/td").text
    except Exception as e:
        print(f"Error scraping services: {e}")
    return data


def scrape_users(driver):
    """Scrape 利用者等 tab."""
    data = {}
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[5]/a'))
        ).click()
        data["交通費の額及び算定方法"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-6']/table/tbody/tr[3]/td").text
    except Exception as e:
        print(f"Error scraping users: {e}")
    return data


def main():
    excel_path = '事業所.xlsx'
    urls_df = pd.read_excel(excel_path, usecols=["URL"])
    driver = setup_webdriver()

    results = []
    for url in urls_df['URL']:
        driver.get(url)
        row = {}
        row.update(scrape_basic_info(driver))
        row.update(scrape_location(driver))
        row.update(scrape_staff(driver))
        row.update(scrape_services(driver))
        row.update(scrape_users(driver))
        results.append(row)

    driver.quit()

    df = pd.DataFrame(results)
    df.to_excel('scraped_data_all.xlsx', index=False)


if __name__ == '__main__':
    main()
