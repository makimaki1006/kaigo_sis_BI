from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

def setup_webdriver():
    options = webdriver.ChromeOptions()
    # For headless mode, uncomment the next line
    # options.add_argument('--headless')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.implicitly_wait(10)
    return driver

def scrape_website(driver, url):
    data = {}
    try:
        driver.get(url)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="kihonItemTab"]/li[3]/a'))).click()    
        data["key_1"] = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr/th"))).text
        data["key_2"] = driver.find_element(By.ID,"rowSpanId20").text
        data["key_3"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/th").text
        data["key_4"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/td").text
        data["key_5"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/td[2]").text
        data["key_6"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/td[3]").text
        data["key_7"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/td[4]").text
        data["key_8"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/td[5]").text
        data["key_9"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[4]/td[6]").text
        data["key_10"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/th[2]").text
        data["key_11"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/td").text
        data["key_12"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/td[2]").text
        data["key_13"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/td[3]").text
        data["key_14"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/td[4]").text
        data["key_15"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/td[5]").text
        data["key_16"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[5]/td[6]").text
        data["key_17"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/th").text
        data["key_18"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/td").text
        data["key_19"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/td[2]").text
        data["key_20"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/td[3]").text
        data["key_21"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/td[4]").text
        data["key_22"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/td[5]").text
        data["key_23"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[6]/td[6]").text
        data["key_24"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/th").text
        data["key_25"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/td").text
        data["key_26"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/td[2]").text
        data["key_27"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/td[3]").text
        data["key_28"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/td[4]").text
        data["key_29"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/td[5]").text
        data["key_30"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[7]/td[6]").text
        data["key_31"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[8]/th").text
        data["key_32"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[8]/td").text
        data["key_33"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[9]/th").text
        data["key_34"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[10]/th").text
        data["key_35"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[10]/th[2]").text
        data["key_36"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[10]/td").text
        data["key_37"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[10]/th[3]").text
        data["key_38"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[10]/td[2]").text
        data["key_39"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[11]/th").text
        data["key_40"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[14]/th").text
        data["key_41"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[14]/td").text
        data["key_42"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[14]/td[2]").text
        data["key_43"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[14]/td[3]").text
        data["key_44"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[14]/td[4]").text
        data["key_45"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[15]/th").text
        data["key_46"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[15]/td").text
        data["key_47"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[15]/td[2]").text
        data["key_48"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[15]/td[3]").text
        data["key_49"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[15]/td[4]").text
        data["key_50"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[16]/th").text
        data["key_51"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[16]/td").text
        data["key_52"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[16]/td[2]").text
        data["key_53"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[16]/td[3]").text
        data["key_54"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[16]/td[4]").text
        data["key_55"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[17]/th").text
        data["key_56"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[17]/td").text
        data["key_57"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[17]/td[2]").text
        data["key_58"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[17]/td[3]").text
        data["key_59"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[17]/td[4]").text
        data["key_60"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[18]/th").text
        data["key_61"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[18]/td").text
        data["key_62"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[18]/td[2]").text
        data["key_63"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[18]/td[3]").text
        data["key_64"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[18]/td[4]").text
        data["key_65"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[19]/th").text
        data["key_66"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[19]/td").text
        data["key_67"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[19]/td[2]").text
        data["key_68"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[19]/td[3]").text
        data["key_69"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[19]/td[4]").text
        data["key_70"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[20]/th").text
        data["key_71"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[20]/td").text
        data["key_72"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[20]/td[2]").text
        data["key_73"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[20]/td[3]").text
        data["key_74"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[20]/td[4]").text
        data["key_75"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[21]/th").text
        data["key_76"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[21]/td").text
        data["key_77"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[21]/td[2]").text
        data["key_78"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[21]/td[3]").text
        data["key_79"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[21]/td[4]").text
        data["key_80"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[22]/th").text
        data["key_81"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[22]/td").text
        data["key_82"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[22]/td[2]").text
        data["key_83"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[22]/td[3]").text
        data["key_84"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[22]/td[4]").text
        data["key_85"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[23]/th").text
        data["key_86"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[23]/td").text
        data["key_87"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[23]/td[2]").text
        data["key_88"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[23]/td[3]").text
        data["key_89"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[23]/td[4]").text
        data["key_90"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[24]/th").text
        data["key_91"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[24]/td").text
        data["key_92"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[24]/td[2]").text
        data["key_93"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[24]/td[3]").text
        data["key_94"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[24]/td[4]").text
        data["key_95"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[25]/th").text
        data["key_96"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[25]/td").text
        data["key_97"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[25]/td[2]").text
        data["key_98"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[25]/td[3]").text
        data["key_99"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[25]/td[4]").text
        data["key_100"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[26]/th").text
        data["key_101"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[26]/td").text
        data["key_102"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[26]/td[2]").text
        data["key_103"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[26]/td[3]").text
        data["key_104"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[26]/td[4]").text
        data["key_105"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[27]/th").text
        data["key_106"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[27]/td").text
        data["key_107"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[27]/td[2]").text
        data["key_108"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[27]/td[3]").text
        data["key_109"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[27]/td[4]").text
        data["key_110"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[28]/th").text
        data["key_111"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[28]/td").text
        data["key_112"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[28]/td[2]").text
        data["key_113"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[28]/td[3]").text
        data["key_114"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[28]/td[4]").text
        data["key_115"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[29]/th").text
        data["key_116"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[29]/td").text
        data["key_117"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[29]/td[2]").text
        data["key_118"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[29]/td[3]").text
        data["key_119"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[29]/td[4]").text
        data["key_120"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[30]/th").text
        data["key_121"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[30]/td").text
        data["key_122"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[30]/td[2]").text
        data["key_123"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[30]/td[3]").text
        data["key_124"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[30]/td[4]").text
        data["key_125"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[31]/th").text
        data["key_126"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[31]/td").text
        data["key_127"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[31]/td[2]").text
        data["key_128"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[31]/td[3]").text
        data["key_129"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[31]/td[4]").text
        data["key_130"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[32]/th").text
        data["key_131"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[32]/td").text
        data["key_132"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[32]/td[2]").text
        data["key_133"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[32]/td[3]").text
        data["key_134"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[32]/td[4]").text
        data["key_135"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[33]/th").text
        data["key_136"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[33]/td").text
        data["key_137"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[33]/td[2]").text
        data["key_138"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[33]/td[3]").text
        data["key_139"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[33]/td[4]").text
        data["key_140"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[34]/th").text
        data["key_141"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[34]/td").text
        data["key_142"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[34]/td[2]").text
        data["key_143"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[34]/td[3]").text
        data["key_144"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[34]/td[4]").text
        data["key_145"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[35]/th").text
        data["key_146"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[35]/td").text
        data["key_147"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[35]/td[2]").text
        data["key_148"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[35]/td[3]").text
        data["key_149"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[35]/td[4]").text
        data["key_150"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[36]/th").text
        data["key_151"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[36]/td").text
        data["key_152"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[36]/td[2]").text
        data["key_153"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[36]/td[3]").text
        data["key_154"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[36]/td[4]").text
        data["key_155"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[37]/th").text
        data["key_156"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[37]/td").text
        data["key_157"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[37]/td[2]").text
        data["key_158"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[37]/td[3]").text
        data["key_159"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[37]/td[4]").text
        data["key_160"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[38]/th").text
        data["key_161"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[38]/td").text
        data["key_162"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[38]/td[2]").text
        data["key_163"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[38]/td[3]").text
        data["key_164"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[38]/td[4]").text
        data["key_165"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[39]/th").text
        data["key_166"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[39]/td/img").text
        data["key_167"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[40]/th").text
        data["key_168"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[40]/td/img").text
        data["key_169"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[41]/th[2]").text
        image_element_1 = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[41]/td/img")
        data["key_170"] = image_element_1.get_attribute('src') if image_element_1 else None.text
        data["key_171"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[42]/th[2]").text
        data["key_172"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[42]/td").text
        data["key_173"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[43]/th").text
        data["key_174"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[44]/th[3]").text
        data["key_175"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[46]/th").text
        data["key_176"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[45]/th").text
        data["key_177"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[46]/td").text
        data["key_178"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[45]/th[2]").text
        data["key_179"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[46]/td[2]").text
        data["key_180"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[47]/th").text
        data["key_181"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[45]/th").text
        data["key_182"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[47]/td").text
        data["key_183"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[45]/th[2]").text
        data["key_184"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[47]/td[2]").text
        data["key_185"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th").text
        data["key_186"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[49]/th[2]").text
        data["key_187"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[2]").text
        data["key_188"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[49]/td").text
        data["key_189"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[3]").text
        data["key_190"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[49]/td[2]").text
        data["key_191"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[50]/th").text
        data["key_192"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[2]").text
        data["key_193"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[50]/td").text
        data["key_194"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[3]").text
        data["key_195"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[50]/td[2]").text
        data["key_196"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[51]/th").text
        data["key_197"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[2]").text
        data["key_198"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[51]/td").text
        data["key_199"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[3]").text
        data["key_200"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[51]/td[2]").text
        data["key_201"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[52]/th").text
        data["key_202"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[2]").text
        data["key_203"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[52]/td").text
        data["key_204"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[3]").text
        data["key_205"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[52]/td[2]").text
        data["key_206"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[53]/th").text
        data["key_207"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[2]").text
        data["key_208"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[53]/td").text
        data["key_209"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[48]/th[3]").text
        data["key_210"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[53]/td[2]").text
        data["key_211"] = driver.find_element(By.ID,"colSpanId27").text
        data["key_212"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[55]/th").text
        data["key_213"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[55]/td/img").text
        data["key_214"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[56]/th").text
        data["key_215"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[57]/th[2]").text
        data["key_216"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[58]/td").text
        data["key_217"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[59]/th").text
        data["key_218"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[60]/th[2]").text
        data["key_219"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[60]/td").text
        data["key_220"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[61]/th").text
        data["key_221"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[61]/td").text
        data["key_222"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[62]/th").text
        data["key_223"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[62]/td").text
        data["key_224"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[63]/th").text
        data["key_225"] = driver.find_element(By.XPATH, "//div[@id='tableGroup-4']/table/tbody/tr[63]/td").text



        
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

    # Convert list to DataFrame and write to Excel
    scraped_data = pd.DataFrame(scraped_data_list)
    scraped_data.to_excel("scraped_data_results.xlsx", index=False)

if __name__ == "__main__":
    main()
