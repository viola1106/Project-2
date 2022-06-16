#!/usr/bin/env python
# coding: utf-8

# In[1]:


#爬取Agoda
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import openpyxl as op
import time
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, StaleElementReferenceException

book1=op.Workbook()
sh1=book1['Sheet']
sh1.title="20220501"
#如果要重跑可以用的
# book2=op.load_workbook("Agoda_台北_0501.xlsx")
# sh2=book2.active

sh1.append(["資料來源","飯店名稱","地址","入住日期","房型","價錢","整體狀況及整潔度","設施與設備","舒適程度","位置","員工素質與服務"])

def find_link(item_list):
    
    for x in range(4): #滾動頁面，取得每一頁所有的飯店Links
        item=driver.find_elements(By.XPATH,'//a[@class="PropertyCard__Link"]')
        time.sleep(2)
        driver.find_element_by_xpath('//body').send_keys(Keys.END)#滾動到頁面最下方
        item=driver.find_elements(By.XPATH,'//a[@class="PropertyCard__Link"]')
    item_list.extend(item)
    print("讀取成功")
    time.sleep(5)
    
    for item_data in item_list: #點進Links取得飯店資訊
        time.sleep(10)
            
        try:
            driver.execute_script('arguments[0].click()',item_data)
            time.sleep(2)
            
        #因為有開新分頁需要切換視窗，-1是最新的視窗
        num = driver.window_handles
        driver.switch_to.window(num[-1])

        except StaleElementReferenceException:
            continue

        time.sleep(10)#要給時間讓網頁資料顯示
        
        #取飯店名稱
        try:
            item1=driver.find_element_by_css_selector('h1[class="HeaderCerebrum__Name"]')
            name=item1.get_attribute("innerText")
            print(name)
            
        except TimeoutException:
            driver.execute_script("window.stop();")#網頁停止loading
            continue

        #地址
        item2=driver.find_element_by_css_selector('span[class="Spanstyled__SpanStyled-sc-16tp9kb-0 gwICfd kite-js-Span HeaderCerebrum__Address"]')
        address=item2.get_attribute("innerText")

        #房型
        try:
            if driver.find_element_by_css_selector('span[class="MasterRoom__HotelName"]').is_displayed():
                item3=driver.find_element_by_css_selector('span[class="MasterRoom__HotelName"]')
                bed=item3.text
                print(bed)

            elif driver.find_element_by_css_selector('span[class="MasterRoom-headerTitle--text"]').is_displayed():
                item3=driver.find_element_by_css_selector('span[class="MasterRoom-headerTitle--text"]')
                bed=item3.text
                print(bed)

        except NoSuchElementException: #排除沒有空房的飯店
            #建議 from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, StaleElementReferenceException
            driver.close()#關閉視窗
            driver.switch_to.window(num[0]) #返回原始視窗
            continue

        #價錢
        item4=driver.find_element_by_css_selector('span[class="Spanstyled__SpanStyled-sc-16tp9kb-0 gwICfd kite-js-Span pd-price PriceDisplay pd-color"]')
        price=item4.text
        if price.find(",")>0:
            price=price.replace(",","")
        print(price)

        #評價
        item5=driver.find_elements(By.CSS_SELECTOR,'div[class="Review-travelerGrade"]')
        score_list=[]
        for score in item5: 

            if score.get_attribute("innerText")[:-5]=="整體狀況及整潔度" or "設施與設備" or "客房舒適度" or "位置" or "服務" :
                comment2=score.get_attribute("innerText")[-3:]
                score_list.append(comment2)
            elif "整體狀況及整潔度" not in score.get_attribute("innerText") or "設施與設備" not in score.get_attribute("innerText") or "客房舒適度" not in score.get_attribute("innerText") or "位置" not in score.get_attribute("innerText") or "服務" not in score.get_attribute("innerText"):
                score_list.append("none")
 
        result=[]           
        result.extend(["Agoda",name,address_f,"2022-05-01",bed,price])
        result.extend(score_list)
        sh1.append(result)
        book1.save("Agoda_台北_0501.xlsx")
        time.sleep(3)

        driver.close() #關閉視窗
        driver.switch_to.window(num[0]) #返回查詢頁面
        time.sleep(5)
   
    #以上是取出一頁所有飯店資訊    
        
    driver.find_element_by_xpath('//body').send_keys(Keys.END)#滾動到查詢頁面最下面
    
    #按下一頁
    try:
        button1=driver.find_element(By.XPATH,'//button[@class="btn pagination2__next"]')
        button1.click()
        item_list.clear()
        find_link(item_list) #呼喚def繼續執行

    except NoSuchElementException:
        pass

    return item_list

desired_capabilities = DesiredCapabilities.CHROME  #調整selinium讀取方式
desired_capabilities["pageLoadStrategy"] = "none" #直接使用DOM樹，進行操作
driver=webdriver.Chrome("C:\\Users\\Lihome\\chromedriver.exe")
driver.get("https://www.agoda.com/zh-hk/search?city=4951&locale=zh-hk&ckuid=f97ea9a2-e509-4de5-a1aa-c728fb6593e7&prid=0&currency=TWD&correlationId=e70a84b0-6875-4311-a64d-a2cfb3df3650&pageTypeId=103&realLanguageId=7&languageId=7&origin=TW&cid=1844104&userId=f97ea9a2-e509-4de5-a1aa-c728fb6593e7&whitelabelid=1&loginLvl=0&storefrontId=3&currencyId=28&currencyCode=TWD&htmlLanguage=zh-hk&cultureInfoName=zh-hk&machineName=hk-acmweb-2007&trafficGroupId=1&sessionId=uq42ojyvxglel1twsrkeopil&trafficSubGroupId=84&aid=130589&useFullPageLogin=true&cttp=4&isRealUser=true&mode=production&checkIn=2022-05-01&checkOut=2022-05-02&rooms=1&adults=2&children=0&priceCur=TWD&los=1&textToSearch=%E5%8F%B0%E5%8C%97%E5%B8%82&productType=-1&travellerType=1&familyMode=off")

#點掉防疫諮詢的彈跳視窗
WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'button[class="Buttonstyled__ButtonStyled-sc-5gjk6l-0 kYHirW Box-sc-kv6pi1-0 hVPGaU"]'))).click()

#選擇住宿類型
driver.find_elements(By.CSS_SELECTOR,'button[class="btn PillDropdown__Button"]')[4].click()

#點飯店
driver.find_element(By.CSS_SELECTOR,'span[class="filter-item-info AccomdType-34 "]').click()

#點確認
try:
    button_1=driver.find_element(By.XPATH,'//*[@id="FilterBar"]/div[1]/div[5]/div/aside/footer/div/span/button[2]')
    driver.execute_script("arguments[0].click()",button_1) #排除按鈕被擋住的問題
    
except ElementClickInterceptedException:
     print("擋住了")
        
item_list=[]
find_link(item_list)

print("結束")

