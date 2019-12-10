from customar_list_gen import script_capture as cpt
from selenium import webdriver
import chromedriver_binary
import time

def browser_controller():
    results = []

    driver = webdriver.Chrome()
    driver.get("https://fumasalse.com/search/?search_from_top=1&tab_btn=on&tab_btn_menu1=on&chu_code%5B%5D=28&chu_code%5B%5D=29&chu_code%5B%5D=31&tab_btn_data=on&listed=1&jugyoinsu%5B%5D=5&jugyoinsu%5B%5D=6")

    time.sleep(3)

    #for i in range(5):
    url = driver.current_url
    time.sleep(1)
    result = cpt(url)
    for i in range(len(result)):
        results.append(result[i])
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    xp = '//*[@id="contents_main_box2"]/div[52]/div[1]/span/a[{}]'.format(3)
    driver.find_element_by_xpath(xp).click()
    time.sleep(3)

    url = driver.current_url
    time.sleep(1)
    result = cpt(url)
    for i in range(len(result)):
        results.append(result[i])
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    xp = '//*[@id="contents_main_box2"]/div[52]/div[1]/span/a[{}]'.format(5)
    driver.find_element_by_xpath(xp).click()
    time.sleep(3)

    url = driver.current_url
    time.sleep(1)
    result = cpt(url)
    for i in range(len(result)):
        results.append(result[i])
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    xp = '//*[@id="contents_main_box2"]/div[52]/div[1]/span/a[{}]'.format(6)
    driver.find_element_by_xpath(xp).click()
    time.sleep(3)

    url = driver.current_url
    time.sleep(1)
    result = cpt(url)
    for i in range(len(result)):
        results.append(result[i])
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    xp = '//*[@id="contents_main_box2"]/div[52]/div[1]/span/a[{}]'.format(6)
    driver.find_element_by_xpath(xp).click()
    time.sleep(3)

    url = driver.current_url
    time.sleep(1)
    result = cpt(url)
    for i in range(len(result)):
        results.append(result[i])

    #print(results)

    return results
