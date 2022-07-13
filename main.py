import base64
import hashlib
import json
import os
import time

import requests as requests
from selenium import webdriver

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import glob
import win32com.client as win32
import hmac

def send_key(xpath, keys):
    driver.find_element(By.XPATH, xpath).send_keys(keys)

def click(xpath):
    driver.find_element(By.XPATH, xpath).click()

def wait_until_clickable(time, xpath):
    WebDriverWait(driver, time).until(EC.element_to_be_clickable((By.XPATH, xpath)))

def mark_as_delivery_completed():
    click("//input[@type='checkbox']")
    click("//button[@id='btn_shpp_cmpl']")

    WebDriverWait(driver, 20).until(EC.alert_is_present())

    alert = driver.switch_to.alert
    alert.accept()

    WebDriverWait(driver, 20).until(EC.alert_is_present())

    alert = driver.switch_to.alert
    alert.accept()

def parse_order_data():
    path = "*.xls"

    while True:
        try:
            xls_filename = os.getcwd() + '\\' + glob.glob(path)[0]

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(xls_filename)

            wb.SaveAs(xls_filename + 'x', FileFormat=51)
            wb.Close()
            excel.Application.Quit()

            break
        except Exception as e:
            print(e)
            time.sleep(1)

    path = "*.xlsx"

    xlsx_filename = glob.glob(path)[0]
    excel = pd.read_excel(xlsx_filename)

    item = list()

    for i in range(len(excel)):
        temp_dict = dict()
        temp_dict['ordernum'] = str(excel['주문번호'].iloc[i])
        temp_dict['orderid'] = str(excel['회원ID'].iloc[i])
        temp_dict['ordername'] = str(excel['회원명'].iloc[i])
        temp_dict['pcode'] = str(excel['상품코드'].iloc[i])
        temp_dict['pname'] = str(excel['상품명'].iloc[i])
        temp_dict['price_sale'] = int(excel['판매가'].iloc[i])
        temp_dict['quant'] = int(excel['수량'].iloc[i])
        temp_dict['discount'] = int(excel['할인금액'].iloc[i])
        temp_dict['price_total'] = int(excel['결제금액'].iloc[i])
        temp_dict['orderhtel'] = str(excel['수취인 휴대폰번호'].iloc[i])
        temp_dict['recvname'] = str(excel['수취인명'].iloc[i])
        temp_dict['orderdate'] = str(excel['주문완료일시'].iloc[i])
        temp_dict['mall'] = str(excel['몰구분'].iloc[i])
        temp_dict['orderstatus'] = str(excel['주문상태'].iloc[i])
        temp_dict['msg'] = str(excel['발송메시지'].iloc[i])
        temp_dict['supplier'] = str(excel['공급업체'].iloc[i])

        item.append(temp_dict)

    print(item)
    os.remove(xls_filename)
    os.remove(xlsx_filename)

    return item

def download_excel():
    click("//a[@class='ico_excel']")

    while len(driver.window_handles) < 2:
        time.sleep(0.5)

    driver.switch_to.window(driver.window_handles[-1])
    wait_until_clickable(20, "//label[@for='piDwldPropTypeCd1']")

    click("//label[@for='piDwldPropTypeCd1']")
    click("//label[@for='piDwldPropTypeCd2']")
    click("//label[@for='piDwldPropTypeCd3']")
    click("//label[@for='piDwldPropTypeCd4']")

    send_key("//textarea[@id='plDwldRsnCntt']", '쿠폰발송')

    click("//label[@for='maskRlsYnCd2']")

    send_key("//textarea[@id='maskRlsRsnCntt']", '쿠폰발송')

    click("//button[@class='btn btn_ty12']")

    while len(driver.window_handles) >= 2:
        time.sleep(0.5)

    driver.switch_to.window(driver.window_handles[-1])
    driver.switch_to.frame('iframe_5000001207_1')

def is_there_order():
    wait_until_clickable(20, "//button[@id='searchBtn']")
    click("//button[@id='searchBtn']")

    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present())

        alert = driver.switch_to.alert
        alert.accept()

        log_in()
        move_to_mobile_gift_order()
        select_condition()

        return False
    except Exception as e:
        pass

    num_of_order = len(driver.find_elements(By.CLASS_NAME, 'objbox')[1].find_elements(By.CLASS_NAME, "ev_dhx_skyblue")) + \
            len(driver.find_elements(By.CLASS_NAME, 'objbox')[1].find_elements(By.CLASS_NAME, "odd_dhx_skyblue"))

    if num_of_order > 0:
        return True

    return False

def select_condition():
    while len(driver.window_handles) >= 2:
        driver.switch_to.window(driver.window_handles[-1])
        driver.close()
        time.sleep(0.5)
        driver.switch_to.window(driver.window_handles[-1])

    driver.switch_to.window(driver.window_handles[-1])

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//iframe[@id='iframe_5000001207_1']")))
    driver.switch_to.frame('iframe_5000001207_1')

    wait_until_clickable(20, "//button[@id='searchBtn']")

    select = Select(driver.find_element(By.ID, 'comboId'))
    select.select_by_visible_text('최근 3일')

    select = Select(driver.find_element(By.ID, 'lb_order_state'))
    select.select_by_visible_text('주문완료')


def move_to_mobile_gift_order():
    click("//button[@id='bookmarkBtn']")

    bookmark_clicked = False

    while not bookmark_clicked:
        elements = driver.find_elements(By.XPATH, "//span[@class='gnb_nav_tx']")

        for each in elements:
            if '모바일기프트 주문조회' in each.text:
                each.click()
                bookmark_clicked = True
                break

        time.sleep(0.5)

def log_in():
    print('Logging in...')

    driver.get(url=URL)

    try:
        wait_until_clickable(20, "//input[@id='userId']")
    except Exception as e:
        return

    time.sleep(1)

    print(len(driver.window_handles))
    send_key("//input[@id='userId']", ssg_id)
    send_key("//input[@id='userPwd']", ssg_pwd + Keys.ENTER)

    while len(driver.window_handles) < 2:
        time.sleep(0.5)
    driver.switch_to.window(driver.window_handles[-1])

    wait_until_clickable(20, "//button[@class='pop_login_sendbtn']")

    driver.find_elements(By.CLASS_NAME, 'pop_login_sendbtn')[1].click()

    while True:
        try:
            time.sleep(7)

            response = requests.get(url=sms_api_url, params={'key': sms_api_key})

            auth_code = response.json()
            print('Auth code : ' + str(auth_code['auth']))

            send_key("//input[@id='certNum1']", auth_code['auth'])
            click("//button[@onclick='certNum1();']")

            WebDriverWait(driver, 5).until(EC.alert_is_present())

            alert = driver.switch_to.alert
            alert.accept()

            print('Alert accepted')

            break
        except Exception as e:
            print(e)
            try:
                wait_until_clickable(5, "//button[@onclick='certNum1();']")
            except Exception as e:
                print(e)
                driver.get(url=URL)
                log_in()
                return

    while len(driver.window_handles) >= 2:
        time.sleep(0.5)

    driver.switch_to.window(driver.window_handles[-1])





    # WebDriverWait(driver, 10000).until(EC.alert_is_present())
    #
    # alert = driver.switch_to.alert
    # alert.accept()




    while len(driver.window_handles) >= 2:
        time.sleep(0.5)

    driver.switch_to.window(driver.window_handles[-1])

    wait_until_clickable(20, "//button[@id='bookmarkBtn']")

    print('Logging in done!')

if __name__ == "__main__":
    with open("config.json", "r", encoding='utf-8') as st_json:
        json_data = json.load(st_json)

    URL = 'https://po.ssgadm.com/main.ssg'

    ssg_id = json_data['ssg_id']
    ssg_pwd = json_data['ssg_pwd']
    sms_api_url = json_data['sms_api_url']
    sms_api_key = json_data['sms_api_key']
    post_api_url = json_data['post_api_url']
    secrete_key = json_data['secrete_key']
    auth_key = json_data['auth_key']
    encrypt_key = json_data['encrypt_key']

    profile = {'savefile.default_directory': os.getcwd(), 'download.default_directory': os.getcwd()}
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', profile)

    driver = webdriver.Chrome(executable_path='chromedriver', options=chrome_options)


    log_in()
    move_to_mobile_gift_order()
    select_condition()

    while True:
        try:
            if is_there_order():
                download_excel()
                item = parse_order_data()

                parse_data = dict()

                parse_data['secretkey'] = secrete_key
                parse_data['item'] = item

                digest = hmac.new(encrypt_key.encode('utf-8'), str(parse_data).encode('utf-8'), hashlib.sha256).digest()
                digest_b64 = base64.b64encode(digest)  # bytes again
                Hmac = auth_key + ':' + digest_b64.decode('utf-8')

                header = {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'Authorization': auth_key,
                    'Hmac': Hmac
                }

                response = requests.post(url=post_api_url, headers=header, data=json.dumps(parse_data))
                reply = response.json()

                print(response)
                print(response.json())

                if reply['msg'] == '성공':
                    mark_as_delivery_completed()
                    select_condition()
        except Exception as e:
            print('error! resetting...')
            while len(driver.window_handles) >= 2:
                driver.switch_to.window(driver.window_handles[-1])
                driver.close()
                time.sleep(0.5)

            driver.switch_to.window(driver.window_handles[-1])

            log_in()
            move_to_mobile_gift_order()
            select_condition()
