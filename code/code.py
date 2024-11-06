from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
import sys

#定义通过代码寻找中文货币名称的函数
def find_currency_name_by_code(currency_code):
    currency_df = pd.read_excel('货币名称符号对照表.xlsx', engine='openpyxl')

    currency_name_row = currency_df[currency_df.iloc[:, 1] == currency_code]

    if not currency_name_row.empty:
        currency_name = currency_name_row.iloc[0, 0]
        return currency_name
    else:
        raise ValueError("No code found")

def fetch_exchange_rate(date, currency_code):
    # 配置WebDriver
    driver = webdriver.Chrome()

    try:
        # 打开目标网页
        driver.get('https://www.boc.cn/sourcedb/whpj/')

        # 等待页面加载完毕
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'historysearchform')))

        # 定位并输入日期
        start_date_input_xpath = '//*[@id="historysearchform"]/div/table/tbody/tr/td[2]/div/input'
        start_date_input = driver.find_element(By.XPATH, start_date_input_xpath)
        start_date_input.clear()
        start_date_input.send_keys(date)

        end_date_input_xpath = '//*[@id="historysearchform"]/div/table/tbody/tr/td[4]/div/input'
        end_date_input = driver.find_element(By.XPATH, end_date_input_xpath)
        end_date_input.clear()
        end_date_input.send_keys(date)

        # 定位并选择货币名称
        dropdown_xpath = '//*[@id="pjname"]'
        dropdown = driver.find_element(By.XPATH, dropdown_xpath)
        select = Select(dropdown)
        select.select_by_visible_text(currency_code)

        # 点击搜索按钮
        search_button_xpath = '//*[@id="historysearchform"]/div/table/tbody/tr/td[7]/input'
        search_button = driver.find_element(By.XPATH, search_button_xpath)
        search_button.click()

        # 等待结果加载
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'BOC_main')))

        # 获取所需信息
        target_xpath = '/html/body/div/div[4]/table/tbody/tr[2]/td[4]'
        target_element = driver.find_element(By.XPATH, target_xpath)
        return target_element.text

    finally:
        # 关闭浏览器
        driver.quit()

def write_to_file(filename, date, currency, data):
    # 将数据写入文件
    try:
        with open(filename, 'a') as file:
            file.write(date + ' ' + currency + ' ' + data + '\n')
    except IOError as e:
        print(f"Error writing to file: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python yourcode.py [date] [currency_code]")
        sys.exit(1)

    date = sys.argv[1]
    currency_code = sys.argv[2]
    
    # 获取现汇卖出价
    currency = find_currency_name_by_code(currency_code)
    sell_price = fetch_exchange_rate(date, currency)

    if sell_price:
        print(f"Exchange Rate: {sell_price}")
        # 将结果写入文件
        write_to_file('result.txt', date, currency_code, sell_price)
    else:
        print("Failed to fetch exchange rate.")