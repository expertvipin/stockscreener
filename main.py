import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import datetime
from selenium.webdriver.firefox.options import Options
import pandas as pd


options = Options()
options.set_preference("dom.webnotifications.enabled", False)
options.set_preference("app.update.enabled", False)
path = 'geckodriver'

class CollectData(object):
    def __init__(self, ids=['State Bank Of Inida']):
        self.ids = ids
        self.driver = webdriver.Firefox(executable_path=path,options=options) #executable_path=path, options=options

    def collect(self, typ):
        if typ == "funds":
            self.CollectFunds(self.ids)
        elif typ == "stocks":
            self.CollectStocks(self.ids)
        self.driver.close()

    def CollectStocks(self, stock_names):
        url = "https://www.moneycontrol.com/"

        self.driver.get(url)
        self.driver.maximize_window()
        try:
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/div/div[1]/span/a')))
            self.driver.find_element(
                By.XPATH, "/html/body/div/div[1]/span/a").click()
        except:
            pass
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.ID, 'search_str')))

        dataframes = []
        flag = 1
   
        for stock in stock_names:
            try:
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="form_topsearch"]//descendant::input[5]'))).send_keys(stock)            # input_bar = self.driver.find_element(

            except:
                pass
            if flag == 1:
                WebDriverWait(self.driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/div[1]/header/div[1]/div[1]/div/div/div[2]/div/a')))
                WebDriverWait(self.driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/div[1]/header/div[1]/div[1]/div/div/div[2]/div/a'))).click()

            else:
                WebDriverWait(self.driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/header/div[1]/div[1]/div/div/div[2]/div/a')))

                WebDriverWait(self.driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/header/div[1]/div[1]/div/div/div[2]/div/a'))).click()

            flag = 0
            try:
                WebDriverWait(self.driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="mc_mainWrapper"]/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/p/a')))

                WebDriverWait(self.driver, 40).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="mc_mainWrapper"]/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/p/a'))).click()

            except:
                pass
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'price_perfo')))
            comp_id = self.driver.current_url.split('/')[-1]
            table_data = self.driver.find_element(By.CLASS_NAME, 'price_perfo')

            rows = len(table_data.find_elements(
                By.XPATH, '/html/body/div[10]/div[2]/div[2]/div[6]/div[9]/div[2]/div[1]/table/tbody/tr'))

            cols = len(table_data.find_elements(
                By.XPATH, '/html/body/div[10]/div[2]/div[2]/div[6]/div[9]/div[2]/div[1]/table/tbody/tr[3]/td'))
            
            data, value = [], []
            for r in range(1, rows+1):
                data.append(table_data.find_element(
                    By.XPATH, '/html/body/div[10]/div[2]/div[2]/div[6]/div[9]/div[2]/div[1]/table/tbody/tr['+str(r)+']/td[1]').text)
                if r in (1, 2):
                    value.append(table_data.find_element(
                        By.XPATH, '/html/body/div[10]/div[2]/div[2]/div[6]/div[9]/div[2]/div[1]/table/tbody/tr['+str(r)+']/td[3]').text)
                else:
                    value.append(table_data.find_element(
                        By.XPATH, '/html/body/div[10]/div[2]/div[2]/div[6]/div[9]/div[2]/div[1]/table/tbody/tr['+str(r)+']/td[2]').text)

            for scroll in range(1):
                self.driver.execute_script(
                    "window.scrollTo(0, document.body.scrollHeight);")
                current_value = self.driver.find_element(
                    By.XPATH, "/html/body/div[10]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div[2]").text
                if current_value == '':
                    continue

                new_url = "https://www.moneycontrol.com/stocks/histstock.php?sc_id=" + \
                    comp_id+"&mycomp="+stock
                self.driver.get(new_url)

                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="mycomp"]'))).clear()
                
                mycomp = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="mycomp"]')))
                mycomp.send_keys(stock)
                
                dropdown = WebDriverWait(self.driver,30).until(EC.presence_of_element_located((
                    By.XPATH, '/html/body/div[6]/div[2]/div[1]/div[5]/div[2]/div[6]/table/tbody/tr/td[5]/div[2]/form/select')))
                select = Select(dropdown)
                today = datetime.date.today()
                year = str(today.year-5)
                select.select_by_value(year)
                self.driver.find_element(
                    By.XPATH, '//*[@id="mc_mainWrapper"]/div[2]/div[1]/div[5]/div[2]/div[6]/table/tbody/tr/td[5]/div[2]/form/input[1]').click()
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'MT12')))
                tableData = self.driver.find_element(By.CLASS_NAME, 'MT12')
                rows = len(tableData.find_elements(
                    By.XPATH, '/html/body/div[6]/div[2]/div[1]/div[4]/div[4]/table/tbody/tr'))
                cols = len(tableData.find_elements(
                    By.XPATH, '/html/body/div[6]/div[2]/div[1]/div[4]/div[4]/table/tbody/tr[2]/td'))

                for row in range(2, rows+1):
                    yearData = tableData.find_element(
                        By.XPATH, '/html/body/div[6]/div[2]/div[1]/div[4]/div[4]/table/tbody/tr['+str(row)+']/td[1]').text
                    if yearData == year:
                        data.append("5 year")
                        closePrice = tableData.find_element(
                            By.XPATH, '/html/body/div[6]/div[2]/div[1]/div[4]/div[4]/table/tbody/tr['+str(row)+']/td['+str(cols)+']').text

                        percentage = (
                            (float(current_value) - float(closePrice))/float(closePrice))*100
                        value.append("{:.2f}".format(percentage)+"%")

            Data = []
            for i in range(len(data)):
                aux = []
                aux.append(data[i])
                if (len(value[i]) != 0):
                    aux.append(value[i])
                else:
                    aux.append("NA")
                Data.append(aux)

            df = pd.DataFrame(Data, columns=['Timeline', 'Return(%)'])
            df.index += 1
            dataframes.append(df)
            print(dataframes)
        try:
            generate_excel_sheet(
                stock_names=stock_names, dataframes=dataframes)
            print("success")
        except Exception as e:
            print("somthing went wrong", e)

    def CollectFunds(self, mf_names):
        url = "https://www.moneycontrol.com/"

        self.driver.get(url)
        self.driver.maximize_window()
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/div/div[1]/span/a')))
            self.driver.find_element(
                By.XPATH, "/html/body/div/div[1]/span/a").click()
        except:
            pass
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, 'search_str')))
        MFData = []
        dataframes = []
        for mf in mf_names:
            input_bar = self.driver.find_element(By.ID, "search_str")
            input_bar.send_keys(mf)
            time.sleep(2)
            input_bar.click()
            WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="tab2"]')))
            self.driver.find_element(By.XPATH, '//*[@id="tab2"]').click()
            self.driver.find_element(
                By.XPATH, '//*[@id="autosuggestlist"]/ul/li[1]/a').click()
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="mc_content"]/div/section[9]/div/ul/li[1]/div/div[2]/span[1]')))
            std_deviation = self.driver.find_element(
                By.XPATH, '//*[@id="mc_content"]/div/section[9]/div/ul/li[1]/div/div[2]/span[1]').text
            aux = []
            aux.append(mf)
            if (std_deviation == ""):
                aux.append("NA")
            else:
                aux.append(std_deviation)
            MFData.append(aux)
        now = datetime.datetime.now()
        fdate = str(now.strftime("%d/%m/%Y"))
        fdate = fdate.replace("/", "_")
        ftime = str(now.strftime("%H:%M:%S"))
        ftime = ftime.replace(":", "_")
        file_name = os.path.join("FundsData_"+fdate+"_"+ftime+".xlsx")
        pd.DataFrame(MFData, columns=['MutualFund', 'Std Deviation']).to_excel(
            file_name, sheet_name="MutualFund")
        print(file_name+" save successfully")            


def generate_excel_sheet(dataframes, stock_names):
    now = datetime.datetime.now()
    fdate = str(now.strftime("%d/%m/%Y"))
    fdate = fdate.replace("/", "_")

    ftime = str(now.strftime("%H:%M:%S"))
    ftime = ftime.replace(":", "_")

    file_name = os.path.join("StocksData_"+fdate+"_"+ftime+".xlsx")
    
    with open(file_name, 'w+'):
        with pd.ExcelWriter(file_name) as writer:

            for i in range(len(stock_names)):

                dataframes[i].to_excel(
                    writer, sheet_name="{0}".format(stock_names[i]))
    
