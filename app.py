
from flask import Flask , render_template , request
from flask import send_file
from selenium import webdriver
import selenium
from selenium.webdriver.remote import webelement
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import openpyxl
import os
import time
import pandas as pd


app = Flask(__name__, template_folder='templates')

@app.route('/')
def main():
    return render_template('app.html')

# @app.route('/send')
# def send():
#     return render_template('app.html')


@app.route('/login', methods=['POST', 'GET'])
def login():
    error = None
    if request.method == 'POST':
        f = request.files['file'] 
        # driver = webdriver.Chrome("chromedriver.exe") 
        f.save(f.filename)
        print(f)  
        df = openpyxl.load_workbook(f)
        sheet = df.active
        index = 1
        driver = webdriver.Chrome("chromedriver.exe")
        zone = request.form.get('username')
        if zone == 'Jaipur':
            link = "https://www.amazon.in/hfc/bill/electricity?ref_=apay_deskhome_Electricity"
            driver.get(link)
            driver.find_element_by_class_name("a-dropdown-prompt").click()
            x = driver.find_element_by_class_name("a-dropdown-common")
            ul = x.find_element_by_tag_name("ul")
            li = ul.find_elements_by_tag_name("li")
            li[26].click()
            time.sleep(2)
            x1 =Select(driver.find_element_by_id("ELECTRICITY>hfc-states-rajasthan"))
            x1.select_by_index(1)
            x1._setSelected
            webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(2)
            # driver.find_element_by_class_name("a-form-lable").click()

            
            for k_no in sheet['D']:
                if index == 1:
                    sheet.cell(row = index, column = 6).value = 'Status'
                    index += 1
                else:
                    if sheet.cell(row=index, column= 1).value == None:
                        break
                    ags = sheet.cell(row=index, column=5).value
                    pk = "Continue to Pay ₹"+ str(ags)+".00"
                    print(pk)
                    driver.find_element_by_id("K Number").click()
                    driver.find_element_by_id("K Number").clear()
                    driver.find_element_by_id("K Number").send_keys(k_no.value)
                    driver.find_element_by_id("fetchBtnText").click()
                    daa = driver.find_element_by_id("paymentBtnAmountText")
                    print(daa.get_attribute("innerHTML"))
                    if pk==daa.get_attribute("innerHTML"):
                        print("sdfdf")
                        sheet.cell(row = index, column = 6).value = "unpaid"
                        df.save('status.xlsx')
                    else:
                        print("asd")
                        sheet.cell(row = index, column = 6).value = "paid"
                        df.save('status.xlsx')
                    
                    time.sleep(2)



            # print(zone)
            
        elif zone == 'Jodhpur':
            link = "http://wss.rajdiscoms.com/HDFC_QUICKPAY/index"
            for k_no in sheet['E']:
                if index == 1:
                    sheet.cell(row = index, column = 7).value = 'Status'
                    index += 1
                else:
                    if sheet.cell(row=index, column=2).value == None:
                        break
                    driver.get(link)
                    driver.find_element_by_id("txtKno").click()
                    driver.find_element_by_id("txtKno").clear()
                    driver.find_element_by_id("txtKno").send_keys(k_no.value)
                    driver.find_element_by_id("txtEmail").click()
                    driver.find_element_by_id("txtEmail").clear()
                    driver.find_element_by_id("txtEmail").send_keys("admin@gmail.com")
                    driver.find_element_by_id("btnsearch").click()
                    status = driver.find_element_by_id("lblMessage").text
                    print(status)
                    sheet.cell(row = index, column = 7).value = status
                    df.save('status.xlsx')
                    index += 1
            # driver.close()
            print(zone)
            
        elif zone=="Ajmer":
            link = "https://jansoochna.rajasthan.gov.in/Services/DynamicControls"
            driver.get(link)
            driver.find_element_by_partial_link_text("Know about your Electricity Bill Payment Information - AVVNL").click()
            time.sleep(2)
            driver.find_element_by_id("Enter_your_K_number").click()
            for k_no in sheet['E']:
                if index == 1:
                    sheet.cell(row = index, column = 7).value = 'Status'
                    index += 1
                else:
                    if sheet.cell(row=index, column=2).value == None:
                        break
                    
                    driver.find_element_by_id("Enter_your_K_number").clear()
                    driver.find_element_by_id("Enter_your_K_number").send_keys(k_no.value)  
                    driver.find_element_by_id("btnSubmit").click()
                    time.sleep(2)   
                    amnt= driver.find_element_by_xpath("/html/body/div[1]/section/div[3]/div/div/div/div/div[3]/div[2]/div/div[1]/div[2]/table/tbody/tr[1]/td[9]")
                    
                    # print(asds.get_attribute("innerHTML"))
                    dat =  driver.find_element_by_xpath("/html/body/div[1]/section/div[3]/div/div/div/div/div[3]/div[2]/div/div[1]/div[2]/table/tbody/tr[1]/td[12]")
                    if amnt==sheet.cell(row=index, column=5).value:
                        sheet.cell(row = index, column = 7).value = "paid"
                        df.save('status.xlsx')
                    else:
                        sheet.cell(row = index, column = 7).value = "unpaid"
                        df.save('status.xlsx')
                    index+=1
                    
            # driver.close()

        else:
            link = "https://jansoochna.rajasthan.gov.in/Services/DynamicControls"
            driver.get(link)
            driver.find_element_by_partial_link_text("Know about your Electricity Bill Payment Information - JDVVNL").click()
            time.sleep(2)
            driver.find_element_by_id("Enter_your_K_number").click()
            for k_no in sheet['E']:
                if index == 1:
                    sheet.cell(row = index, column = 7).value = 'Status'
                    index += 1
                else:
                    if sheet.cell(row=index, column=2).value == None:
                        break
                    driver.find_element_by_id("Enter_your_K_number").clear()
                    driver.find_element_by_id("Enter_your_K_number").send_keys(k_no.value)  
                    driver.find_element_by_id("btnSubmit").click()
                    time.sleep(2)   
                    amnt= driver.find_element_by_xpath("/html/body/div[1]/section/div[3]/div/div/div/div/div[3]/div[2]/div/div[1]/div[2]/table/tbody/tr[1]/td[12]")
                    amnt = amnt.get_attribute("innerHTML")
                    print(amnt)
                    dat =  driver.find_element_by_xpath("/html/body/div[1]/section/div[3]/div/div/div/div/div[3]/div[2]/div/div[1]/div[2]/table/tbody/tr[1]/td[13]")
                    dat = dat.get_attribute("innerHTML")
                    if amnt==str(sheet.cell(row=index, column=5).value):
                        sheet.cell(row = index, column = 7).value = "paid"
                        df.save('status.xlsx')
                    else:
                        sheet.cell(row = index, column = 7).value = "unpaid"
                        df.save('status.xlsx')
        
        driver.close()
        data = pd.read_excel('status.xlsx')
   
        return render_template("submit.html", data = data.to_html() )

@app.route('/download', methods=['POST', 'GET'])
def download():
    path="status.xlsx"
    return send_file(path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(debug=True)
