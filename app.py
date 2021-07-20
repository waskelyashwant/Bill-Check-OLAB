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
import threading
import subprocess


chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = os.environ.get("GOOGLE_CHROME_BIN")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")

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
        # data = pd.read_excel(r'mapping.xlsx')
        # df1 = pd.DataFrame(data)
        # length = len(df1.index)
        # mapping_dict={}
        # for i in range(0,length):
        #     mapping_dict[df1.iloc[i]['Distributor as per file']]=str(df1.iloc[i]['Zone'])
        # print(mapping_dict)

        f = request.files['file'] 
        # driver = webdriver.Chrome("chromedriver.exe") 
        f.save(f.filename)

        # print(f)
        df = openpyxl.load_workbook(f)
        sheet = df.active

        df.save("bill_file.xlsx")
   
        # z=0
        # real_list=[]
        # for k in sheet:
        #     if z==0:
        #         z+=1
        #         continue
        #     k=list(k)
        #     if k[1].value==None:
        #         continue
        #     else:
        #     	# k=[str(k[0].value), k[1].value, k[2].value, k[3].value, k[4].value]
        #     	k=["1","2"]
        #     real_list.append(bytes(k))

        # print("real_list", real_list)

        # myList = []
        subprocess.Popen(["python", "side.py"])

        # x=len(real_list)
        # y=int(x/20)
        # main_list=[[0, y],[y, 2*y],[2*y, 3*y],[3*y, 4*y],[4*y, 5*y],[5*y, 6*y],[6*y, 7*y],[7*y, 8*y],[8*y, 9*y],[9*y, 10*y],
        #            [10*y, 11*y],[11*y, 12*y],[12*y, 13*y],[13*y, 14*y],[14*y, 15*y],[15*y, 16*y],[16*y, 17*y],[17*y, 18*y],[18*y, 19*y],[19*y, 20*y+ x %20]]
        # threads=[]
        # results=[]
        # for i in range(0,20):
        #     res=[None]*(y+x%20)
        #     results.append(res)

        # for v in range(19,20):
        #     t=threading.Thread(target=starting, args=(real_list,main_list[v],results[v], mapping_dict))
        #     t.start()
        #     threads.append(t)
        # for v in threads:
        #     v.join()

        # print(results)
        # main_result=[]
        # for v in range(19,20):
        #     for u in range(0,len(results[v])):
        #         if results[v][u]!=None:
        #             main_result.append(results[v][u])
        # index=1
        # sheet.cell(row = index, column = 6).value = 'Status'
        # df.save('status.xlsx')
        # index+=1
        # for k in range(0,len(main_result)):
        #     if sheet.cell(row=index, column= 1).value == None:
        #         break
        #     sheet.cell(row = index, column = 6).value = main_result[k]
        #     index+=1
        #     df.save('status.xlsx')
        # driver.close()
        # data = pd.read_excel('status.xlsx')

        data=pd.read_excel("mapping.xlsx")
        return render_template("submit.html", data = data.to_html() )


@app.route('/download', methods=['POST', 'GET'])
def download():
    path="status.xlsx"
    return send_file(path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(debug=True)
