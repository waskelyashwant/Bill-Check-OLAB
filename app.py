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

        file1 = open("value.txt","w")
        file1.write("-1")
        file1.close() 

        count=open("count.txt", "w")
        data = count.write("0")
        count.close()

        f = request.files['file'] 
        # driver = webdriver.Chrome("chromedriver.exe") 
        f.save(f.filename)

        # print(f)
        df = openpyxl.load_workbook(f)
        sheet = df.active

        df.save("bill_file.xlsx")
        
        data = pd.read_excel(r'bill_file.xlsx')
        df1 = pd.DataFrame(data)
        length = len(df1.index)

        totalfile=open("total.txt", "w")
        totalfile.write(str(length))
        totalfile.close()
   
        subprocess.Popen(["python", "side.py"])

        data=pd.read_excel("mapping.xlsx")
        # return render_template("submit.html", data = data.to_html() )
        return render_template("submit.html")


@app.route('/refresh', methods=['POST', 'GET'])
def refresh():
    f = open("value.txt", "r")
    val = f.read()
    print(val)
    f.close()
    if val == "1":
        print("inside value")
        path="status.xlsx"
        return send_file(path, as_attachment=True)
        # return redirect("http://127.0.0.1:5000/download")
    else :
    	count=open("count.txt", "r")
    	data = count.read()
    	count.close()
    	return render_template("submit.html", data = data)


@app.route('/download', methods=['POST', 'GET'])
def download():
    path="status.xlsx"
    return send_file(path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(debug=True)
