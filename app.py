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

        count=open("count0.txt", "w")
        count.write("0")
        count.close()

        count=open("count1.txt", "w")
        count.write("0")
        count.close()

        count=open("count2.txt", "w")
        count.write("0")
        count.close()
        
        count=open("count.txt", "w")
        count.write("")
        count.close()
        
        stopfile = open("stop.txt", "w")
        stopfile.write("0")
        stopfile.close()

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
        totalfile.write(str(length) + " completed")
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
    	count=open("count0.txt", "r")
    	data1 = count.read()
    	count.close()

    	count=open("count1.txt", "r")
    	data2 = count.read()
    	count.close()

    	count=open("count2.txt", "r")
    	data3 = count.read()
    	count.close()

    	data=int(data1) + int(data2) + int(data3)
        
    	count = open("count.txt", "r")
    	line = count.read()
    	count.close()

    	if line != "":
    		data = line
        
    	totalfile=open("total.txt", "r")
    	total = totalfile.read()
    	totalfile.close()
    	return render_template("submit.html", data = str(data) , total = total)

@app.route('/stop', methods=['POST', 'GET'])
def stop():
    stopfile = open("stop.txt", "r")
    value= stopfile.read()
    stopfile.close()

    if value=="1":
    	data="Script is already stopped. Wait for the file to generate and click on REFRESH button to get the new info"
    	total=" "
    	return render_template("submit.html", data = str(data), total = total)

    stopfile = open("stop.txt", "w")
    stopfile.write("1")
    stopfile.close()
    data="Script is stopped. Wait for the file to generate and click on REFRESH button to get the new info"
    total = " "
    return render_template("submit.html", data = str(data), total = total)


@app.route('/download', methods=['POST', 'GET'])
def download():
    path="status.xlsx"
    return send_file(path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(debug=True)
