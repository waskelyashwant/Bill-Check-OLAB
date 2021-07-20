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
    
@app.route('/login', methods=['POST', 'GET'])
def login():
    error = None
    if request.method == 'POST':

        f = request.files['file'] 
        # driver = webdriver.Chrome("chromedriver.exe") 
        f.save(f.filename)

        # print(f)
        df = openpyxl.load_workbook(f)
        sheet = df.active

        df.save("bill_file.xlsx")
        
        subprocess.Popen(["python", "side.py"]

        data=pd.read_excel("mapping.xlsx")
        return render_template("submit.html", data = data.to_html() )


@app.route('/download', methods=['POST', 'GET'])
def download():
    path="status.xlsx"
    return send_file(path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(debug=True)
