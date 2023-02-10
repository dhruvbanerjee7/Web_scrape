from flask import Flask, render_template, request,make_response,send_file
import pandas as pd
# from sklearn.ensemble import RandomForestRegressor
from selenium import webdriver
import time
import gspread
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
from cmegroup import gen_current_time,driver_define,get_web_data,add_sum,conn_sheet,main

# Configure application
app = Flask(__name__)

@app.route('/index', methods=['POST', 'GET'])
# @app.route('/', methods=['POST', 'GET'])
def index():

    greeting = 'Welcome to My Data Science Portfolio Website'

    return render_template('/index.html')


@app.route('/scrape', methods=['POST', 'GET'])
def home():
    msg = "Process started!! Your File will be downloaded in a few minutes"

    return render_template('boost.html', prediction_text='{}'.format(msg))

@app.route('/scrape_result',methods=['POST', "GET"])
def predict():
    '''
    For rendering results on HTML GUI
    '''
    main()
    x = "Done"
    return x



if __name__ == '__main__':
    app.run(debug=True)