# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 13:42:42 2022

@author: abockelman
"""


#INIT
# (A1) LOAD MODULES
from flask import Flask, render_template, request, make_response
from openpyxl import load_workbook
import datetime
import ercot_daily_rucs

days_ago = 0

output = 'Daily EROCT RUC' + str(datetime.date.today() - datetime.timedelta(days=days_ago)) +'.xlsx'

print(output)
 
# (A2) FLASK SETTINGS + INIT
HOST_NAME = "localhost"
HOST_PORT = 80
app = Flask(__name__)
# app.debug = True
 
# (B) DEMO - READ EXCEL & GENERATE HTML TABLE
@app.route("/")
def index():
  # (B1) OPEN EXCEL FILE + WORKSHEET
  book = load_workbook(output)
  sheet =book.get_sheet_by_name( "Daily RUC")
 
  # (B2) PASS INTO HTML TEMPLATE
  return render_template("s3_excel_table.html", sheet=sheet)

def dynamic_page():
    return ercot_daily_rucs.your_function_in_the_module()

# (C) START
if __name__ == "__main__":
  app.run(HOST_NAME, HOST_PORT)