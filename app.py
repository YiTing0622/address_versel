#app.py

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from urllib.parse import quote
from pymongo.mongo_client import MongoClient
import requests
import xml.etree.ElementTree as ET
import json
import openpyxl
import datetime
import io
import time 
import os
import uvicorn
import webbrowser
import re
from openai import OpenAI
import pandas as pd

app = FastAPI()


@app.get('/')
def hello_world():
    return "Hello,World"

@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile):
    start = time.time()
    # check file 
    if file.filename.endswith('.xlsx'):
        f = await file.read()
        xlsx = io.BytesIO(f)
        wb = openpyxl.load_workbook(xlsx)
        ws = wb.active
        col_name_list = ['序號','地址','編號','站所四碼代號','站所名稱','站所簡碼','郵遞區號','ＳＤ註區','ＳＤ疊區','ＤＤ註區','ＤＤ疊區','ＭＤ註區','ＭＤ疊區','低溫註區','低溫疊區','聯運費用(起碼)','聯運費用(百公斤)','PUTAO_FLAG','新集貨註區','PUTCODE1','PUTCODE2','PUTCODE3','PUTCODE4','PUTCODE5','PUTCODE6','類別','到著簡碼','衛星區','註區','疊區','QRcode','聯運區提醒(*表示有聯運區,空白表示沒有)','MD 到著碼衛星區','MD 到著碼註區','MD 到著碼疊區','訊息']
        address_list = []
        for index, value in enumerate(col_name_list, 1):
            ws.cell(row=1, column=index, value=value)
        print(ws.max_row)
        for cells in ws['B2':'B' + str(ws.max_row)]:
            for cell in cells:
                if cell.value is not None:
                    cleaned_address = clean_address(cell.value)
                    address_list.append(cleaned_address)
                else:
                    address_list.append("")

        # call hsinchu api
        soap_response = send_soap_request_by_address_list(address_list)
        # write to excel
        response = parse_soap_response_by_list_uploadfile(wb, soap_response)
        print(response)
        if response != True:
            raise HTTPException(404, response)
        current_time = datetime.datetime.now()
        
        filename =  "地址比對" + str(current_time.year) + '-' + str(current_time.month) + '-' + str(current_time.day) + '-' + str(current_time.hour) + '-' + str(current_time.minute) + '-' + str(current_time.second) + ".xlsx"
        print(filename)
        wb.save('new.xlsx')
    else:
        raise EOFError
    # return None
    print(time.time()-start)
    return FileResponse('new.xlsx', headers={"Content-Disposition": f"attachment; filename={quote(filename)}"})


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app)