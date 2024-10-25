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



if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app)