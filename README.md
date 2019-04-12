# Web-crawler
grab weather data

import xlwt
import xlrd
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import Select
import urllib.request
import bs4 as bs
from openpyxl import load_workbook
from openpyxl import Workbook
from selenium.webdriver.chrome.options import Options
import time
import pickle


#EXAMPLE
getWeatherData('Yankee Stadium',                                              #Name of ballpark
               'EDT',                                                           #Time zone
               'KLGA',                                                          #Stationcode
               'D:/Wunderground Excel/Yankee Stadium.xlsx',                   #Input excel file
               'D:/Wunderground Excel/Weather Data/WD_Yankee Stadium.xlsx',   #Output excel file
               'E:/Anaconda3/chromedriver.exe'                                  #the version of chromedriver should be correct
               ,'10000')                                                        #Distace between ballpaark and wunderground station
