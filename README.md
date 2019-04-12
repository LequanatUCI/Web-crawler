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

def getWeatherData(ballpark, timezone, stationcode, input_filename, output_filename, chromedriver_path, stationdist):
    def timeToValue(hh, mm, sec, AP):
        if AP == 'AM':
            if hh == '12':
                hh = '00'
            timevalue = float(hh) * 60 + float(mm) + sec / 60
        elif AP == 'PM':
            if hh == '12':
                hh = '00'
            hh = float(hh) + 12
            timevalue = float(hh) * 60 + float(mm) + sec / 60
        return timevalue

    def timeConvert(hour, timezone):
        hour = int(hour)
        if timezone == 'PDT':
            hour = hour - 7
            if hour < 0:
                hour = hour + 24

        if timezone == 'MST':
            hour = hour - 6
            if hour < 0:
                hour = hour + 24

        if timezone == 'CDT':
            hour = hour - 5
            if hour < 0:
                hour = hour + 24

        if timezone == 'EDT':
            hour = hour - 4
            if hour < 0:
                hour = hour + 24

        if (hour >= 0) & (hour <= 11):
            AP = 'AM'
        elif (hour >= 12) & (hour <= 23):
            hour = hour - 12
            AP = 'PM'

        hour = str(hour)
        res = [hour, AP]
        return res

    def getMonth(mon):
        mon = int(mon)
        if mon == 1:
            month = 'January'
        elif mon == 2:
            month = 'February'
        elif mon == 3:
            month = 'March'
        elif mon == 4:
            month = 'April'
        elif mon == 5:
            month = 'May'
        elif mon == 6:
            month = 'June'
        elif mon == 7:
            month = 'July'
        elif mon == 8:
            month = 'August'
        elif mon == 9:
            month = 'September'
        elif mon == 10:
            month = 'October'
        elif mon == 11:
            month = 'November'
        elif mon == 12:
            month = 'December'
        return month

    def getTimeValue(timestring, ampm):
        if ampm == 'PM':
            h = float(timestring.split(':')[0])
            m = float(timestring.split(':')[1])
            if h == 12:
                return 12 * 60 + m
            return (h + 12) * 60 + m
        elif ampm == 'AM':
            h = float(timestring.split(':')[0])
            m = float(timestring.split(':')[1])
            if h == 12:
                return m
            return h * 60 + m

    def grabData(month, day, year, zip):
        try:
            # time1s = time.time()

            # INITIALIZE THE CHROME DRIVER
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            url2 = 'https://www.wunderground.com/history/'
            driver = webdriver.Chrome(chromedriver_path,
                                      chrome_options=chrome_options)  # , chrome_options=chrome_options

            driver.get(url2)

            elem = driver.find_element_by_id("histSearch")

            day = str(int(day))
            month = getMonth(month)
            s1 = Select(driver.find_element_by_name('month'))
            s1.select_by_visible_text(month)
            s2 = Select(driver.find_element_by_name('day'))
            s2.select_by_visible_text(day)
            s3 = Select(driver.find_element_by_name('year'))
            s3.select_by_visible_text(year)

            elem.send_keys(zip)

            temp = driver.find_element_by_xpath('//input[@class="button radius"]')

            temp.click()
            sleep(2)

            # time1e = time.time()
            # print('time1', time1e - time1s )

            url = driver.current_url
            # print(url)
            # driver2.close()

            # chrome_options = webdriver.ChromeOptions()
            # # 使用headless无界面浏览器模式
            # chrome_options.add_argument('--headless')
            # # chrome_options.add_argument('--disable-gpu')
            # driver = webdriver.Chrome(r'E:/Anaconda3/chromedriver.exe',
            #                           chrome_options=chrome_options)  # , chrome_options=chrome_options
            # time2s = time.time()

            driver.get(url)
            #driver.implicitly_wait(10)
            sleep(5)

            allcontents = driver.find_elements_by_class_name('tablesaw-sortable-btn')

            list_of_content = []
            for e in allcontents:
                list_of_content.append(e.text)

            ind_Temp = 0
            ind_Hum = 0
            ind_Pres = 0

            for i in range(len(list_of_content)):
                if list_of_content[i] == 'Temperature':
                    ind_Temp = i
                if list_of_content[i] == 'Humidity':
                    ind_Hum = i
                if list_of_content[i] == 'Pressure':
                    ind_Pres = i
                if ind_Hum * ind_Temp * ind_Pres != 0:
                    break

            table = driver.find_element_by_xpath("//table[@class='tablesaw-sortable'][@id='history-observation-table']")
            all_tr = table.find_elements_by_tag_name('tr')
            # num_of_rows = len(all_tr) - 1

            datalist_of_Time = []
            datalist_of_Temperature = []
            datalist_of_Humidity = []
            datalist_of_Pressure = []
            c1 = 0
            for tr in all_tr:
                if c1 == 0:
                    c1 = c1 + 1
                    continue
                data_in_row = tr.find_elements_by_tag_name('td')

                timestring = data_in_row[0].text
                timevalue = getTimeValue(timestring.split()[0], timestring.split()[1])

                datalist_of_Time.append(timevalue)

                datalist_of_Temperature.append(float(data_in_row[ind_Temp].text.split()[0]))
                datalist_of_Humidity.append(float(data_in_row[ind_Hum].text.split()[0]))
                datalist_of_Pressure.append(float(data_in_row[ind_Pres].text.split()[0]))

            driver.close()
            # time2e = time.time()
            # print('time2',time2e - time2s)
            return [datalist_of_Time, datalist_of_Temperature, datalist_of_Humidity, datalist_of_Pressure]
        except:
            print('!EXCEPT HAPPENS!' + '  ' +year + '-' + month + '-'+ day + '!!!')
            exceptionlist.append(p)
            # sleep(2)
            # tmonth = month
            # tday = day
            # tyear = year
            # tzip = zip
            # return grabdata(tmonth, tday, tyear, tzip)

    def findInd(t_hour, t_minute, sec, ampm, timelist):
        sec = float(sec)
        list_diff = []
        tv = timeToValue(t_hour, t_minute, sec, ampm)

        for j in range(len(timelist)):
            diff = abs(timelist[j] - tv)
            list_diff.append(diff)

        minabs = min(list_diff)
        ind = list_diff.index(minabs)
        diff = tv - timelist[ind]
        diff = round(diff, 2)

        return [ind, diff]

    workbook = load_workbook(filename=input_filename)
    sheetnames = workbook.get_sheet_names()
    first_sheet = sheetnames[0]
    booksheet = workbook.get_sheet_by_name(first_sheet)
    rows = booksheet.rows

    i = 0
    indd = 0
    datelist = []  # ALL DATES NEEDED TO SEARCH
    for row in rows:
        line = [col.value for col in row]
        if i == 0:
            i = i + 1
            continue
        if i == 1:
            datelist.append(line[0])

        if i > 1:
            if line[0] != datelist[indd]:
                datelist.append(line[0])
                indd = indd + 1
        i = i + 1

    length_of_datelist = len(datelist)
    print('number of total dates', length_of_datelist)
    exceptionlist = []

    print('/////////////////////////////////////////////////////')
    print('/////////////////START WEB SCRAPING//////////////////')
    print('/////////////////////////////////////////////////////')

    for p in range(length_of_datelist):

        if p < 0:
            continue
        print('#####', p)

        mon = datelist[p][5:7]
        day = datelist[p][8:10]
        year = datelist[p][0:4]

        file = ballpark + '_' + stationcode + '_' + str(p) + '_' + year + '-' + mon + '-' + day + '_weather_data.txt'
        wdatafile = open(file, 'wb')

        time2s = time.time()
        fourlistinone = grabData(mon, day, year, stationcode)  # FOUR LISTS: TIME, TEMPERATURE,

        pickle.dump(fourlistinone, wdatafile)
        wdatafile.close()

        time2e = time.time()
        print('## time consumed', time2e - time2s)


    print('#####EXCEPTION LIST', exceptionlist)
    print('#####',exceptionlist)

    print('/////////////////////////////////////////////////////')
    print('/////////////////PROCESS EXCEPTIONS//////////////////')
    print('/////////////////////////////////////////////////////')

    timeofloop = 0

    while len(exceptionlist) > 0:

        templist = []
        for element in exceptionlist:
            templist.append(element)

        timeofloop = timeofloop + 1
        for p in templist:

            print('###This is' + str(timeofloop) + 'th time of while loop###', p)

            mon = datelist[p][5:7]
            day = datelist[p][8:10]
            year = datelist[p][0:4]

            file = ballpark + '_' + stationcode + '_' + str(p) + '_' + year + '-' + mon + '-' + day + '_weather_data.txt'
            wdatafile = open(file, 'wb')

            time2s = time.time()

            fourlistinone = grabData(mon, day, year, stationcode)  # FOUR LISTS: TIME, TEMPERATURE,

            if fourlistinone:
                print('## this p is removed from exceptionlist', p)
                exceptionlist.remove(p)

            pickle.dump(fourlistinone, wdatafile)
            wdatafile.close()

            time2e = time.time()
            print('# time consumed', time2e - time2s)

    print('//////Number of loop is////////', timeofloop)

    #============================================================
    #============================================================

    ##########EXTRACT DATA FROM FILE STORING DATA
    print('/////////////////////////////////////////////////////')
    print('////////EXTRACT DATA FROM FILE STORING DATA//////////')
    print('/////////////////////////////////////////////////////')
    biglist = []
    for c in range(length_of_datelist):
        mon = datelist[c][5:7]
        day = datelist[c][8:10]
        year = datelist[c][0:4]
        fr1 = open(ballpark + '_' + stationcode + '_' + str(c) + '_' + year + '-' + mon + '-' + day + '_weather_data.txt', 'rb')
        templist = pickle.load(fr1)
        biglist.append(templist)
        fr1.close()

    workbook = load_workbook(filename=input_filename)
    sheetnames = workbook.get_sheet_names()
    first_sheet = sheetnames[0]
    booksheet = workbook.get_sheet_by_name(first_sheet)
    rows = booksheet.rows
    columns = booksheet.columns

    wb = Workbook()
    ws = wb.active
    c = 0

    ########################### STORE NEEDED DATA INOT EXCEL FILE
    for row in rows:
        line = [col.value for col in row]
        ind = 0
        print(c)
        if c == 0:
            c = c + 1
            continue

        zuluhour = line[1][7:9]
        minu = line[1][9:11]
        seco = line[1][11:13]

        temp = timeConvert(zuluhour, timezone)
        hour = temp[0]
        ampm = temp[1]

        date = line[0]
        date_ind = datelist.index(date)

        outind = findInd(hour, minu, seco, ampm, biglist[date_ind][0])  ####################### SEARCH THE CLOSEST TIME

        indd1 = outind[0]
        t1diff = outind[1]
        wtemp1 = biglist[date_ind][1][indd1]
        wRH1 = biglist[date_ind][2][indd1]
        wpres1 = biglist[date_ind][3][indd1]

        k = c + 1
        ws.cell(row=k, column=1).value = stationcode
        ws.cell(row=k, column=2).value = float(stationdist)
        ws.cell(row=k, column=3).value = float(t1diff)
        ws.cell(row=k, column=4).value = float(wtemp1)
        ws.cell(row=k, column=5).value = float(t1diff)
        ws.cell(row=k, column=6).value = float(wRH1)
        ws.cell(row=k, column=7).value = float(t1diff)
        ws.cell(row=k, column=8).value = float(wpres1)

        c = c + 1

    ws.cell(row=1, column=1).value = 'Stationcode'
    ws.cell(row=1, column=2).value = 'Stationdist'
    ws.cell(row=1, column=3).value = 't1diff'
    ws.cell(row=1, column=4).value = 'Temperature'
    ws.cell(row=1, column=5).value = 't1diff'
    ws.cell(row=1, column=6).value = 'Humidity'
    ws.cell(row=1, column=7).value = 't1diff'
    ws.cell(row=1, column=8).value = 'Pressure'

    wb.save(filename=output_filename)


#EXAMPLE
getWeatherData('Yankee Stadium',                                              #Name of ballpark
               'EDT',                                                           #Time zone
               'KLGA',                                                          #Stationcode
               'D:/Wunderground Excel/Yankee Stadium.xlsx',                   #Input excel file
               'D:/Wunderground Excel/Weather Data/WD_Yankee Stadium.xlsx',   #Output excel file
               'E:/Anaconda3/chromedriver.exe'                                  #the version of chromedriver should be correct
               ,'10000')                                                        #Distace between ballpaark and wunderground station
