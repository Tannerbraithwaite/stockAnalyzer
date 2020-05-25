import urllib.request
import zipfile
import os
import sqlite3
import time
import os, csv, datetime
from datetime import datetime
import xlsxwriter

conn = sqlite3.connect('stocksearch.db')
c = conn.cursor()

# c.execute('CREATE TABLE prices (SYMBOL text, SERIES text, OPEN real, HIGH real, LOW real, CLOSE real, LAST real, PREVCLOSE real, TOTTRDQTY real, TOTTRDVAL real, TIMESTAMP date, TOTALTRADES real, ISIN text, PRIMARY KEY (SYMBOL, SERIES, TIMESTAMP))')
# conn.commit()


# def download(localZipFilePath,urlOfFileName):
#     hdr = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1 Safari/605.1.15',
#            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
#            'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
#            'Accept-Encoding': 'none',
#            'Accept-Language': 'en-US,en;q=0.8',
#            'Connection': 'keep-alive'}
#     webRequest = urllib.request.Request(urlOfFileName, headers=hdr)
#     try:
#         page = urllib.request.urlopen(webRequest)
#         content = page.read()
#         output = open(localZipFilePath, "wb")
#         output.write(bytearray(content))
#         output.close()
#     except(urllib.request.HTTPError, e):
#         print(e.fp.read())
#         print("Looks like the download did not go through. Please download manually \nFROM:" + urlOfFileName + "\nTO:" + localZipFilePath)

def unzip(localZipFilePath, localExtractFilePath):
    if os.path.exists(localZipFilePath):
        print("Cool! " + localZipFilePath + " exists..proceeding")
        listOfFiles = []
        fh = open(localZipFilePath, 'rb')
        zipFileHandler = zipfile.ZipFile(fh)
        for name in zipFileHandler.namelist():
            zipFileHandler.extract(name, localExtractFilePath)
            listOfFiles.append(localExtractFilePath + name)
            print("Extracted " + name + " from the zip file, and saved to " + (localExtractFilePath + name))
        print("Extracted " + str(len(listOfFiles)) + " file in total")
        fh.close()

def unzipForPeriod(listOfMonths, listOfyears):
    for year in listOfYears:
        for month in listOfMonths:
            for dayOfMonth in range(31) :
                date = dayOfMonth + 1
                dateStr = str(date)
                if date < 10:
                    dateStr = "0"+dateStr
                print(dateStr, "-", month,"-", year)
                fileName = "cm" +str(dateStr) + str(month) + str(year) + "bhav.csv.zip"
                localZipFilePath = "/Users/tannerbraithwaite/github/stockAnalyzer/stockData/NSE_2006-16/" + fileName
                unzip(localZipFilePath,localExtractFilePath)
                time.sleep(5)
    print("OK, all done extracting")

# localExtractFilePath = "/Users/tannerbraithwaite/github/stockAnalyzer/stockData/extractedData/"
#
#
#
# listOfMonths = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
# listOfYears = ['2015']
#
# unzipForPeriod(listOfMonths,listOfYears)

def insertRows(fileName,conn):
    c = conn.cursor()
    lineNum = 0
    with open(fileName, 'r') as csvfile:
        lineReader = csv.reader(csvfile, delimiter = ',', quotechar = "\"")
        for row in lineReader:
            lineNum = lineNum + 1
            if lineNum ==1:
                print("Header row, skipping")
                continue
            date_object = datetime.strptime(row[10], '%d-%b-%Y')
            oneTuple = [row[0], row[1], float(row[2]),float(row[3]),float(row[4]),float(row[5]),float(row[6]),float(row[7]),float(row[8]),float(row[9]),date_object,float(row[11]),row[12]]
            c.execute("INSERT INTO prices VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",oneTuple)
        conn.commit()
        print("Done iterating over file contents - the file has been closed now!")

localExtractFilePath = "/Users/tannerbraithwaite/github/stockAnalyzer/stockData/extractedData/"
#
# for file in os.listdir(localExtractFilePath):
#     if file.endswith('.csv'):
#         insertRows(localExtractFilePath+file,conn)

t1 = 'ICICIBANK'
series = 'EQ'
c = conn.cursor()
cursor = c.execute ('SELECT symbol, max(close), min(close), max(timestamp), min(timestamp), count(timestamp) FROM prices WHERE symbol = ? and series = ? GROUP BY symbol ORDER BY timestamp', (t1,series))
for row in cursor:
    print(row)


def createExcelWithDailyPriceMoves(ticker,conn):
    c = conn.cursor()
    cursor = c.execute('SELECT symbol, timestamp, close FROM prices where symbol = ? and series = ? ORDER BY timestamp',(ticker,series))
    excelFileName = "/Users/tannerbraithwaite/github/stockAnalyzer/stockData/processedCSVs/"+ticker+".xlsx"
    workbook = xlsxwriter.Workbook(excelFileName)
    worksheet = workbook.add_worksheet("Summary")
    worksheet.write_row("A1",["Top Traded Stocks"])
    worksheet.write_row("A2",['Stock','Date','Closing'])
    lineNum = 3
    for row in cursor:
        worksheet.write_row("A"+str(lineNum), list(row))
        print("A"+str(lineNum),list(row))
        lineNum = lineNum + 1
    chart1 = workbook.add_chart({'type':'line'})
    chart1.add_series({
            'categories': '=Summary!$B$3:$B$' + str(lineNum),
            'values': '=Summary!$C$3:$C$'+str(lineNum)
        })
    chart1.set_title({'name': ticker})
    chart1.set_x_axis({'name':'Date'})
    chart1.set_y_axis({'name': 'Closing Price'})
    worksheet.insert_chart('F2',chart1,{'x_offset': 25, 'y_offset': 10})
    workbook.close()

conn=sqlite3.connect('stocksearch.db')
createExcelWithDailyPriceMoves('ICICIBANK',conn)


# conn=sqlite3.connect('stocksearch.db')
# c = conn.cursor()
# c.execute('DROP TABLE prices')
# conn.commit()
# conn.close()
