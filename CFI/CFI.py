import requests
from bs4 import BeautifulSoup
import xlsxwriter
import datetime
from enum import Enum
import os
import time

totalTaskCount = 0
tasksDoneCount = 0
def taskDone():
    global tasksDoneCount
    tasksDoneCount = tasksDoneCount + 1
    print("progress: " + str(
        round(tasksDoneCount / totalTaskCount * 100)
        ).rjust(3) + "%", end = "\r")

def isCFIDown(res):
    resc = BeautifulSoup(res.content, "html.parser")
    return resc.html is None

def safeCFIget(URL):
    res = requests.get(URL)
    if isCFIDown(res):
        for i in range(60):
            print("progress: " + str(
                round(tasksDoneCount / totalTaskCount * 100)
                ).rjust(3) + "%% | CFI IS DOWN, reattempting in %02d seconds!" % (60 - i), end="\r")
            time.sleep(1)
        print("progress: " + str(
            round(tasksDoneCount / totalTaskCount * 100)
            ).rjust(3) + "%" + " " * 43, end="\r")
        return safeCFIget(URL)
    return res

thisYear = datetime.datetime.now().year
listGen = lambda yearsToGoBack : [str(thisYear - i) for i in range (yearsToGoBack + 1)]
yearsToQuery = listGen(5)
yearsToQuery_重大事件 = listGen(3)

def getCFIStockCode(A股上市公司代码):
    CFI_HOME_URL = "http://quote.cfi.cn/quote_" + A股上市公司代码 + ".html"
    soup = BeautifulSoup(safeCFIget(CFI_HOME_URL).content, "html.parser")
    taskDone()
    return(soup.find(id="nodea1").nobr.a["href"].split("/")[2])

#---------begin: vertical tables---------------------------

class CFI_VPage(Enum):
    财务分析指标 = "cwfxzb"
    资产负债表 = "zcfzb_x"
    利润分配表 = "lrfpb_x"
    现金流量表 = "xjll"
    非经常性损益合计表 = "FJCXSY_HJ"

def isVTableEnd(tr):
    return (tr is not None and
    tr.td is not None and
    tr.td.a is not None and
    tr.td.a.string == "返回页顶")

def vTable_item_name(tr):
    if tr.td.string is not None: return tr.td.string
    if tr.td.a is not None: return " " + tr.td.a.string
    return

def fetchNewestQuarter(cfiVPage, stockid, year):
    URL = ("http://quote.cfi.cn/quote.aspx?contenttype=" + cfiVPage.value
     + "&stockid=" + stockid
      + "&jzrq=" + year)
    soup = BeautifulSoup(safeCFIget(URL).content, "html.parser")

    vertical_table = soup.find(class_="vertical_table")

    if vertical_table is None or vertical_table.string == "我们在该表没找到该只股票。":
        taskDone()
        return (None, None)

    # skipping the table title; should point to the "截止日期" row
    row = vertical_table.tr.next_sibling

    (item_names, values) = ([], [])

    while not isVTableEnd(row):
        item_names.append(vTable_item_name(row))
        values.append(row.td.next_sibling.string)
        row = row.next_sibling

    taskDone()

    return (item_names, values)

def writeToExcelNewestQuarter(worksheet, A股上市公司s, A股上市公司代码s, cfiVPage, year):
    first_company_item_names = None
    for i in range(len(A股上市公司s)):
        worksheet.write(0, i + 1, A股上市公司s[i])
        (item_names, values) = fetchNewestQuarter(cfiVPage, A股上市公司代码s[i], year)

        if item_names is not None:
            if (first_company_item_names is None):
                first_company_item_names = item_names
                for j in range(len(item_names)):
                    worksheet.write(j + 1, 0, item_names[j])
            elif (first_company_item_names != item_names):
                worksheet.write(1, i + 1, "Err: Vector Mismatch")
                continue

            for j in range(len(values)):
                worksheet.write(j + 1, i + 1, values[j])

# ---------boundary between vertical tables (above) and horizontal tables (below)-----
def fetchHTable_产品分布(stockid, year):
    URL = ("http://quote.cfi.cn/quote.aspx?stockid="
    + stockid + "&contenttype=zyfb&jzrq="
    + year + "-12-31")
    soup = BeautifulSoup(safeCFIget(URL).content, "html.parser")
    tables_found = soup.find_all(id="tabh")
    return None if len(tables_found) != 3 else tables_found[2]

def fetchHTable_重大事件(stockid, year):
    URL = ("http://quote.cfi.cn/quote.aspx?stockid="
    + stockid + "&contenttype=zdsj&jzrq="
    + year)
    soup = BeautifulSoup(safeCFIget(URL).content, "html.parser")
    return soup.find(id="tabh")

def is重大事件LastLine(tr):
    return (tr is not None and
    tr.td is not None and
    tr.td.has_attr("colspan") and
    tr.td["colspan"] == "3" and
    tr.td.has_attr("style") and
    tr.td["style"] == "text-align:right;" and
    tr.td.string is None)

def readHTable(table):
    row = table.tr.next_sibling
    (item_names, values) = ([td.string for td in row.contents], [])
    row = row.next_sibling

    while (row is not None) and (not is重大事件LastLine(row)):
        values.append([td.string for td in row.contents])
        row = row.next_sibling

    taskDone()

    return (item_names, values)

def writeHTableValuesToExcel(worksheet, rowToWrite, values):
    for row in values:
        for i in range(len(row)):
            worksheet.write(rowToWrite, 1 + i, row[i])
        rowToWrite = rowToWrite + 1
    return rowToWrite

def writeToExcelHTables(workbook, is重大事件, A股上市公司s, A股上市公司代码s):
    first_company_first_year_item_names = None
    sheet = workbook.add_worksheet("重大事件" if is重大事件 else "产品分布")
    rowToWrite = 0
    for i in range(len(A股上市公司s)):
        if i != 0: sheet.write(rowToWrite, 0, A股上市公司s[i])
        for year in yearsToQuery_重大事件 if is重大事件 else yearsToQuery:
            fetcher = fetchHTable_重大事件 if is重大事件 else fetchHTable_产品分布
            hTable = fetcher(A股上市公司代码s[i], year)
            if hTable is not None:
                (item_names, values) = readHTable(hTable)
                if first_company_first_year_item_names is None:
                    first_company_first_year_item_names = item_names
                    sheet.write(rowToWrite, 0, "公司名称")
                    for j in range(len(first_company_first_year_item_names)):
                        sheet.write(rowToWrite, 1 + j, first_company_first_year_item_names[j])
                    rowToWrite = rowToWrite + 1
                    sheet.write(rowToWrite, 0, A股上市公司s[i])
                    rowToWrite = writeHTableValuesToExcel(sheet,
                    rowToWrite, values)
                elif item_names != first_company_first_year_item_names:
                    sheet.write(rowToWrite, 1, year)
                    sheet.write(rowToWrite, 2, "Err: Vector Mismatch")
                    rowToWrite = rowToWrite + 1
                else:
                    rowToWrite = writeHTableValuesToExcel(sheet,
                    rowToWrite, values)
            else:
                taskDone()

def downloadExcelFromCFI(上市公司列表Path, outputDir, numYearsToQuery, numYearsToQuery_重大事件):
    with open(上市公司列表Path, "r") as f:
        A股上市公司s = f.read().splitlines()

    global yearsToQuery, yearsToQuery_重大事件
    yearsToQuery = listGen(numYearsToQuery)
    yearsToQuery_重大事件 = listGen(numYearsToQuery_重大事件)

    global totalTaskCount
    totalTaskCount = (len(A股上市公司s) + len(yearsToQuery) * len(A股上市公司s) * len(CFI_VPage)
    + (len(yearsToQuery) + len(yearsToQuery_重大事件)) * len(A股上市公司s))

    A股上市公司代码s = [getCFIStockCode(line.split("\\")[1]) for line in A股上市公司s]
    for i in range(len(A股上市公司s)):
        A股上市公司s[i] = A股上市公司s[i].split("\\")[0]

    with xlsxwriter.Workbook(os.path.join(outputDir, "CFI.xlsx")) as workbook:
        for cfiVPage in CFI_VPage:
            for year in yearsToQuery:
                sheet = workbook.add_worksheet(cfiVPage.name + "-" + year)
                writeToExcelNewestQuarter(sheet, A股上市公司s, A股上市公司代码s, cfiVPage, year)
        writeToExcelHTables(workbook, False, A股上市公司s, A股上市公司代码s)
        writeToExcelHTables(workbook, True, A股上市公司s, A股上市公司代码s)

if __name__ == "__main__":
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename, askdirectory
    Tk().withdraw()
    
    print("请选择要下载中财网数据的公司列表（该文件由QCC_IPO产生）")
    qcc_company_names_file = askopenfilename()
    print("请选择保存位置")
    outputDir = askdirectory()
    print("请输入你需要过去多少年的数据:")
    numYearsToQuery = int(input())
    print("请输入重大事件需要过去多少年的数据:")
    numYearsToQuery_重大事件 = int(input())
    downloadExcelFromCFI(qcc_company_names_file, outputDir, numYearsToQuery, numYearsToQuery_重大事件)
