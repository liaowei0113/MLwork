from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter,time


def urlOpen(url):
    res = urllib.request.Request(url)
    res.add_header('User-Agent',
                   'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/'
                   '537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36')
    req = urllib.request.urlopen(res)
    html = req.read()
    print("++++++")
    bs = BeautifulSoup(html, 'html.parser')

    table = bs.find_all('div', {'class': 'api_month_list'})[0].find('table')
    tr_lists = table.find_all('tr')
    return tr_lists

# 写入数据
def writeData(WorkSheet,row,col):
    # 数据时间有重复，用作去重
    time_list = []
    for i in range(1, len(tr_lists)):
        td_lists = tr_lists[i].find_all('td')
        # 用作去重
        flag = 1
        for j in range(len(td_lists)):
            content = td_lists[j].get_text().lstrip()
            # 去重
            if j == 0 and content in time_list:
                flag = 0
                break
            time_list.append(content)
            WorkSheet.write(row, col, content)  # 表格第row行第col列
            col += 1
        col = 0
        if flag == 1:
            row += 1

years = [2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020]
months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
headUrl = "http://www.tianqihoubao.com/aqi/chongqing-"


CQQualityExcel = xlsxwriter.Workbook('重庆空气质量.xlsx')
WorkSheet = CQQualityExcel.add_worksheet()
row = 1
col = 0
WorkSheet.write(0, col, '日期')  # 表格第一行
WorkSheet.write(0, col + 1, '质量等级')
WorkSheet.write(0, col + 2, 'AQI')
WorkSheet.write(0, col + 3, 'PM2.5')
WorkSheet.write(0, col + 4, 'PM10')
WorkSheet.write(0, col + 5, 'SO2')
WorkSheet.write(0, col + 6, 'NO2')
WorkSheet.write(0, col + 7, 'CO')
WorkSheet.write(0, col + 8, 'O3')

for year in years:
    for month in months:
        URL = headUrl + str(year) + month + '.html'
        tr_lists = urlOpen(URL)
        # 数据时间有重复，用作去重
        time_list = []
        for i in range(1, len(tr_lists)):
            td_lists = tr_lists[i].find_all('td')
            # 用作去重
            flag = 1
            for j in range(len(td_lists)):
                if j==3:
                    # 不写入排名
                    continue
                content = td_lists[j].get_text().lstrip()
                # 去重
                if j == 0 and content in time_list:
                    flag = 0
                    break
                time_list.append(content)
                WorkSheet.write(row, col, content)  # 表格第row行第col列
                col += 1
            col = 0
            if flag == 1:
                row += 1
        print("已写入"+str(year)+month+"的数据++++++++++++++++++")
        # time.sleep(3)


CQQualityExcel.close()
