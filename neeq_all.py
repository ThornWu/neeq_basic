#coding=utf-8
import re
import json
import string
import urllib
import sys
import xlwt


reload(sys)
sys.setdefaultencoding('utf8')
count=0

def getHtml(url):
    global count
    html=urllib.urlopen(url).read().replace("null(","").replace(")","")
    if html !="null":
        count+=1
        formatString=json.loads(html)["baseinfo"]
        sheet.write(count, 0, (formatString['code'] if 'code' in formatString else "null"))
        sheet.write(count, 1, (formatString['shortname'] if 'shortname' in formatString else "null"))
        sheet.write(count, 2, (formatString['name'] if 'name' in formatString else "null"))
        sheet.write(count, 3, (formatString['address'] if 'address' in formatString else "null"))
        sheet.write(count, 4, (formatString['listingDate'] if 'listingDate' in formatString else "null"))
        sheet.write(count, 5, (formatString['area'] if 'area' in formatString else "null"))
        sheet.write(count, 5,(formatString['broker'] if 'broker' in formatString else "null"))
        sheet.write(count, 7,(formatString['website'] if 'website' in formatString else "null"))


workbook =xlwt.Workbook(encoding = 'utf-8')
sheet = workbook.add_sheet('data',cell_overwrite_ok=True)
style = xlwt.XFStyle()
font = xlwt.Font()
font.name = 'SimSun' # 指定“宋体”
style.font = font

sheet.write(0,0,"股票代码")
sheet.write(0,1,"股票名称")
sheet.write(0,2,"公司名称")
sheet.write(0,3,"注册地址")
sheet.write(0,4,"挂牌日期")
sheet.write(0,5,"所在地")
sheet.write(0,6,"主办券商")
sheet.write(0,7,"网址")


i=430000
while(i<430800):
    url="http://www.neeq.com.cn/nqhqController/detailCompany.do?zqdm="+str(i)
    getHtml(url)
    i+=1
workbook.save('all_company.xlsx')
print "430ok"

i=830000
while(i<840000):
    url="http://www.neeq.com.cn/nqhqController/detailCompany.do?zqdm="+str(i)
    getHtml(url)
    i += 1
workbook.save('all_company.xlsx')
print "830ok"

i=870000
while(i<872300):
    url="http://www.neeq.com.cn/nqhqController/detailCompany.do?zqdm="+str(i)
    getHtml(url)
    i += 1

workbook.save('all_company.xlsx')
print "870ok"
