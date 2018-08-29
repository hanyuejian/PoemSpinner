# coding=utf-8
import sys
import urllib2
from bs4 import BeautifulSoup
import xlwt

reload(sys)
sys.setdefaultencoding('utf-8')


def geturl(url):
    response = urllib2.urlopen(url).read()
    response = unicode(response, 'utf-8').encode('utf-8')
    soup = BeautifulSoup(response, 'html.parser')
    return soup


def decode():
    url_list = []
    for i in range(1, 1001):
        print "正在解析" + str(i) + "页"
        url = "https://www.gushiwen.org/shiwen/default_0A0A" + str(i) + ".aspx"
        soup = geturl(url)
        poem_list = soup.findAll('a', attrs={"style": "font-size:18px; line-height:22px; height:22px;"})

        for item in poem_list:
            if poem_list is not None:
                url_list.append(item.get('href'))
    return url_list


def spinner():
    wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = wb.add_sheet('sheet1', cell_overwrite_ok=True)
    url_list = decode()
    for i in range(0, len(url_list)):
        href_soup = geturl(url_list[i])
        print "正在爬取" + str(i) + "个"
        name = href_soup.find('h1',
                              attrs={"style": "font-size:20px; line-height:22px; height:22px; margin-bottom:10px;"})
        if name is not None:
            sheet.write(i, 0, name.text.strip())
        info = href_soup.find('p', attrs={'class': "source"})
        if info is not None:
            message = info.text.strip().split("：")
            if len(message) > 1:
                sheet.write(i, 1, message[0])
                sheet.write(i, 2, message[1])
        content = href_soup.find('div', attrs={'class': "contson"})
        if content is not None:
            sheet.write(i, 3, content.text.strip())
        tag = href_soup.find('div', attrs={'class': "tag"})
        if tag is not None:
            sheet.write(i, 4, tag.text.strip())
        wb.save('poem.xlsx')


spinner()
