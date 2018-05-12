from openpyxl import load_workbook
from urllib.parse import quote  #字符串编码
import time

import requests
from bs4 import BeautifulSoup


# 获取网页页面
def getHTMLText(url, encoding='utf-8'):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36'
        }  # 构造头部
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        r.encoding = encoding
        return r.text
    except:
        return ''


# excel读写
def excelReader(file_name, sheet_name):
    wb = load_workbook(file_name)
    sh = wb.get_sheet_by_name(sheet_name)
    words = []
    for i in sh['A']:
        words.append(i.value)

    wordsWritter(words, sh)

    wb.save(file_name)
    wb.close()


# 主控
def wordsWritter(words, sh):
    row_index = 1
    count = 0
    for word in words:
        url = 'http://youdao.com/w/' + quote(word, 'utf-8') #字符串编码
        result = crawler(url)
        if result == []:
            count += 1
        col_index = 'B'
        for col in result:
            index = col_index + str(row_index)
            sh[index] = col
            col_index = chr(ord(col_index)+1)
        row_index = row_index + 1
        time.sleep(0.2) #TODO:缓一下，有空做多线程

    print('添加单词失败%s个', count)


# 爬虫
def crawler(url):
    try:
        result = []
        html = getHTMLText(url)
        soup = BeautifulSoup(html, "html.parser")

        #获取音标，英标和美标，和词意
        phrsListTab = soup.find(id='phrsListTab')
        phonetics = phrsListTab.find_all('span', attrs={'class', 'phonetic'})[0]
        # phonetic = ''
        # for i in phonetics:
        #     phonetic=phonetic+i.text+'\n'
        result.append(phonetics.text)

        means = phrsListTab.find('div', attrs={'class', 'trans-container'})
        mean = means.text.replace(' ', '').replace('\n\n', '\n')
        result.append(mean)

        #例句
        examples = soup.find(id='examples').ul
        [s.extract() for s in examples('p', attrs={'class', 'example-via'})]#去除例句来源
        sentences = examples.get_text().strip().replace("\n",'')#格式整理
        new = ''
        for char in sentences:
            if (char == '。'):
                new = new + '。\n'
            elif (char == '.'):
                new = new + '.\n'
            else:
                new = new + char
        result.append(new)

        return result
    except:
        print("出错")
        return []


if __name__ == '__main__':
    file_name = 'test.xlsx'
    sheet_name = 'Sheet1'
    excelReader(file_name, sheet_name)