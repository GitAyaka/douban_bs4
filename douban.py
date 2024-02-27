import csv
from bs4 import BeautifulSoup
import requests
import os
import xlsxwriter

# 设置headers 请求反爬
headers = {

    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.3 Safari/605.1.15"
}

# 结果写入csv文件
file = open('豆瓣电影Top250.csv', 'w', newline='', encoding='utf-8')
csvwriter = csv.writer(file)
csvwriter.writerow(['电影名称', '导演', '主演', '类型', '国家', '年份', '评分', '评语'])
oneresult = []  # 存一部电影的信息
result = []  # 存所有信息
count = [0] * 93  # 年产量

for i in range(10):
    url = 'https://movie.douban.com/top250?start=' + str(i*25) + '&filter='  # 访问豆瓣网站 网站分了十页 依次访问
    response = requests.get(url, headers=headers)
    response.encoding = 'utf-8'
    page_content = response.text
    soup = BeautifulSoup(response.text, 'html.parser')  # 访问网页源代码
    # 图片
    pics = soup.find_all('div', 'pic')
    imgs = []
    for pic in pics:
        imgs.append(pic.find('img').attrs['src'])
    for x in range(25):
        if not os.path.exists('images'):
            os.makedirs('images')
        path = 'images/{}.jpg'.format(i * 25 + x + 1)
        response = requests.get(imgs[x])
        with open(path, 'wb') as f:
            f.write(response.content)

    # 其他信息
    soup = soup.find('ol', class_='grid_view')
    for item in soup.find_all('div', 'info'):
        # 名称
        title = item.div.a.span.string
        # 导演 主演 年份 类型 国家
        line = item.find('div', 'bd').find('p').text.strip('').split('\n')  # 找到信息所在行
        directorline = line[1].strip('').split('\xa0\xa0\xa0')  # 找到导演
        yearline = line[2].strip('').split('\xa0/\xa0')  # 找到年份

        # 如果导演太多会导致主演没法显示，此时做省略处理
        if '\xa0\xa0\xa0' in line[1]:
            director = directorline[0].strip(' ')
            director = director.replace('导演: ', '')
            actors = directorline[1]
            actors = actors.replace('主演: ', '')
        else:
            director = directorline[0].strip('')
            actors = ''

        year = yearline[0].strip(' ')
        # 有些年份后跟了注释，统计年产量时只取前四位作年份
        if len(year) == 4:
            count[int(year) - 1930] += 1
        else:
            Year = year[0:4]
            count[int(Year) - 1930] += 1
        country = yearline[1].strip(' ')
        type = yearline[2].strip(' ')

        # 评分
        score = item.find('span', {'class':'rating_num'}).get_text()

        # 判断是否有评论并写入csv文件
        if item.find('span', class_="inq") != None:
            comment = item.find('span', class_="inq").text
            csvwriter.writerow([title, director, actors, type, country, year, score, comment])
        else:
            comment = ''
            csvwriter.writerow([title, director, actors, type, country, year, score, comment])

        oneresult = [title, director, actors, type, country, year, score, comment]
        result.append(oneresult)

# 创建Excel文件
workbook = xlsxwriter.Workbook('豆瓣电影Top250.xlsx')
sheet = workbook.add_worksheet('豆瓣电影Top250')
# 设置列宽
sheet.set_column('A:A', 25)
sheet.set_column('B:C', 50)
sheet.set_column('D:D', 25)
sheet.set_column('E:E', 35)
sheet.set_column('F:G', 10)
sheet.set_column('H:H', 100)
# 开始写入数据
col = ['电影名称', '导演', '主演', '类型', '国家', '年份', '评分', '评语']
for i in range(8):
    sheet.write(0, i, col[i])
for i in range(250):
    data = result[i]
    for j in range(8):
        sheet.write(i+1, j, data[j])
workbook.close()

# 统计年产量
workbook = xlsxwriter.Workbook('年产量.xlsx')
sheet = workbook.add_worksheet('豆瓣电影Top250')
sheet.write(0, 0, '年份')
sheet.write(0, 1, '电影数量')
p = 1
for i in range(93):
    if count[i] != 0:
        sheet.write(p, 0, i + 1930)
        sheet.write(p, 1, count[i])
        p += 1
workbook.close()
# 关闭csv文件
file.close()
print("over!")
