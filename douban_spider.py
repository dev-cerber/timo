import requests
from lxml import etree
import xlwt
import re

# 要访问的地址
# 未登录状态获取数据受限，需登录后拷贝cookie
url = "https://movie.douban.com/subject/1303967/"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36',
    'Cookie': 'bid=cISTxamqZx4; __utmc=30149280; __utmz=30149280.1632654805.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); ll="108288"; __utmz=223695111.1632654807.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); __utmc=223695111; __gads=ID=0d3e97fa31910b1e-2256505febcb00fe:T=1632654809:RT=1632654809:S=ALNI_Mb-YhHaLeblpoTYEYjZxYJ4u2lF6A; _vwo_uuid_v2=D4605DF7A3F3F6BF497F2B01AF70DEAED|878f40575bedeb04d67e9798dbb6c101; __yadk_uid=A3rNILD6PVzPl6pBHBY2l3WhOLFBbNKo; dbcl2="247328215:KJUIbBeOj64"; ck=Brqk; push_noty_num=0; push_doumail_num=0; __utmv=30149280.24732; ap_v=0,6.0; __utmb=30149280.0.10.1632663981; __utma=30149280.1911116475.1632654805.1632657647.1632663981.3; __utma=223695111.991280750.1632654807.1632657647.1632663981.3; __utmb=223695111.0.10.1632663981; _pk_ses.100001.4cf6=*; _pk_id.100001.4cf6=1a3464067bab4ed1.1632654805.3.1632664033.1632662076.'
}


resp = requests.get(url=url, headers=headers)
wb_data = resp.content.decode()

html = etree.HTML(wb_data)
html_data = html.xpath('//*[@id="content"]/h1/span[1]')
rating_num = html.xpath('//*[@id="interest_sectl"]/div[1]/div[2]/strong')
comments_url = html.xpath('//*[@id="comments-section"]/div[1]/h2/span/a/@href')
counters = html.xpath('//*[@id="comments-section"]/div[1]/h2/span/a/text()')
counters_num = int(re.findall('\d+', counters[0])[0])
# counters_num = 40
if counters_num > 1000:
    counters_num = 1000
data_list = []


def sub_net(index):
    global data_list
    params = {
        "percent_type": '',
        'start': index * 20,
        'limit': 20,
        'status': 'P',
        'sort': 'new_score',
        'comments_only': 1
    }

    resp = requests.get(url=comments_url[0], headers=headers, params=params)
    final_html = resp.json().get('html')
    final_htm = etree.HTML(final_html)
    comment_divs = final_htm.xpath('//*[@class="comment-item "]')
    username = comment_divs[0].xpath('//div[@class="comment"]//span[@class="comment-info"]/a/text()')
    rating = comment_divs[0].xpath('//div[@class="comment"]//span[@class="comment-info"]/span[2]/@title')
    comment_time = comment_divs[0].xpath('//div[@class="comment"]//span[@class="comment-info"]/span[@class="comment-time "]/@title')
    comment = comment_divs[0].xpath('//div[@class="comment"]/p/span/text()')
    each = zip(username, comment_time, rating, comment)
    print('获取数据中...', username, comment_time, rating, comment)
    data_list.extend(list(each))


def write2execl(rows_num, data_list):
    excelpath = ('./douban_data.xls')  # 新建excel文件
    workbook = xlwt.Workbook(encoding='utf-8')  # 写入excel文件
    sheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)
    headlist = [u'序号', u'评论人', u'评论时间', u'评分', u'评论']  # 写入数据头
    row = 0
    col = 0
    for head in headlist:
        sheet.write(row, col, head)
        col = col + 1
    for i in range(1, rows_num + 1):

        username = data_list[i-1][0]
        comment_time = data_list[i-1][1]
        rating = data_list[i-1][2]
        comment = data_list[i-1][3]

        sheet.write(i, 0, i)
        sheet.write(i, 1, username)
        sheet.write(i, 2, comment_time)
        sheet.write(i, 3, rating)
        sheet.write(i, 4, comment)
        workbook.save(excelpath)  # 保存


if __name__ == '__main__':
    print("任务开始执行...")
    times = (counters_num // 20)
    for index in range(times):
        try:
            sub_net(index)
        except:
            print("网页访问受限")
    try:
        write2execl(counters_num, data_list)
    except:
        pass
    print("任务执行完成...")

