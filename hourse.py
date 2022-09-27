import requests
from multiprocessing.dummy import Pool
import csv
from lxml import html
from pymongo import MongoClient

mongo_client = MongoClient('mongodb://a:b@localhost:3717')

# 二手房页面的url
house_url_template = "https://jn.ke.com/ershoufang/{district}/pg{page}/"
districts = ['lixia', 'gaoxin', 'licheng']
headers = {
    'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE'
}


def download(district, page):
    """
    下载页面
    :param page: 页码
    :param district: 小区code
    :return:
    """
    response = requests.get(house_url_template.format(page=str(page), district=district), headers=headers)
    content = response.text
    xpath = html.fromstring(content)
    xpath.xpath()
    district_li_tag_list = re.findall('<li class="clear xiaoquListItem CLICKDATA"(.*?)>(.*?)</li>', content, re.S)
    district_li_txt_list = [x[1] for x in district_li_tag_list]
    districts = []
    for district_txt in district_li_txt_list:
        process(district_txt, districts)
    store(districts)


def process(district_txt, result_list):
    """
    提取小区基本信息
    :param district_txt:
    :param result_list:
    :return:
    """
    match_price = re.search('<div class="totalPrice">.*?<span>([0-9]+)</span>', district_txt, re.S)
    match_amount = re.search('<div class="xiaoquListItemSellCount">.*?<span>([0-9]+)</span>', district_txt, re.S)
    match_link = re.search('<a class="img maidian-detail" href="(.*?)"', district_txt, re.S)
    match_title = re.search('<a class="maidian-detail".*?>(.*?)</a>', district_txt, re.S)
    print(match_title)
    info = {"title": match_title.group(1) if match_title else '',
            "link": match_link.group(1) if match_link else '',
            "no": re.search('[0-9]+', district_txt, re.S).group(0),
            "district": re.search('<a href.*?class="district".*?>(.*?)</a>', district_txt, re.S).group(1),
            "area": re.search('<a href.*?class="bizcircle".*?>(.*?)</a>', district_txt, re.S).group(1),
            "price": match_price.group(1) if match_price else 0,
            "amount": match_amount.group(1) if match_amount else 0}
    info['layout'] = layout(info["link"])
    info['three'] = '3室' in info['layout']
    result_list.append(info)


def layout(link):
    """
    爬取户型数据
    :param link: 小区详情页链接
    :return:
    """
    response = requests.get(link, headers=headers)
    content = response.text
    layout_li_tag_list = re.findall('<ol class="frameDealListItem">.*?<li>(.*?)</li>', content, re.S)
    return ";".join([process_layout(x) for x in layout_li_tag_list])


def process_layout(layout_li_tag):
    """
    处理每个户型标签数据，提取户型名称，大小
    :param layout_li_tag:  标签数据
    :return: 户型名称及大小
    """
    title = re.search('class="frameDealTitle">(.*?)</a>', layout_li_tag, re.S).group(1)
    square = re.search('class="frameDealArea">(.*?)平米</div>', layout_li_tag, re.S).group(1)
    return title + ":" + square


def store(info_list):
    """
    保存到csv文件
    :param info_list: 小区的字典列表
    :return: none
    """
    with open('./districts.csv', 'a', encoding='utf-8') as f:
        writer \
            = csv.DictWriter(f, fieldnames=['title', 'link', 'no', 'district',
                                            'area', 'price', 'amount', 'layout', 'three'])
        writer.writeheader()
        writer.writerows(info_list)
    print(info_list)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pool = Pool(16)
    pool.map(download, range(100))
