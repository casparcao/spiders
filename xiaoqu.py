import requests
import re
from concurrent.futures import ThreadPoolExecutor
import threading
import headers as hd
from pymongo import MongoClient


# 小区页面的url
xiaoqu_url_template = "https://jn.ke.com/xiaoqu/pg{page}/"
mongo_client = MongoClient()
db = mongo_client.bk
collection = db.xiaoqu


def crawls(page):
    # 下载
    content = downloads(page)
    # 提取信息
    xiaoqu = extracts(content)
    # 保存数据
    save(xiaoqu)


def extracts(content):
    """
    提取页面中的信息，并转为实体对象
    :param content: 页面内容
    :return: 实体对象
    """
    xiaoqu_li_tag_list = re.findall('<li class="clear xiaoquListItem CLICKDATA"(.*?)>(.*?)</li>', content, re.S)
    xiaoqu_li_txt_list = [x[1] for x in xiaoqu_li_tag_list]
    xiaoqu = []
    for xiaoqu_text in xiaoqu_li_txt_list:
        xiaoqu.append(extract0(xiaoqu_text))
    return xiaoqu


def downloads(page):
    """
    下载指定页面
    :param page: 页码
    :return: 返回页面内容
    """
    response = requests.get(xiaoqu_url_template.format(page=str(page)), headers={'User-Agent': hd.agent()})
    content = response.text
    return content


def extract0(xiaoqu_txt):
    """
    提取小区基本信息
    :param xiaoqu_txt:
    :return:
    """
    match_price = re.search('<div class="totalPrice">.*?<span>([0-9]+)</span>', xiaoqu_txt, re.S)
    match_amount = re.search('<div class="xiaoquListItemSellCount">.*?<span>([0-9]+)</span>', xiaoqu_txt, re.S)
    match_link = re.search('<a class="img maidian-detail" href="(.*?)"', xiaoqu_txt, re.S)
    match_title = re.search('<a class="maidian-detail".*?>(.*?)</a>', xiaoqu_txt, re.S)
    info = {"title": match_title.group(1) if match_title else '',
            "link": match_link.group(1) if match_link else '',
            "no": re.search('[0-9]+', xiaoqu_txt, re.S).group(0),
            "district": re.search('<a href.*?class="district".*?>(.*?)</a>', xiaoqu_txt, re.S).group(1),
            "area": re.search('<a href.*?class="bizcircle".*?>(.*?)</a>', xiaoqu_txt, re.S).group(1),
            "price": match_price.group(1) if match_price else 0,
            "amount": match_amount.group(1) if match_amount else 0}
    info['layout'] = layout(info["link"])
    info['three'] = '3室' in info['layout']
    return info


def layout(link):
    """
    爬取户型数据
    :param link: 小区详情页链接
    :return:
    """
    response = requests.get(link, headers={'User-Agent': hd.agent()})
    content = response.text
    layout_li_tag_list = re.findall('<ol class="frameDealListItem">.*?<li>(.*?)</li>', content, re.S)
    return ";".join([parse_layout(x) for x in layout_li_tag_list])


def parse_layout(layout_li_tag):
    """
    处理每个户型标签数据，提取户型名称，大小
    :param layout_li_tag:  标签数据
    :return: 户型名称及大小
    """
    title = re.search('class="frameDealTitle">(.*?)</a>', layout_li_tag, re.S).group(1)
    square = re.search('class="frameDealArea">(.*?)平米</div>', layout_li_tag, re.S).group(1)
    return title + ":" + square


lock = threading.Lock()


def save(info_list):
    """
    保存到csv文件
    :param info_list: 小区的字典列表
    :return: none
    """
    with lock:
        # df = DataFrame(info_list)
        # index += 1
        # with pd.ExcelWriter("./districts/" + str(index) + ".xlsx") as writer:
        #     df.to_excel(writer, sheet_name="Sheet")
        # df.to_csv("./districts2.csv", encoding="utf-8")
        print("--------------------------------------------------")
        print(threading.current_thread().name)
        print("--------------------------------------------------")
        print(info_list)
        if info_list:
            collection.insert_many(info_list)


if __name__ == '__main__':
    with ThreadPoolExecutor(max_workers=16) as pool:
        for p in range(16):
            pool.submit(crawls, p)
