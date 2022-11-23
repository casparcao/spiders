import requests
from lib import headers as hd
from pymongo import MongoClient
from lxml import etree
from datetime import datetime


# 赛程页面的url
schedule_url_template = "https://www.jianlizhi.com/#scb"
mongo_client = MongoClient()
db = mongo_client.wager
collection = db.game


def crawls():
    # 下载
    content = downloads()
    # 提取信息
    schedule = extracts(content)
    # 保存数据
    save(schedule)


def extracts(content):
    """
    提取页面中的信息，并转为实体对象
    :param content: 页面内容
    :return: 实体对象
    """
    html = etree.HTML(content)
    daily_parts = html.xpath('//div[@class="word_cup_saicheng"]')
    games = []
    count = 0
    for daily_part in daily_parts:
        extract0(games, daily_part, count)
        count = len(games)
    return games


def downloads():
    """
    下载指定页面
    :return: 返回页面内容
    """
    response = requests.get(schedule_url_template, headers={'User-Agent': hd.agent()})
    content = response.text
    return content


def extract0(games, daily_part, count):
    """
    每日赛程列表解析，每天大概有4场比赛
    :param games
    :param daily_part:
    :param count
    :return:
    """
    date_h3 = daily_part.xpath('div[@id="word_cup_riqi"]/h3/text()')
    if len(date_h3) <= 0:
        return
    date = date_h3[0]
    game_parts = daily_part.xpath('div[@id="word_cup_gm"]')
    ic = count
    for game_part in game_parts:
        time = game_part.xpath('div[@class="time"]/span[@class="word_cup_time"]/text()')
        title = game_part.xpath('div[@class="time"]/span[@class="word_cup_leixing1"]/a/text()')
        turns = game_part.xpath('div[@class="time"]/span[@class="word_cup_changci"]/text()')
        home = game_part.xpath('div[@class="qiudui"]/span[@class="word_cup_jiemu_left"]/a/text()')
        guest = game_part.xpath('div[@class="qiudui"]/span[@class="word_cup_jiemu_right"]/a/text()')
        home_icon = game_part.xpath('div[@class="qiudui"]/span[@class="word_cup_zpic"]/img/@src')
        guest_icon = game_part.xpath('div[@class="qiudui"]/span[@class="word_cup_kpic"]/img/@src')
        result = game_part.xpath('div[@class="qiudui"]/span[@class="word_cup_vs"]/text()')
        status = game_part.xpath('div[@class="zhibo"]/span[@class="word_cup_links"]/a/i/text()')
        time = time[0] if time else ""
        title = title[0] if title else ""
        turns = turns[0] if turns else ""
        home = home[0] if home else ""
        guest = guest[0] if guest else ""
        home_icon = home_icon[0] if home_icon else ""
        guest_icon = guest_icon[0] if guest_icon else ""
        result = result[0] if result else ""
        status = status[0] if status else ""
        ic = ic + 1
        game = {
            "_id": ic,
            "date": datetime.strptime("2022年" + date[:7] + time + ":00", '%Y年%m月%d日 %H:%M:%S'),
            "title": title + turns,
            "home": {
                "name": home,
                "icon": home_icon,
                "score": int(result[0]) if "已完赛" == status else 0
            },
            "guest": {
                "name": guest,
                "icon": guest_icon,
                "score": int(result[2:]) if "已完赛" == status else 0
            },
            "result": result,
            "status": status,
            "clear": False
        }
        games.append(game)
        print(game)


def save(info_list):
    """
    保存到csv文件
    :param info_list: 小区的字典列表
    :return: none
    """
    print(info_list)
    if info_list:
        collection.insert_many(info_list)


if __name__ == '__main__':
    crawls()
    # with ThreadPoolExecutor(max_workers=16) as pool:
    #     pool.submit(crawls)
