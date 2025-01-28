import json
import winreg
from pathlib import Path

#如果有数据请从这里修改

def sheetall():
    dict_fakeid_all = {
        "OUC就业直通车": "MzA4ODUzNzUyNA==",
        "中国海大研究生": "MzA3MjQxMzgzMg==",
        "中国海洋大学校友会": "MzIzODU4MjExNQ==",
        "中国海洋大学信息学部": "MzA4ODE3OTMyMg==",
        "中国海大文新学院": "MjM5Mzk2NjczMQ==",
        "OUC法学院": "MzU3OTczNTQxMQ==",
        "中国海大地学院": "MzU1OTg3OTUwNQ==",
        "中国海大海气学院": "MzI3MjA0MzUxOA==",
        "海大食品食光机": "MzA3NzEzMTQwMA==",
        "OUCFLCSTU": "MzA5Njk4MDQxMw==",
        "中国海大管理力量": "MzA4MzU5OTEyOA==",
        "中海大化学人": "MzA3ODYyNTQwMQ==",
        "中国海大海洋生命学院": "MzA3NTg4NDg2Ng==",
        "OUC筑梦经济": "MzU4OTYwMDY2Mw==",
        "中国海洋大学环科院": "MzUyMzQ0Njg4MA==",
        "OUC基础教学中心": "Mzg4NTU3Mjc1Nw==",
        "OUC小药丸": "MzAwNzE0MTM5OQ==",
        "海大水产之声": "MzIyOTc2Nzk1Mg==",
        "嗨 大工程": "MzkwNzQxNTk4OQ==",
        "中国海洋大学心理中心": "MzU3MTcwMDAxNQ==",
        "中国海大材料学院": "Mzg2MTkzMTEwNQ==",
        "麦思微斯基": "MzA5NDQyNDczNg==",
        "OUC海德团小青": "MzkyNzMzMzMyOQ==",
        "中国海洋大学未来海洋学院": "MzkyMzMzMjM3NQ==",
        "中国海洋大学海德学院": "MzAwMjc5ODk2MA==",
        "OUC国管人": "MzU3MjczMzQyNA==",
        "海鸥剧社": "MzA3MTcxNDM0NA==",
        "中国海洋大学崇本学院": "MzA3NjQ5NTc1Mw==",
        "海大经济学院研究生": "MjM5Nzg1NDIyNQ==",
        "海大电影课": "MzI0NzY4Njc4OQ==",
        "中国海洋大学信息青年": "MjM5MTk2Nzg4MQ==",
        "OUC思源": "MzU2MTcxMzI4NQ==",
        "中国海大马院": "Mzg3Mzc1MDM2OQ==",
        "中国海大研究生会": "MzA4MTYwODAzOA==",
        "OUC文新研会": "MzU1OTIxNjIwMw==",
        "新闻传播OUC": "MzIxMzI4NDM0NA==",
        "小文谈": "MzIxNDY5ODMwMg==",
        "Dongxiangxing": "MzAwNjAzNTI0Nw==",
        "中海大语办": "MzkzMTAyNDE1Mg==",
        "躲进故事的角落": "MzUzOTU1ODgyMw==",
        "悦享政治学": "MzIwMzM3MTk2OQ==",
        "OUC浩海": "MzIzNTI5NTQyNw==",
        "海大博雅文学社": "MzA3MTA2MDQzNQ==",
        "Math先锋": "MzAxODA3NzAwMA==",
        "OUC理论学习研究会": "Mzg5NDMxMDE3OA==",
        "祥说近代史": "MzI4MDA4OTcxMw==",
        "海大科幻": "MzIwMTExODUyMQ==",
        "OUC田径队": "MzI3MTUyNjQ2NQ==",
        "中海大工程自强社": "MzA4NTQ1NzM0Mw==",
        "海岩社": "MjM5NDc4NjA1OA==",
        "海大国际教育中心": "MzI2ODI3MDQ2Nw==",
        "新长城中海大自强社": "MjM5MzgzMjE5NQ==",
        "金融俱乐部ouc": "MzI3MTg2NTQyNw==",
        "OUC文新自强社": "MzA3MzUzNzg1NA==",
        "OUCbaseball": "MjM5Mjc3NTc2Nw==",
        "海大爱心社": "MzA3MDQ4MTMwNQ==",
        "海大英语角OUTALK": "MzA3MzcyMzMzNA==",
        "OUC求是学会": "MzAxNTAwNDk5NA==",
        "中海大环协": "MzA4MTA3ODM2Ng==",
        "ouc大艺团": "MzAxMTM4NjI2Mw==",
        "ouc法协": "MzI1MjA3MDg2OA==",
        "海大工院科协": "MzI0NDY1ODYxNQ==",
        "OUC环保人": "MzA5MzA5MDA5Nw==",
        "OUC海之心志愿服务队": "MzIxMzI1ODc4NQ==",
        "鱼山科协": "MjM5MzQ4NzM4MQ==",
        "海大小助": "MzA5MDQ0MDA3Mw==",
        "经济先锋": "MzIwOTA4NDI0MA==",
        "海大基础教学中心艺术系": "MjM5MjU5MzY5OQ==",
        "海大国学社": "MzAxNTM1MDE0NQ==",
        "中海大外院自强社": "MzI4MDA3ODgyNA==",
        "海大国防生": "MzA3ODQwMTczMQ==",
        "中海大港澳台文化交流": "MzU0MDQ0NzI0Nw==",
        "海大百科": "MjM5OTExNjc2MA==",
        "骨语文创": "MzI3MDY3NTQ3Mw==",
        "海大化院研究会": "MzIxNDU4MzM4NA==",
        "化院自强社": "MzIzODQ0NjgxNg==",
        "中国海大后勤": "MjM5NTkxMzI4NA==",
        "缘定海大": "MzIwMzU3OTM5Ng==",
        "海大网球": "MzI0MjUzNzQ3Ng==",
        "樱海击剑社": "MzIwMzg2NDM0Ng==",
        "中国海大环境学子": "MzA5NDYyNDAxMg==",
        "海大定向运动俱乐部": "MzI5NDAyMTU4OA==",
        "中国海洋大学明勤社": "MzA4MDA2OTc4NQ==",
        "中国海大学生公共部门发展研究会": "MzU0NDU0MDU4OQ==",
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号信息.json","w",encoding="utf-8") as file:
        json.dump(dict_fakeid_all ,file)
    dict_fakeid = {
    "OUC就业直通车": "校务机构",
    "中国海大研究生": "校务机构",
    "中国海洋大学校友会": "校务机构",
    "中国海洋大学信息学部": "院系学生会",
    "中国海大文新学院": "院系学生会",
    "OUC法学院": "院系学生会",
    "中国海大地学院": "院系学生会",
    "中国海大海气学院": "院系学生会",
    "海大食品食光机": "院系学生会",
    "OUCFLCSTU": "院系学生会",
    "中国海大管理力量": "院系学生会",
    "中海大化学人": "院系学生会",
    "中国海大海洋生命学院": "院系学生会",
    "OUC筑梦经济": "院系学生会",
    "中国海洋大学环科院": "院系学生会",
    "OUC基础教学中心": "院系学生会",
    "OUC小药丸": "院系学生会",
    "海大水产之声": "院系学生会",
    "嗨 大工程": "院系学生会",
    "中国海洋大学心理中心": "校务机构",
    "中国海大材料学院": "院系学生会",
    "麦思微斯基": "院系学生会",
    "OUC海德团小青": "院系学生会",
    "中国海洋大学未来海洋学院": "校务机构",
    "中国海洋大学海德学院": "院系学生会",
    "OUC国管人": "院系学生会",
    "海鸥剧社": "社团",
    "中国海洋大学崇本学院": "院系学生会",
    "海大经济学院研究生": "院系学生会",
    "海大电影课": "自媒体",
    "中国海洋大学信息青年": "院系学生会",
    "OUC思源": "自媒体",
    "中国海大马院": "院系学生会",
    "中国海大研究生会": "校务机构",
    "OUC文新研会": "院系学生会",
    "新闻传播OUC": "院系学生会",
    "小文谈": "院系学生会",
    "Dongxiangxing": "社团",
    "中海大语办": "自媒体",
    "躲进故事的角落": "自媒体",##
    "悦享政治学": "社团",
    "OUC浩海": "社团",
    "海大博雅文学社": "社团",
    "Math先锋": "院系学生会",
    "OUC理论学习研究会": "院系学生会",
    "祥说近代史": "自媒体",
    "海大科幻": "社团",
    "OUC田径队": "自媒体",
    "中海大工程自强社": "院系学生会",
    "海岩社": "社团",
    "海大国际教育中心": "校务机构",
    "新长城中海大自强社": "社团",
    "金融俱乐部ouc": "社团",
    "OUC文新自强社": "社团",
    "OUCbaseball": "社团",
    "海大爱心社": "社团",
    "海大英语角OUTALK": "社团",
    "OUC求是学会": "社团",
    "中海大环协": "社团",
    "ouc大艺团": "社团",
    "ouc法协": "社团",
    "海大工院科协": "社团",
    "OUC海之心志愿服务队": "院系学生会",
    "鱼山科协": "校务机构",
    "海大小助": "院系学生会",
    "经济先锋": "社团",
    "OUC环保人": "院系学生会",
    "海大基础教学中心艺术系": "社团",
    "海大国学社": "社团",
    "中海大外院自强社": "院系学生会",##
    "海大国防生": "社团",
    "中海大港澳台文化交流": "社团",
    "海大百科": "自媒体",
    "骨语文创": "自媒体",
    "海大化院研究会": "院系学生会",
    "化院自强社": "院系学生会",##
    "中国海大后勤": "校务机构",
    "缘定海大": "自媒体",
    "海大网球": "社团",
    "樱海击剑社": "社团",
    "中国海大环境学子": "院系学生会",
    "海大定向运动俱乐部": "社团",
    "中国海洋大学明勤社": "社团",
    "中国海大学生公共部门发展研究会": "社团",
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号and类别.json","w",encoding="utf-8") as file:
        json.dump(dict_fakeid,file)

def sheet5():
    dict_sheet_5 = {
        "中国海大文新学院": "MjM5Mzk2NjczMQ==",
        "中国海大海气学院": "MzI3MjA0MzUxOA==",
        "OUCFLCSTU": "MzA5Njk4MDQxMw==",
        "中国海大管理力量": "MzA4MzU5OTEyOA==",
        "中国海洋大学心理中心": "MzU3MTcwMDAxNQ==",
        "中国海大后勤": "MjM5NTkxMzI4NA==",
        "新闻传播OUC": "MzIxMzI4NDM0NA==",
        "缘定海大": "MzIwMzU3OTM5Ng==",
        "海大网球": "MzI0MjUzNzQ3Ng==",
        "樱海击剑社": "MzIwMzg2NDM0Ng==",
        "中国海大环境学子": "MzA5NDYyNDAxMg==",
        "海大定向运动俱乐部": "MzI5NDAyMTU4OA==",
        "中国海洋大学明勤社": "MzA4MDA2OTc4NQ==",
        "Dongxiangxing": "MzAwNjAzNTI0Nw==",
        "中国海洋大学校友会": "MzIzODU4MjExNQ==",
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号信息.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_5 ,file)

    dict_sheet_5_category = {
        "中国海大文新学院": "院系学生会",
        "中国海大海气学院": "院系学生会",
        "OUCFLCSTU": "院系学生会",
        "中国海大管理力量": "院系学生会",
        "中国海洋大学心理中心": "校务机构",
        "中国海大后勤": "校务机构",
        "新闻传播OUC": "院系学生会",
        "缘定海大": "自媒体",
        "海大网球": "社团",
        "樱海击剑社": "社团",
        "中国海大环境学子": "院系学生会",
        "海大定向运动俱乐部": "社团",
        "中国海洋大学明勤社": "社团",
        "Dongxiangxing": "社团",
        "中国海洋大学校友会": "校务机构",
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号and类别.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_5_category,file)

def sheet4():
    dict_sheet_4 = {
        "中国海大研究生": "MzA3MjQxMzgzMg==",
        "OUC筑梦经济": "MzU4OTYwMDY2Mw==",
        "中海大化学人": "MzA3ODYyNTQwMQ==",
        "中国海洋大学未来海洋学院": "MzkyMzMzMjM3NQ==",
        "麦思微斯基": "MzA5NDQyNDczNg==",
        "中国海大材料学院": "Mzg2MTkzMTEwNQ==",
        "中国海大地学院": "MzU1OTg3OTUwNQ==",
        "OUC思源": "MzU2MTcxMzI4NQ==",
        "小文谈": "MzIxNDY5ODMwMg==",
        "OUC浩海": "MzIzNTI5NTQyNw==",
        "悦享政治学": "MzIwMzM3MTk2OQ==",
        "中海大港澳台文化交流": "MzU0MDQ0NzI0Nw==",
        "海大百科": "MjM5OTExNjc2MA==",
        "骨语文创": "MzI3MDY3NTQ3Mw==",
        "海大化院研究会": "MzIxNDU4MzM4NA==",
        "化院自强社": "MzIzODQ0NjgxNg=="
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号信息.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_4 ,file)

    dict_sheet_4_category = {
    "中国海大研究生": "院系学生会",
    "OUC筑梦经济": "院系学生会",
    "中海大化学人": "院系学生会",
    "中国海洋大学未来海洋学院": "校务机构",
    "麦思微斯基": "院系学生会",
    "中国海大材料学院": "院系学生会",
    "中国海大地学院": "院系学生会",
    "OUC思源": "自媒体",
    "小文谈": "院系学生会",
    "OUC浩海": "社团",
    "悦享政治学": "社团",
    "中海大港澳台文化交流": "社团",
    "海大百科": "自媒体",
    "骨语文创": "自媒体",
    "海大化院研究会": "院系学生会",
    "化院自强社": "院系学生会"
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号and类别.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_4_category,file)

def sheet3():
    dict_sheet_3 = {
    "OUC就业直通车": "MzA4ODUzNzUyNA==",
    "OUC法学院": "MzU3OTczNTQxMQ==",
    "海大经济学院研究生": "MjM5Nzg1NDIyNQ==",
    "OUC国管人": "MzU3MjczMzQyNA==",
    "中国海洋大学崇本学院": "MzA3NjQ5NTc1Mw==",
    "中国海洋大学信息青年": "MjM5MTk2Nzg4MQ==",
    "OUC海之心志愿服务队": "MzIxMzI1ODc4NQ==",
    "OUC小药丸": "MzAwNzE0MTM5OQ==",
    "海大食品食光机": "MzA3NzEzMTQwMA==",
    "鱼山科协": "MjM5MzQ4NzM4MQ==",
    "海大小助": "MzA5MDQ0MDA3Mw==",
    "经济先锋": "MzIwOTA4NDI0MA==",
    "OUC环保人": "MzA5MzA5MDA5Nw==",
    "海大基础教学中心艺术系": "MjM5MjU5MzY5OQ==",
    "海大国学社": "MzAxNTM1MDE0NQ==",
    "中海大外院自强社": "MzI4MDA3ODgyNA==",
    "海大国防生": "MzA3ODQwMTczMQ=="
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号信息.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_3 ,file)

    dict_sheet_3_category = {
    "OUC就业直通车": "校务机构",
    "OUC法学院": "院系学生会",
    "海大经济学院研究生": "院系学生会",
    "OUC国管人": "院系学生会",
    "中国海洋大学崇本学院": "院系学生会",
    "中国海洋大学信息青年": "院系学生会",
    "OUC海之心志愿服务队": "院系学生会",
    "OUC小药丸": "院系学生会",
    "海大食品食光机": "院系学生会",
    "鱼山科协": "校务机构",
    "海大小助": "院系学生会",
    "经济先锋": "社团",
    "OUC环保人": "院系学生会",
    "海大基础教学中心艺术系": "社团",
    "海大国学社": "社团",
    "中海大外院自强社": "院系学生会",
    "海大国防生": "社团"
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号and类别.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_3_category,file)

def sheet2():
    dict_sheet_2 = {
    "海大电影课": "MzI0NzY4Njc4OQ==",
    "海大水产之声": "MzIyOTc2Nzk1Mg==",
    "中国海大海洋生命学院": "MzA3NTg4NDg2Ng==",
    "OUC基础教学中心": "Mzg4NTU3Mjc1Nw==",
    "中海大语办": "MzkzMTAyNDE1Mg==",
    "金融俱乐部ouc": "MzI3MTg2NTQyNw==",
    "中国海洋大学海德学院": "MzAwMjc5ODk2MA==",
    "OUC文新自强社": "MzA3MzUzNzg1NA==",
    "OUCbaseball": "MjM5Mjc3NTc2Nw==",
    "海大爱心社": "MzA3MDQ4MTMwNQ==",
    "海大英语角OUTALK": "MzA3MzcyMzMzNA==",
    "OUC求是学会": "MzAxNTAwNDk5NA==",
    "中海大环协": "MzA4MTA3ODM2Ng==",
    "中国海大马院": "Mzg3Mzc1MDM2OQ==",
    "ouc大艺团": "MzAxMTM4NjI2Mw==",
    "ouc法协": "MzI1MjA3MDg2OA==",
    "海大工院科协": "MzI0NDY1ODYxNQ==",
    "中国海大研究生会": "MzA4MTYwODAzOA==",
    "中国海大学生公共部门发展研究会": "MzU0NDU0MDU4OQ==",
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号信息.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_2 ,file)

    dict_sheet_2_category = {
    "海大电影课": "自媒体",
    "海大水产之声": "院系学生会",
    "中国海大海洋生命学院": "院系学生会",
    "OUC基础教学中心": "院系学生会",
    "中海大语办": "自媒体",
    "金融俱乐部ouc": "社团",
    "中国海洋大学海德学院": "院系学生会",
    "OUC文新自强社": "社团",
    "OUCbaseball": "社团",
    "海大爱心社": "社团",
    "海大英语角OUTALK": "社团",
    "OUC求是学会": "社团",
    "中海大环协": "社团",
    "中国海大马院": "院系学生会",
    "ouc大艺团": "社团",
    "ouc法协": "社团",
    "海大工院科协": "社团",
    "中国海大研究生会": "校务机构",
    "中国海大学生公共部门发展研究会": "社团",
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号and类别.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_2_category,file)

def sheet1():
    dict_sheet_1 = {
    "嗨 大工程": "MzkwNzQxNTk4OQ==",
    "中国海洋大学环科院": "MzUyMzQ0Njg4MA==",
    "OUC海德团小青": "MzkyNzMzMzMyOQ==",
    "海鸥剧社": "MzA3MTcxNDM0NA==",
    "OUC文新研会": "MzU1OTIxNjIwMw==",
    "躲进故事的角落": "MzUzOTU1ODgyMw==",
    "Math先锋": "MzAxODA3NzAwMA==",
    "OUC理论学习研究会": "Mzg5NDMxMDE3OA==",
    "祥说近代史": "MzI4MDA4OTcxMw==",
    "海大科幻": "MzIwMTExODUyMQ==",
    "OUC田径队": "MzI3MTUyNjQ2NQ==",
    "中海大工程自强社": "MzA4NTQ1NzM0Mw==",
    "海岩社": "MjM5NDc4NjA1OA==",
    "海大国际教育中心": "MzI2ODI3MDQ2Nw==",
    "中国海洋大学信息学部": "MzA4ODE3OTMyMg==",
    "海大博雅文学社": "MzA3MTA2MDQzNQ==",
    "新长城中海大自强社": "MjM5MzgzMjE5NQ=="
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号信息.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_1 ,file)

    dict_sheet_1_category = {
    "嗨 大工程": "院系学生会",
    "中国海洋大学环科院": "院系学生会",
    "OUC海德团小青": "院系学生会",
    "海鸥剧社": "社团",
    "OUC文新研会": "院系学生会",
    "躲进故事的角落": "自媒体",
    "Math先锋": "院系学生会",
    "OUC理论学习研究会": "院系学生会",
    "祥说近代史": "自媒体",
    "海大科幻": "社团",
    "OUC田径队": "自媒体",
    "中海大工程自强社": "院系学生会",
    "海岩社": "社团",
    "海大国际教育中心": "校务机构",
    "中国海洋大学信息学部": "院系学生会",
    "海大博雅文学社": "社团",
    "新长城中海大自强社": "社团"
    }
    with open(f"{Desktop_path_web_spider}/数据_公众号and类别.json","w",encoding="utf-8") as file:
        json.dump(dict_sheet_1_category,file)


if __name__ =='__main__':

    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    desktop_path_1 = winreg.QueryValueEx(key, "Desktop")[0]
    Desktop_path = Path(desktop_path_1.replace("\\","/"))
    Desktop_path_web_spider = Desktop_path / 'web spider'

    D = {
        'all':sheetall,
        '1':sheet1,
        '2':sheet2,
        '3':sheet3,
        '4':sheet4,
        '5':sheet5,
    }
    
    while True:
        sheet = input(f"请输入:{' '.join(list(D.keys()))}\n")
        if sheet in D.keys():
            break
        else:
            print('无效输入!')
            
    D[sheet]()