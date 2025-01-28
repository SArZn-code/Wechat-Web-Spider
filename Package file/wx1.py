from requests import get
from time import sleep
from time import strftime
from time import localtime
from datetime import timedelta
from datetime import date
from sys import exit
from json import load
from json import loads
from json import JSONDecodeError
from json import dump
from warnings import filterwarnings
import os
import winreg
from selenium.webdriver import EdgeOptions
from selenium.webdriver import Edge
from selenium.webdriver.edge.service import Service
filterwarnings("ignore")

print("\n\n")
print("==========================================*^_^*微信公众号链接采集*^_^*============================================")
print("请确定是否要执行该程序\n输入yes或no:", end="")
while(True):
    bl_1 = input().lower()
    if(bl_1 == "yes"):
        break
    elif(bl_1 == "no"):
        exit()
    else:
        print("无效输入，请重新确认是否要执行该程序:", end="")
print("==================================================================================================================")

key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
path_1 = winreg.QueryValueEx(key, "Desktop")[0]
path = path_1.replace("\\","/")

if not os.path.exists(f"{path}/web spider"):
    os.makedirs(f"{path}/web spider")
else:
    pass

today_date = date.today()
today_week = today_date.strftime("%A")
DICT_WEEK = {
    "Monday":5,
    "Tuesday":6,
    "Wednesday":0,
    "Thursday":1,
    "Friday":2,
    "Saturday":3,
    "Sunday":4,
}
this_Wednesday = today_date - timedelta(days=DICT_WEEK[today_week])
last_Thrusday = this_Wednesday - timedelta(days=6)
with open(f"{path}/web spider/采集时间.txt","w+") as get_limit_date:
    old_content = get_limit_date.read()
    get_limit_date.seek(0)
    get_limit_date.write(str(last_Thrusday)+ "--" + str(this_Wednesday) + "\n" + old_content)
print(last_Thrusday,this_Wednesday)
INPUT = input("请确认:Enter ; 如果非此时间段, 请输入对应时间段, 格式:xxxx-xx-xx;xxxx-xx-xx \n")
if INPUT == '':
    pass
else:
    last_Thrusday, this_Wednesday = INPUT.split(";")

print("开始清除上次爬取的文件")
sleep(2)
print("清除完毕")

url = "https://mp.weixin.qq.com"
url_new = ''
option = EdgeOptions()
option.add_experimental_option('excludeSwitches', ['enable-automation'])
option.add_experimental_option('useAutomationExtension', False)
driver = Edge(service=Service(f"{path}/web spider/msedgedriver.exe"),options=option)
driver.get(url)
sleep(25)
while(True):
    if(url == driver.current_url):
        sleep(1)
    else:
        url_new = driver.current_url
        break
token = ''
cookies = ''
agent = ''
fragment = url_new.split('&')
for i in fragment:
    if("token" in i):
        token = i.split("=")[1]
for cookie in driver.get_cookies():
    cookies = cookies + cookie['name'] + '=' + cookie['value'] + ';'
cookies = cookies.strip(';')
try:
    file = open(f"{path}/web spider/agent.txt","r",encoding='utf-8')
    agent = file.readlines()[0]
except:
    agent = driver.execute_script("return navigator.userAgent;")
    with open(f"{path}/web spider/agent.txt","w",encoding='utf-8') as file:
        file.write(agent)
driver.quit()

list_file =[f"{path}/web spider/all of url",f"{path}/web spider/mid_with_title",f"{path}/web spider/title_with_time"]
for file_i in list_file:
    if not os.path.exists(file_i):
        os.makedirs(file_i)
        continue
    else:
        file_list = os.listdir(file_i)
        for file_name in file_list:
            file_path = os.path.join(file_i, file_name)
            os.remove(file_path)

with open(f"{path}/web spider/数据_公众号信息.json","r",encoding="utf-8") as file:
    dict_fakeid = load(file)
    # 在这里添加
    
    
for key in range(0,len(dict_fakeid)):
    account = list(dict_fakeid.keys())[key]

    sleep(3)

    print(f"{account},总计还剩{len(dict_fakeid)-key-1}个公众号")

    link_list = []
    dict_title_with_time = {}
    dict_mid_with_time = {}

    each_fakeid = dict_fakeid[account]

    page = 100 
    bool_date_limit = True

    for each_page in range(0,page):

        # url = f"https://mp.weixin.qq.com/cgi-bin/appmsg?action=list_ex&begin={each_page*5}&count=5&fakeid={each_fakeid}&type=9&query=&token={token}&lang=zh_CN&f=json&ajax=1"
        url = f"https://mp.weixin.qq.com/cgi-bin/appmsgpublish?sub=list&search_field=null&begin={each_page*5}&count=5&query=&fakeid={each_fakeid}&type=101_1&free_publish_type=1&sub_action=list_ex&token={token}&lang=zh_CN&f=json&ajax=1"
        header = {
            'User-Agent': agent,
            'Cookie':cookies,
        }
        
        result = get(url, headers=header, verify=False)
        receive_dict = result.json()
        def convert_to_dict(value):
            try:
                return loads(value)
            except (JSONDecodeError, TypeError):
                return value  
        processed_dict = {key: convert_to_dict(value) for key, value in receive_dict.items()}

        try:
            processed_dict['publish_page']['publish_list']
        except KeyError:
            print("请求过于频繁,等待继续")
            suspend_time = 5000 
            for suspend in range(suspend_time, -1, -1):
                sleep(1)
                print("等待中...{:>4}秒".format(suspend), end="\r")
            print("\r")
            processed_dict['publish_page']['publish_list']

        for i in processed_dict['publish_page']['publish_list']:

            ii = {key: convert_to_dict(value) for key, value in i.items()}

            for son_list in ii['publish_info']['appmsgex']:
                # 优化 - 删除
                if son_list['is_deleted']:
                    continue
                
                local_time = localtime(int(son_list["create_time"]))
                formal_time = strftime("%Y-%m-%d %H:%M:%S", local_time).split(" ")[0]

                if(formal_time > str(this_Wednesday)):
                    continue

                if(str(last_Thrusday) <= formal_time):
                    pass
                else:
                    bool_date_limit = False

                if(bool_date_limit == True):
                    pass
                else:
                    break
                
                link = son_list["link"]
                title = son_list["title"]
                # mid = link.split("&")[1].split("=")[1]
                mid = son_list["aid"]
                link_list.append(link)
                dict_mid_with_time[mid] = title 
                dict_title_with_time[mid + '|' +title] = formal_time

                sleep_time = 5 
                for remaining_time in range(sleep_time, -1, -1):
                    sleep(1)
                    print(f"等待爬取下一篇文章中...{remaining_time}秒", end="\r")
                print("\r")
            if(bool_date_limit == True):
                pass
            else:
                break

        if(bool_date_limit == True):
            pass
        else:
            break
        sleep_time = 10 
        for remaining_time in range(sleep_time, -1, -1):
            sleep(1)
            print("等待爬取下一页中...{:>2}秒".format(remaining_time), end="\r")
        print("\r")


    if(link_list == []):
        print(f"{account}本周无文章发表\n")
    else:
        print(f"{account}已全部爬取完毕,正在写入文件中...")

        with open(f"{path}/web spider/all of url/{account}_url.txt","w",encoding="UTF-8") as file_url:
            for link in link_list:
                file_url.write(link+"\n")
        print(f"{account}的链接读写完成;",end="")
        with open(f"{path}/web spider/mid_with_title/{account}.json","w") as f_mid_json:
            dump(dict_mid_with_time,f_mid_json)
        with open(f"{path}/web spider/title_with_time/{account}.json","w") as f_title_json:
            dump(dict_title_with_time,f_title_json)
        print("相关文件读写完成.已全部保存到桌面web spider文件夹中\n")
os.system("pause")
#################################################################################################################################