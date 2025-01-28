# First
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
from pathlib import Path
import winreg
from selenium.webdriver import EdgeOptions
from selenium.webdriver import Edge
from selenium.webdriver.edge.service import Service
import driver_download
import keyboard


class Create_1:
    
    def __init__(self) -> None:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        desktop_path_1 = winreg.QueryValueEx(key, "Desktop")[0]
        self.desktop_path = Path(desktop_path_1.replace("\\","/"))

    def start(self):#提示
        print("\n\n==========================================*^_^*微信公众号链接采集*^_^*============================================")
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

    def first_check_file_and_dir(self):
        path = self.desktop_path / 'web spider'
        list_file =["all of url","mid_with_title","title_with_time"]
        if not path.exists():
            path.mkdir()
        else:
            # 更新
            path_drive = ''
            # 检测如果存在多个, 该如何选择. 采用通配符模式
            path_1 = path / '数据_公众号and类别.json'
            path_2 = path / '数据_公众号信息.json'
            while not path_1.exists() and not path_2.exists():
                print('关键json文件未找到,请添加后继续程序')

            for file_i in list_file:
                path_i = path / file_i
                if not path_i.exists():
                    path_i.mkdir()
                else:
                    for son_file in os.listdir(path_i):
                        os.remove(path_i / son_file)

    def time_limit(self):
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
        self.this_Wednesday = today_date - timedelta(days=DICT_WEEK[today_week])
        self.last_Thrusday = self.this_Wednesday - timedelta(days=6)


        def color(current):
            color_0 = ''
            color_1 = ''
            if current < 0:
                # 绿色
                color_0 = '\033[32m'
                color_1 = '\033[0m'
            elif current > 0:
                # 红色
                color_0 = '\033[31m'
                color_1 = '\033[0m'
            
            return color_0,color_1

        with open(f"{self.desktop_path / 'web spider'/ '采集时间.txt'}","w+") as get_limit_date:
            old_content = get_limit_date.read()
            get_limit_date.seek(0)
            get_limit_date.write(str(self.last_Thrusday)+ "--" + str(self.this_Wednesday) + "\n" + old_content)
        print("如果非此时间段,请通过左右键来切换日期, 成功后Enter")
        print(self.last_Thrusday,self.this_Wednesday, end='\r')


        def on_key_event(event):
            nonlocal current
            if event.name == 'left':
                current -= 1
                color_0 = color(current)[0]
                color_1 = color(current)[1]
                self.this_Wednesday = self.this_Wednesday - timedelta(days=7)
                self.last_Thrusday = self.last_Thrusday - timedelta(days=7)
                print(f'{color_0}{self.last_Thrusday} {self.this_Wednesday}{color_1}', end='\r')

            elif event.name == 'right':
                current += 1
                color_0 = color(current)[0]
                color_1 = color(current)[1]                
                self.this_Wednesday = self.this_Wednesday + timedelta(days=7)
                self.last_Thrusday = self.last_Thrusday + timedelta(days=7)
                print(f'{color_0}{self.last_Thrusday},{self.this_Wednesday}{color_1}', end='\r')
        
        current = 0
        keyboard.on_press(on_key_event)
        keyboard.wait('enter')


    def drive(self):
        url = "https://mp.weixin.qq.com"
        self.url_new = ''
        option = EdgeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        option.add_experimental_option('useAutomationExtension', False)
        # 优化
        driver_root = self.desktop_path / 'web spider'

        exe_list = list(driver_root.glob('*.exe'))
        try:
            exe_list.remove('edge_driver.exe')
        except:
            pass
        driver_path = exe_list[0]

        
        loop = 0
        can = 1
        while can:
            try:
                self.driver = Edge(service=Service(f"{driver_path}"),options=option)
                self.driver.get(url)
                can = 0 
            except:
                print("\033[31m浏览器驱动版本不正确, 尝试更新中...\033[0m")
                driver_download.main(self.desktop_path,loop)
                sleep(1)
                loop += 1

        #更新
        sleep(25)
        while(True):
            if(url == self.driver.current_url):
                sleep(1)
            else:
                self.url_new = self.driver.current_url
                break


    def get_url(self):

        self.token = ''
        self.cookies = ''
        self.agent = ''
        fragment = self.url_new.split('&')
        for i in fragment:
            if("token" in i):
                self.token = i.split("=")[1]
        for cookie in self.driver.get_cookies():
            self.cookies = self.cookies + cookie['name'] + '=' + cookie['value'] + ';'
        self.cookies = self.cookies.strip(';')

        self.agent = self.driver.execute_script("return navigator.userAgent;")

        file_agent = self.desktop_path / 'web spider' / 'agent.txt'
        if not file_agent.exists():
            with open(file_agent,encoding='utf-8') as agent_0:
                agent_0.write(self.agent)

        self.driver.quit()

        with open(f"{self.desktop_path / 'web spider'/ '数据_公众号信息.json'}","r",encoding="utf-8") as file:
            self.dict_fakeid = load(file)
            #在这里添加


    def backbone(self):

        for key in range(0,len(self.dict_fakeid)):
            account = list(self.dict_fakeid.keys())[key]

            sleep(3)

            print(f"{account},总计还剩{len(self.dict_fakeid)-key-1}个公众号")

            link_list = []
            dict_title_with_time = {}
            dict_mid_with_time = {}

            each_fakeid = self.dict_fakeid[account]

            page = 100 
            bool_date_limit = True

            for each_page in range(0,page):

                # url = f"https://mp.weixin.qq.com/cgi-bin/appmsg?action=list_ex&begin={each_page*5}&count=5&fakeid={each_fakeid}&type=9&query=&token={token}&lang=zh_CN&f=json&ajax=1"
                url = f"https://mp.weixin.qq.com/cgi-bin/appmsgpublish?sub=list&search_field=null&begin={each_page*5}&count=5&query=&fakeid={each_fakeid}&type=101_1&free_publish_type=1&sub_action=list_ex&token={self.token}&lang=zh_CN&f=json&ajax=1"
                header ={
                    'User-Agent': self.agent,
                    'Cookie':self.cookies,
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
                    print("请求过于频繁,等待继续. 若是本周初次执行,说明程序寄了")
                    suspend_time = 5000 
                    for suspend in range(suspend_time, -1, -1):
                        sleep(1)
                        print("等待中...{:>4}秒".format(suspend), end="\r")
                    print("\r")
                    processed_dict['publish_page']['publish_list']

                for i in processed_dict['publish_page']['publish_list']:
                    ii = {key: convert_to_dict(value) for key, value in i.items()}
                    for son_list in ii['publish_info']['appmsgex']:
                        if son_list['is_deleted']:
                            continue
            
                        local_time = localtime(int(son_list["create_time"]))
                        formal_time = strftime("%Y-%m-%d %H:%M:%S", local_time).split(" ")[0]

                        if(formal_time > str(self.this_Wednesday)):
                            continue

                        if(str(self.last_Thrusday) <= formal_time):
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

            self.write_in(link_list,account,dict_mid_with_time,dict_title_with_time)

    def write_in(self,link_list,account,dict_mid_with_time,dict_title_with_time):
        if(link_list == []):
            print(f"{account}本周无文章发表\n")
        else:
            print(f"{account}已全部爬取完毕,正在写入文件中...")

            with open(f"{self.desktop_path / 'web spider' / 'all of url' / f'{account}_url.txt'}","w",encoding="UTF-8") as file_url:
                for link in link_list:
                    file_url.write(link+"\n")
            print(f"{account}的链接读写完成;",end="")
            with open(f"{self.desktop_path / 'web spider' / 'mid_with_title' / f'{account}.json'}","w") as f_mid_json:
                dump(dict_mid_with_time,f_mid_json)
            with open(f"{self.desktop_path / 'web spider' / 'title_with_time' / f'{account}.json'}","w") as f_title_json:
                dump(dict_title_with_time,f_title_json)
            print("相关文件读写完成.已全部保存到桌面web spider文件夹中\n")
class Create_2:
    def __init__(self) -> None:
        pass


#-------------------------------------------------------
def main():
    filterwarnings("ignore")

    
    create = Create_1()
    create.start()
    create.first_check_file_and_dir()
    create.time_limit()
    create.drive()
    create.get_url()
    create.backbone()

    os.system("pause")

if __name__=='__main__':
    main()