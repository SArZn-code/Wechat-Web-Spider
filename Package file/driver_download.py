import requests
from bs4 import BeautifulSoup as BS
import re
import platform
import zipfile
import shutil
import os
from pathlib import Path
import winreg


def get_url(index):
    os_name = platform.system()
    #这里没有添加不同系统下的不同版本
    os_CPU = {'windows': 'x64', 'linux':'linux', 'mac':'mac', 'arm64':'arm64'}

    RE = re.compile(r'\d+.\d+.\d+.\d+')
    url = 'https://developer.microsoft.com/zh-cn/microsoft-edge/tools/webdriver/?form=MA13LH'
    response = requests.get(url)
    html = response.text

    soup = BS(html,'html.parser')
    # 第一个box的selector 选择器
    first_box = soup.select('#main > div > div.block-page.block-page--ready.block-page--theme-default > section:nth-child(5) > div.block-container__container.block-centered__container > div > div > div > div.block-centered__media > div > div > div > div > div.common-pager__pages > div > div > div > div:nth-child(1) > div')
                             
    a_box = first_box[index].select('a')

    CPU_version_link = []

    for i in a_box:
        ii = i.select_one('span').text.strip()
        href = i.get('href')
        result = RE.search(href)
        version = result.group()
        if os_CPU[os_name.lower()] == ii.lower():
            CPU_version_link.append([ii, version, href])

    latest = CPU_version_link[0]
    link = latest[2]

    print(f'Latest\n{os_name}----{latest[0]}\n')
    print(f'Version:  {latest[1]}')
    print(f'Link:  {link}\n')

    return link

def download(Desktop_path,link,zip_path):
    
    response = requests.get(link,stream=True)
    # 字节
    per_size = 1024*1024
    total_size = int(response.headers['content-length'])
    print('文件大小 -> %.2f Mb' %(total_size / per_size))

    # 进度条
    size = 0
    with open(zip_path, 'wb') as file:
        for byte in response.iter_content(chunk_size=per_size):
            file.write(byte)
            size += len(byte)

            a = 100 * size / total_size # 100 指的是进度条的总数
            print('%s %.2f' %('|'*int(a), a),end='\r')
        print('\n')

    #解压
    zip = zipfile.ZipFile(zip_path)

    RE = re.compile(r'.*\.exe$')
    for i in zip.namelist():
        result = RE.search(i)
        try:
            exe = result.group()
        except:
            pass
    temp = Desktop_path / '1'
    zip.extract(exe,temp)
    shutil.move(f'{temp / exe}', f'{Desktop_path /'web spider' / 'msedgedriver.exe'}')
    shutil.rmtree(temp)


def main(Desktop_path,index):
    link = get_url(index)
    zip_path = Desktop_path / '1.zip'

    count = 1
    while count < 4:
        try:
            download(Desktop_path,link,zip_path)
            break
        except:
            print('下载失败重试: 第 %s 次 / 共 3 次 \n' %count)
            count += 1
    else:
        print('\033[31m浏览器驱动下载失败, 请手动复制上方链接, 并解压里面的exe文件到web spider后再运行\033[0m')
        os.system('pause')

    os.remove(zip_path)

if __name__=='__main__':
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    desktop_path_1 = winreg.QueryValueEx(key, "Desktop")[0]
    Desktop_path = Path(desktop_path_1.replace("\\","/"))
    print('下载的是最新版本的浏览器驱动')
    main(Desktop_path,0)