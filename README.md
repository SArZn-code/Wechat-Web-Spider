# Wechat-Web-Spider
This repository record multiple-file Web crawlers to collect Wetchat Official Accounts info including likes and views for Ocean University of China Community.
There are also some useful script like download edge browser agency automatically. 

language: Python & some Typescript
Third-party Tools: Packet capture, Fiddler

# Tree
```
│  agent.txt
│  all_data.txt
│  fiddler配置.png
│  msedgedriver.exe
│  py_2 need url.txt
│  read me.txt
│  一些参数.txt
│  数据_公众号and类别.json
│  数据_公众号信息.json
│  注意事项.txt
│  采集时间.txt
│
├─all of url   --> the temporary file for each essay url
├─mid_with_title  --> the temporary file for unique id of each essay with title 
├─Package file   --> reposit all script
│  │  01优化.txt
│  │  driver_download.py    --> download edge browser agency automatically
│  │  sheet_select.py       ---> select objected accounts classes
│  │  wx1.py
│  │  wx2.py
│  │  重写.py
│  │  重写2_浏览器可以打开了_测试版.py
│  │  重写2_浏览器打不开.py
│  └─  重写整合版~(不用了).py
|
└─title_with_time  --> the temporary file for each essay title with time 

重写*.py -> reconstruct code by Object-Oriented
all of the 1* file get info from browser
all of the 2* file get info from Wechat Client via Fiddler
```
# WorkFlow
1* file -> Fiddler -> 2* file
- 1* file get info for Wechat Account Officials
- Fiddler get info carried 1* output
- 2* file get core result carried Fiddler output
