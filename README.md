## Auto-StudyStrongCountry 
## 自动强国项目（仅网页端），扫码登陆，自动阅读文章，看视频，当日积分查询

### 前言
这个项目主要是最近接触了学习强国，对于刷分比较反感，做的一点便于摸鱼的微小工作。
主要思路来自于 https://blog.csdn.net/qq_38636998/article/details/89279195 这篇文章，做了一些小改动。py小白，第一次认真写代码，主要还是使用了python+selenium的方法，供各位学习参考使用。
### 目录
- 运行环境
- 使用方法
- 各模块组成
- 版本记录
- 存在问题记录
### 运行环境
- Windows 10 
- python 3.7 
- anaconda 3 （64-bit）
- Google Chorme  75.0.3770.142 
- Chromedriver （正常应该找对应版本的，不过我使用的是 https://chromedriver.storage.googleapis.com/index.html?path=74.0.3729.6/；
2019.08.08更新：找到一个可以下载完整版本的地址http://chromedriver.storage.googleapis.com/index.html ，各位按需下载即可)
### 使用方法
- V0.1版本
  - 安装完Google Chrome和Chromedriver，Python后，运行代码即可
### 模块组成
- login_simulation 模拟登录模块
- read_articles 阅读文章模块
- watch_videos 观看视频模块
- get_scores 查询积分模块
### 版本记录
- V0.1版本  2019.07.28  
  - 使用网络上的代码修改后编写，主要是修改了read_articles()和watch_videos()两个模块。
### 存在问题记录
- V0.1版本
  - 阅读文章、观看视频有时候会点开往日已经看过的文章
  - 观看视频时有时视频播完了还在窗口等待
  - 没有完整观看到有效时间
  - 要自己先安装谷歌浏览器和Chromedriver，代码才能调用成功
