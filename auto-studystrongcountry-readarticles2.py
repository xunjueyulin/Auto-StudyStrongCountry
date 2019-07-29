# _*_ coding: utf-8 _*_

from selenium import webdriver
import time


HOME_PAGE = 'https://www.xuexi.cn/' #定义主页
VIDEO_LINK = 'https://www.xuexi.cn/4426aa87b0b64ac671c96379a3a8bd26/db086044562a57b441c24f2af1c8e101.html#1novbsbi47k-5' #定义视频链接
LONG_VIDEO_LINK = 'https://www.xuexi.cn/f65dae4a57fe21fcc36f3506d660891c/b2e5aa79be613aed1f01d261c4a2ae17.html'
LONG_VIDEO_LINK2 = 'https://www.xuexi.cn/0040db2a403b0b9303a68b9ae5a4cca0/b2e5aa79be613aed1f01d261c4a2ae17.html'
TEST_VIDEO_LINK = 'https://www.xuexi.cn/8e35a343fca20ee32c79d67e35dfca90/7f9f27c65e84e71e1b7189b7132b4710.html'
SCORES_LINK = 'https://pc.xuexi.cn/points/my-points.html'
LOGIN_LINK = 'https://pc.xuexi.cn/points/login.html'
ARTICLES_LINK = 'https://www.xuexi.cn/72ac54163d26d6677a80b8e21a776cfa/9a3668c13f6e303932b5e0e100fc248b.html'

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-automation'])
browser = webdriver.Chrome(options=options)
#browser = webdriver.Chrome(executable_path=r'D:\OneDrive\Python\selenium\chromedriver.exe',options=options)


def login_simulation():
    """模拟登录"""
    # 方式一：使用cookies方式
    # 先自己登录，然后复制token值覆盖
    # cookies = {'name': 'token', 'value': ''}
    # browser.add_cookie(cookies)

    # 方式二：自己扫码登录
    browser.get(LOGIN_LINK)
    browser.maximize_window()  # 窗口最大化
    browser.execute_script("var q=document.documentElement.scrollTop=1000")
    time.sleep(20)
    browser.get(HOME_PAGE)
    print("模拟登录完毕\n")

def read_articles():
    """阅读文章"""
    browser.get(ARTICLES_LINK)
    time.sleep(5)
    articles = set(browser.find_elements_by_class_name("text")) #获取文章连接,用set函数去重
    print(type(articles))
    print(articles)
    for index,article in enumerate(articles): # 遍历文章连接
        if index > 7: # 点击6个文章链接
            break            
        print(index,article)        
        article.click()
        # browser.get(browser.current_url)  #获取当前窗口连接
        time.sleep(5)
        # browser.close()
        print(browser.current_url)
    all_handles = browser.window_handles #获取当前窗口的句柄
    for handle in all_handles[1:]:
        browser.switch_to.window(handle) #切换到当前窗口
        browser.execute_script("var q=document.documentElement.scrollTop=100000")  # 窗口滑动到底部
        time.sleep(20)
        print(browser.current_url + "阅读完毕")
        browser.close()	
	# browser.switch_to.window(all_handles[-1])  #切换到倒数第一个窗口  
    # browser.execute_script("var q=document.documentElement.scrollTop=100000")  # 窗口滑动到底部
    # time.sleep(20)
    browser.switch_to.window(all_handles[0]) #回到第一个窗口
    # browser.get(HOME_PAGE)
    time.sleep(5)
    print("阅读文章完毕\n")

def watch_videos():
    """观看视频"""
    browser.get(VIDEO_LINK)
    time.sleep(5)
    #videos = browser.find_elements_by_xpath("//div[@id='Ck3ln2wlyg3k00']")
    videos = set(browser.find_elements_by_class_name("textWrapper")) # 获取视频链接，用set函数去重
    print(type(videos))
    print(videos)
    
    for i , video in enumerate(videos):  # 遍历视频链接
        if i > 7:  # 点击6个视频链接
            break
        print(i,video)
        video.click()
        time.sleep(5)
        print(browser.current_url)
    all_handles = browser.window_handles # 获取当前窗口的句柄
    for handle in all_handles[1:]:  # 对除第一个窗口句柄以外的句柄进行操作
        browser.switch_to.window(handle) # 切换到当前窗口
        video_duration_str = browser.find_element_by_class_name("duration").get_attribute('innerText')  #获取视频时长的字段内容，几分几秒，这里用find_elements方法会报错
        video_duration = int(video_duration_str.split(':')[0])* 60 + int(video_duration_str.split(':')[1])  # 将时长转换成秒数
        time.sleep(video_duration + 3) # 保持窗口到视频时长结束
        print(browser.current_url + '观看完毕')
        browser.close()
    browser.switch_to.window(all_handles[0]) # 回到第一个窗口
    time.sleep(5)
    print("观看视频完毕\n")


	
def get_scores():
    """获取当前积分"""
    browser.get(SCORES_LINK)
    time.sleep(2)
    gross_score = browser.find_element_by_xpath("//*[@id='app']/div/div[2]/div/div[2]/div[2]/span[1]")\
        .get_attribute('innerText')
    today_score = browser.find_element_by_xpath("//span[@class='my-points-points']").get_attribute('innerText')
    print("当前总积分：" + str(gross_score))
    print("今日积分：" + str(today_score))
    print("获取积分完毕，即将退出\n")
	
if __name__ == '__main__':
    login_simulation()  # 模拟登录
    read_articles()     # 阅读文章
    watch_videos()      # 观看视频
    get_scores()        # 获得今日积分
    # browser.quit()	

	
