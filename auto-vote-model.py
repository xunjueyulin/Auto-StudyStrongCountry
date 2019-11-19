#!/usr/bin/env python 
# -*- coding:utf-8 -*-
from selenium import webdriver
import time
# import random

options = webdriver.ChromeOptions()
options.add_argument('--headless')  # 使用无头模式
options.add_argument('--disable-gpu')
# options.add_experimental_option('excludeSwitches',['enable-automation'])  # 添加实验性质的参数，好像不写也没关系

browser = webdriver.Chrome()
# 确定投票页面
VOTE_PAGE = 'http://data.zaihujk.com/f/GG0RAh'
i = 0
# 加入无头浏览模式
# 打开投票页面


def open_vote_page():
    browser.get(VOTE_PAGE)
    # browser.maximize_window()  # 窗口最大化
    time.sleep(5)
    print("打开投票页面完毕\n")

# 选择选项（随机选择其他）


def choice_option():
    browser.find_element_by_xpath("//div[contains(text(),'曾淑倩')]").click()  # 必选项
    time.sleep(3)
    print("第2项已选择")
    # num = random.choice(list(range(101,124)))
    # browser.find_element_by_xpath("//div[contains(text(),'中日友好')]").click()  # 在视频组选项里选一个作为搭配，防止暴露
    browser.find_element_by_xpath("//div[contains(text(),'梦诗')]").click()
    # print("第"+str(num)+"项作为搭配被选择，选择完毕")
    # 如果选择搭配这种方法不行,再想想
    # optionalchoices = browser.find_element_by_xpath()

# 提交表单


def submit_form():
    browser.find_element_by_xpath("//input[@value='提交']").click()   # 点击提交按钮
    time.sleep(5)
    # i = 0
    checksuccesses = browser.find_elements_by_class_name("message")   # 检查是否提交成功

    if len(checksuccesses)>0:
        # i = i + 1
        print("投票成功")
    else:
        print("投票不成功，请检查")

# 清理cookie并等待


def delete_cookie_and_quit():
    browser.delete_all_cookies()
    print("清理cookie完毕")
    time.sleep(5)
    # browser.close()   # 如果这里写了close（），使用selenium自动化的时候报错，
    # 错误提示：selenium.common.exceptions.WebDriverException: Message: invalid session id，通过对错误信息进行分析，
    # 无效的sessionid。后来通过对网上进行搜索查询，原因是在使用webdriver之前调用了driver.close()后
    # 将webdriver关闭了，则webdriver就失效了。


if __name__ == '__main__':
    #  一个python文件通常有两种使用方法，第一是作为脚本直接执行，第二是 import 到其他的 python 脚本中被调用（模块重用）执行。
    #  因此 if __name__ == 'main': 的作用就是控制这两种情况执行代码的过程，在 if __name__ == 'main': 下的代码只有在第一种
    #  情况下即文件作为脚本直接执行）才会被执行，而 import 到其他脚本中是不会被执行的。
    i = 0

    while i < 50:
        open_vote_page()
        choice_option()
        submit_form()
        i = i + 1
        print("第"+str(i)+"次投票成功")
        delete_cookie_and_quit()
        time.sleep(5)
    else:
        print("本次刷票完成")
        browser.quit()

