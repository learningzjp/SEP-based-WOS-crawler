# -*- coding: utf-8 -*-
"""
Created on Thu Apr  7 16:14:56 2022

@author: zjp
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common import exceptions

def rename_file(SAVE_TO_DIRECTORY,name,record_format='excel'):
    #导出文件重命名
    #SAVE_TO_DIRECTORY -> 导出记录存储位置(文件夹)； 
    #name -> 重命名为
    while True:
        files = list(filter(lambda x:'savedrecs' in x and len(x.split('.'))==2,os.listdir(SAVE_TO_DIRECTORY)))
        if len(files)>0:
            break
    files = [os.path.join(SAVE_TO_DIRECTORY, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getctime(x))
    newest_file = files[-1]
    if record_format=='excel':
        os.rename(newest_file, os.path.join(SAVE_TO_DIRECTORY, name+".xls"))
    else:
        os.rename(newest_file, os.path.join(SAVE_TO_DIRECTORY, name+".txt"))
        
def send_key(browser,path,value):
    # browser -> browser
    # path -> css选择器
    # value -> 填入值
    markto=browser.find_element_by_css_selector(path)
    markto.clear()
    markto.send_keys(value)
    
def iselement(browser, ids):
    # 实现判断元素是否存在
    # param browser: 浏览器对象
    # param ids: id表达式
    # return: 是否存在
    try:
        browser.find_element_by_id(ids)
        return True
    except exceptions.NoSuchElementException:
        return False

def login(browser):
    # 通过SEP登录WOS
    # 此处填写sep用户名
    browser.find_element_by_id('userName1').send_keys('请填写自己的sep账号')
    time.sleep(0.5)
    # 此处填写sep密码
    browser.find_element_by_id('pwd1').send_keys('请填写自己的sep账号密码')
    time.sleep(0.5)
    browser.find_element_by_id('sb1').click()
    time.sleep(4)
    browser.find_element_by_xpath('/html/body/div[2]/ul/li[2]/a[4]').click()
    time.sleep(5) 
    
def startdownload(url,record_num,SAVE_TO_DIRECTORY,record_format='excel',reverse=False):
    # url -> 检索结果网址; 
    # record_num -> 需要导出的记录条数(检索结果数); 
    # SAVE_TO_DIRECTORY -> 记录导出存储路径(文件夹);
    # reverse -> 是否设置检索结果降序排列, default=False 
    fp = webdriver.FirefoxProfile()
    fp.set_preference('browser.download.dir', SAVE_TO_DIRECTORY)
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/plain")
    browser = webdriver.Firefox(executable_path=r'geckodriver',firefox_profile=fp)
    browser.get(url)
    time.sleep(4)
    login(browser)
    browser.get(url) # 登陆后会跳转SEP文献下载页面,这里直接重新打开检索结果页面
    time.sleep(10)
    #关闭WOS常见的一些弹窗
    for num in range(0,3):
        if iselement(browser,'pendo-close-guide-7176fce7'):
            browser.find_element_by_id('pendo-close-guide-7176fce7').click()
            time.sleep(2) 
        if iselement(browser,'pendo-close-guide-8be2c6ae'):
            browser.find_element_by_id('pendo-close-guide-8be2c6ae').click()
            time.sleep(2)    
        if iselement(browser,'pendo-close-guide-1d939896'):
            browser.find_element_by_id('pendo-close-guide-1d939896').click()
            time.sleep(2) 
    time.sleep(3)
    # 获取需要导出的文献数量
    record_num1 = int(browser.find_element_by_css_selector('.brand-blue').text.replace(",",""))
    if record_num >= record_num1:
        record_num = record_num1
    # 按时间降序排列
    if reverse: 
        browser.find_element_by_css_selector('.top-toolbar wos-select:nth-child(1) button:nth-child(1) span:nth-child(2)').click()
        browser.find_element_by_css_selector("div.wrap-mode:nth-child(2) span:nth-child(1)").click() 
        time.sleep(3)
    # 叉掉获取cookie的弹窗。
    if iselement(browser,'onetrust-accept-btn-handler'):
        browser.find_element_by_id('onetrust-accept-btn-handler').click()
        time.sleep(2) 
    # 开始导出
    start = 1 # 起始记录
    i = 0 # 导出记录的数字框id随导出次数递增
    flag = 1 # mac文件夹默认有一个'.DS_Store'文件
    while start<record_num:
        ele = browser.find_element_by_xpath('/html/body/app-wos/div/div/main/div/div[2]/app-input-route/app-base-summary-component/div/div[2]/app-page-controls[1]/div/app-export-option/div/app-export-menu/div/button')# 导出
        browser.execute_script('arguments[0].click()',ele)
        time.sleep(3)
        if record_format=='excel':
            browser.find_element_by_xpath('//*[@id="exportToExcelButton"]').click() # 选择导出格式为excel
            time.sleep(1)
            browser.find_element_by_css_selector('#radio3 label:nth-child(1) span:nth-child(1)').click() # 选择自定义记录条数
            send_key(browser,'#mat-input-%d'%i,start)#mat-input-2
            send_key(browser,'#mat-input-%d'%(i+1),start+999)
            browser.find_element_by_css_selector('.margin-top-5 button:nth-child(1)').click() # 更改导出字段
            browser.find_element_by_css_selector('div.wrap-mode:nth-child(3) span:nth-child(1)').click() # 选择所需字段(excel:3完整/4自定义; txt:3完整/4完整+引文)
            browser.find_element_by_css_selector('div.flex-align:nth-child(3) button:nth-child(1)').click() # 点击导出
            while len(os.listdir(SAVE_TO_DIRECTORY))==flag:
                time.sleep(3) # 等待下载完毕
            # 导出文件按照包含的记录编号重命名
            rename_file(SAVE_TO_DIRECTORY,'record-'+str(start)+'-'+str(start+999),record_format=record_format)
            start = start + 1000
            time.sleep(3)
        else:
            browser.find_element_by_css_selector('#exportToFieldTaggedButton').click() # 选择导出格式为txt
            time.sleep(1)
            browser.find_element_by_css_selector('#radio3 label:nth-child(1) span:nth-child(1) span:nth-child(1)').click() # 选择自定义记录条数
            send_key(browser,'#mat-input-%d'%i,start)#mat-input-2
            send_key(browser,'#mat-input-%d'%(i+1),start+499)
            browser.find_element_by_css_selector('.margin-top-5 button:nth-child(1)').click() # 更改导出字段
            browser.find_element_by_css_selector('div.wrap-mode:nth-child(4) span:nth-child(1)').click() # 选择所需字段(excel:3完整/4自定义; txt:3完整/4完整+引文)
            browser.find_element_by_css_selector('div.flex-align:nth-child(3) button:nth-child(1)').click() # 点击导出
            while len(os.listdir(SAVE_TO_DIRECTORY))==flag:
                time.sleep(1) # 等待下载完毕
            # 导出文件按照包含的记录编号重命名
            rename_file(SAVE_TO_DIRECTORY,'record-'+str(start)+'-'+str(start+499),record_format=record_format)
            start = start + 500
        i = i + 2
        flag = flag + 1

    time.sleep(10)
    browser.quit()
    
if __name__=='__main__':

    # WOS“检索结果”页面的网址，将检索地址放到这里
    url = 'https://libyw.ucas.ac.cn/https/1syHScyMuiZyDype4n6jffBUif9YVRKxLJkemI2Ixa8d1/wos/alldb/summary/ddf8d5bb-e4bf-4160-a75d-6c0be9648450-2f44d3ff/relevance/1'
    # 导出到本地的存储路径(自行修改)
    download_path = r'C:\Users\zjp\Desktop\test'
    #下载函数
    startdownload(url,10000, download_path,record_format='excel',reverse=False) # 主要函数
    print('Done')
    
