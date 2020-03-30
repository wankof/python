from selenium import webdriver
import json,time
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
import lxml
from bs4 import BeautifulSoup
import os

# 初始化浏览器，并设置浏览器默认下载地址"D:\python数据库"

options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups':0,'download.default_directory':r'd:\python数据库'}
options.add_experimental_option('prefs',prefs)
wd = webdriver.Chrome('E:/chromedriver.exe',chrome_options=options)



wd.get(r'https://web.zhsmjxc.com/UCenter-webapp/Login/Init.htm')


username = wd.find_element_by_id('username')
password = wd.find_element_by_id('userpassword')
username.send_keys('18970812655')
password.send_keys('JXGZ2019\n')
time.sleep(2)

# ------------------------------------------------------
#               关闭悬浮窗口
# ------------------------------------------------------
# 方法一：F12在console中输入JavaScript语句找到是第10个参数，定位后关闭悬浮窗口
# js = "document.getElementsByClassName('closeButton')[10].click()"
# wd.execute_script(js)

# 方法二：定位文本内容为"确认"的元素，点击
wd.find_element_by_link_text('确认').click()


# ------------------------------------------------------
# 由selenium的ActionChains类完成模拟鼠标操作
# ------------------------------------------------------
from selenium.webdriver.common.action_chains import ActionChains

ActionChains(wd).move_to_element(wd.find_element_by_link_text('分析')).perform()
ok = wd.find_element_by_link_text('销售明细表')
ok.click()


xf = wd.find_element_by_xpath('(//*[@id="iframe"])[2]')
wd.switch_to.frame(xf)
sleep(1)

js = "$('input[id=end]').attr('readonly',false)"
wd.execute_script(js)


# ------------------------------------------------------
# 创建月份循环列表
# ------------------------------------------------------
import datetime
import calendar


def get_time_range_list(startdate, enddate):
    """
    获取时间参数列表
    :param startdate: 起始月初时间 --> str
    :param enddate: 结束时间 --> str
    :return: date_range_list -->list
    """
    date_range_list = []
    startdate = datetime.datetime.strptime(startdate,'%Y-%m-%d')
    enddate = datetime.datetime.strptime(enddate,'%Y-%m-%d')
    while 1:
        next_month = startdate + datetime.timedelta(days=calendar.monthrange(startdate.year, startdate.month)[1])
        month_end = next_month - datetime.timedelta(days=1)
        if month_end < enddate:
            date_range_list.append((datetime.datetime.strftime(startdate,
                                                               '%Y-%m-%d'),
                                    datetime.datetime.strftime(month_end,
                                                               '%Y-%m-%d')))
            startdate = next_month
        else:
            return date_range_list


days = get_time_range_list('2017-06-01','2020-04-01')


# -----------------------------------------------------
#               清空文件夹下的文件
# -----------------------------------------------------

path = 'D:/python数据库'
for i in os.listdir(path):
   path_file = os.path.join(path,i)
   if os.path.isfile(path_file):
      os.remove(path_file)
   else:
     for f in os.listdir(path_file):
         path_file2 =os.path.join(path_file,f)
         if os.path.isfile(path_file2):
             os.remove(path_file2)

# -----------------------------------------------------
#               输入2017-6-1至2017-6-30日，下载数据
# -----------------------------------------------------
for star_day,end_day in days:

    wd.find_element_by_xpath('//*[@id="startdate-fornormal"]').clear()
    wd.find_element_by_xpath('//*[@id="startdate-fornormal"]').send_keys(star_day)
    wd.find_element_by_xpath('//*[@id="enddate-fornormal"]').clear()
    wd.find_element_by_xpath('//*[@id="enddate-fornormal"]').send_keys(end_day)
    sleep(2)
    wd.find_element_by_xpath('//*[@id="enddate-fornormal"]').send_keys(Keys.ENTER)

    wd.find_element_by_link_text('导出').click()
    wd.find_element_by_link_text('导出数据').click()

    WebDriverWait(wd,30).until(EC.presence_of_element_located((By.XPATH,'//span[contains(text(), "下载文件结束")]')))
    print('%s 至 %s 数据已下载' % (star_day,end_day))
    sleep(3)
# path为批量文件的文件夹的路径
    path = 'D:/python数据库/'

    # 文件夹中所有文件的文件名
    f = os.listdir(path)

    # 外循环遍历所有文件名，内循环遍历每个文件名的每个字符
    for oldname in f:
        if '销售明细统计' in oldname:
            print(oldname)
            path_oldname = path + oldname
            path_newname = path + star_day + '至' + end_day + '销售明细' + '.xls'
            os.rename(path_oldname,path_newname)
            print(path_oldname,'重命名为：',path_newname)



wd.close()