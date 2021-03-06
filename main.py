import os

import requests
import time
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
import cv2
import numpy as np
from typing import List
from wxauto import *
from win32com.client import DispatchEx
import datetime
from PIL import ImageGrab


class yidun:
    def crack(self, target_img: str, template_img: str) -> List[int]:
        distance = self.match(target_img, template_img)
        real_distance = [int(distance * 360 / 480)]
        return real_distance

    def change_size(self, file):
        image = cv2.imread(file, 1)
        img = cv2.medianBlur(image, 5)
        b = cv2.threshold(img, 15, 255, cv2.THRESH_BINARY)
        binary_image = b[1]
        binary_image = cv2.cvtColor(binary_image, cv2.COLOR_BGR2GRAY)
        x, y = binary_image.shape
        edges_x = []
        edges_y = []
        for i in range(x):
            for j in range(y):
                if binary_image[i][j] == 255:
                    edges_x.append(i)
                    edges_y.append(j)

        left = min(edges_x)
        right = max(edges_x)
        width = right - left
        bottom = min(edges_y)
        top = max(edges_y)
        height = top - bottom
        pre1_picture = image[left: left + width, bottom: bottom + height]
        return pre1_picture

    def match(self, target, temp) -> int:
        img_gray = cv2.imread(target, 0)
        img_rgb = self.change_size(temp)
        template = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
        run = 1

        L = 0
        R = 1
        while run < 20:
            run += 1
            threshold = (R + L) / 2
            if threshold < 0:
                print("Error")
                return None
            loc = np.where(res >= threshold)
            if len(loc[1]) > 1:
                L += (R - L) / 2
            elif len(loc[1]) == 1:
                break
            elif len(loc[1]) < 1:
                R -= (R - L) / 2
        return loc[1][0]


if __name__ == '__main__':
    url = 'https://icas.jnu.edu.cn/cas/login?service=https://stuhealth.jnu.edu.cn/dashboard/cas'
    browser = webdriver.Chrome()
    browser.get(url)

    browser.set_window_size(2000, 1050)
    browser.set_window_position(-3000, 0)
    time.sleep(2)

    for i in range(1, 6):  # ???????????????
        try:
            bg_img_url = browser.find_element(by=By.XPATH,
                                              value='//*[@id="captcha"]/div/div[1]/div/div[1]/img[1]').get_attribute(
                "src")
            block_img_url = browser.find_element(by=By.XPATH,
                                                 value='//*[@id="captcha"]/div/div[1]/div/div[1]/img[2]').get_attribute(
                "src")

            bg = requests.get(bg_img_url)
            with open('./image/img1.png', 'wb') as f:
                f.write(bg.content)
            block = requests.get(block_img_url)
            with open('./image/img2.png', 'wb') as f:
                f.write(block.content)
            a = yidun()
            tracks = a.crack(r"image\img1.png", r"image\img2.png")
            slide = browser.find_element(By.CLASS_NAME, "yidun_slide_indicator")
            slider = browser.find_element(by=By.XPATH,
                                          value='//*[@id="captcha"]/div/div[2]/div[2]')
            ActionChains(browser).click_and_hold(slider).perform()
            while tracks:
                x = tracks.pop(0)
                ActionChains(browser).move_by_offset(xoffset=x, yoffset=0).perform()
            ActionChains(browser).release().perform()
            time.sleep(1)
            if slide.rect["width"] > 2:
                break
            else:
                continue
        except:
            pass
    time.sleep(1)

    # ?????????????????????
    browser.find_element(by=By.ID, value='un').send_keys(r'')  # ??????
    browser.find_element(by=By.ID, value='pd').send_keys(r'')  # ??????
    browser.find_element(by=By.ID, value='index_login_btn').click()
    time.sleep(3)
    cookies = browser.get_cookies()
    headers = {
        "accept": "application/json, text/plain, */*", "accept-encoding": "gzip, deflate, br",
        "accept-language": "zh,en;q=0.9,en-US;q=0.8,zh-CN;q=0.7", "content-length": "237",
        "content-type": "application/x-www-form-urlencoded",
        'cookie': 'JSESSIONID = ' + cookies[1]['value'] + ';vue_admin_template_token = ' + cookies[0]['value'],
        'x-token': cookies[0]['value'],
        "referer": "https://stuhealth.jnu.edu.cn/dashboard/",
        "sec-ch-ua": "\" Not A;Brand\";v=\"99\", \"Chromium\";v=\"101\", \"Google Chrome\";v=\"101\"",
        "sec-ch-ua-mobile": "?0", "sec-ch-ua-platform": "\"Windows\"", "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors", "sec-fetch-site": "same-origin",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/101.0.4951.54 Safari/537.36 "

    }

    ima = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))  # ?????????????????????
    ima_list = ima.split()
    ima_toki = ima_list[1].split(':')
    if int(ima_toki[0]) < 8:  # ??????0??????
        ima_toki[0] = str(24 - (8 - int(ima_toki[0])))
        t = ''
        for i in ima_toki:
            t += i
        ima_list[1] = t
        yesterday = datetime.date.today() - datetime.timedelta(days=1)
        ima_list[0] = str(yesterday)

    else:
        ima_toki[0] = str(int(ima_toki[0]) - 8)
        t = ''
        for i in ima_toki:
            t += i
        ima_list[1] = t

    data = {
        'declareTime': ima_list[0] + 'T' + ima_list[1] + '.000Z',  # ????????????
        'pycc': '',
        'collegeName': '????????????????????????/????????????????????????',  # ????????????????????????????????????
        'completed': 'false',
        'detailed': 'false'
    }
    download = requests.post(url='https://stuhealth.jnu.edu.cn/dashboard/excel/college', headers=headers, data=data)
    fp = open("today.xlsx", "wb")
    fp.write(download.content)
    fp.close()
    time.sleep(5)
    browser.quit()

    excel = DispatchEx("Excel.Application")  # ??????excel
    excel.Visible = False  # ???????????????

    wb = excel.Workbooks.Open(r'D:/yidun/today.xlsx', UpdateLinks=False, ReadOnly=False, Format=None,
                              Password='')  # ????????????
    sht = wb.Worksheets('sheet1')
    rows = sht.UsedRange.Rows.Count
    for i in range(0, 7):
        sht.Columns(3).Delete()  # ?????????
    for i in range(0, 6):
        sht.Columns(4).Delete()  # ?????????
    for i in range(0, rows):  # ?????????????????????
        if sht.Cells(rows - i, 3).Value != '????????????' and sht.Cells(rows - i, 3).Value != '????????????' and sht.Cells(rows - i,
                                                                                                           3).Value != '????????????????????????' or '2018' != sht.Cells(
                rows - i, 1).Value[0:4]:
            sht.Rows(rows - i).Delete()
    sht_range = "A1:C" + str(sht.UsedRange.Rows.Count)
    sht.Range(sht_range).CopyPicture()
    sht.Paste(sht.Range('D1'))
    sht.Shapes(sht.Shapes.Count).Copy()  # ??????????????????
    img = ImageGrab.grabclipboard()
    img.save(r"D:/yidun/test.png")
    wb.Save()
    wb.Close()

    # ???????????????????????????
    wx = WeChat()

    # ??????????????????
    wx.GetSessionList()

    # ?????????????????????
    msg = 'robot?????????????????????????????????'
    who = ''  # ????????????
    wx.ChatWith(who)  # ??????????????????
    wx.SendMsg(msg)  # ????????????

    # ?????????????????????
    file1 = r'D:/yidun/test.png'
    who = ''
    wx.ChatWith(who)  # ??????????????????
    wx.SendFiles(file1)
    # ?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    os.remove(r'D:/yidun/test.png')
    os.remove(r'D:/yidun/today.xlsx')
    os.remove(r'D:/yidun/image/img1.png')
    os.remove(r'D:/yidun/image/img1.png')
