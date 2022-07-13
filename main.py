import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from bs4 import BeautifulSoup
import time
import datetime
import sys
import os
import requests
from PyQt5.QtWidgets import QWidget, QApplication,QTreeView,QFileSystemModel,QVBoxLayout,QPushButton,QInputDialog,QLineEdit,QMainWindow,QMessageBox
from PyQt5.QtCore import QCoreApplication
from selenium.webdriver import ActionChains
from datetime import datetime,date,timedelta
import time
import requests
from bs4 import BeautifulSoup as bs
import sys,os,shutil
from PyQt5.QtWidgets import QWidget, QApplication,QTreeView,QFileSystemModel,QVBoxLayout,QPushButton,QInputDialog,QLineEdit,QMainWindow
from window import Ui_MainWindow
import chromedriver_autoinstaller




class Example(QMainWindow,Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.path="C:"
        self.index=None
        self.setupUi(self)
        self.setSlot()
        self.show()
        self.progressBar.setValue(0)
        QApplication.processEvents()

    def start(self):
        def createFolder(directory):
            try:
                if not os.path.exists(directory):
                    os.makedirs(directory)
            except OSError:
                print('Error: Creating directory. ' + directory)

        createFolder('C:/auto_search/')

        chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
        driver_path = f'./{chrome_ver}/chromedriver.exe'
        if os.path.exists(driver_path):
            print(f"chrom driver is insatlled: {driver_path}")
        else:
            print(f"install the chrome driver(ver: {chrome_ver})")
            chromedriver_autoinstaller.install(True)

        # print(driver_path)
        options = webdriver.ChromeOptions()
        # options.add_argument('headless')
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        keyword = self.lineEdit_1.text()
        print("키워드는:",keyword)
        browser = webdriver.Chrome(driver_path, options=options)

        url = ['https://www.114.co.kr/search/result/db?query={}'.format(keyword),
               'https://www.114.co.kr/search/result/direct?query={}'.format(keyword),
               'https://www.114.co.kr/search/result/collect?query={}'.format(keyword),
               'https://www.114.co.kr/search/result/public?query={}'.format(keyword)]
        # url = ['https://www.114.co.kr/search/result/direct?query={}'.format(keyword),'https://www.114.co.kr/search/result/collect?query={}'.format(keyword),'https://www.114.co.kr/search/result/public?query={}'.format(keyword)]
        # url = ['https://www.114.co.kr/search/result/collect?query={}'.format(keyword),'https://www.114.co.kr/search/result/public?query={}'.format(keyword)]
        tab = ""

        # 엑셀파일열기
        wb = openpyxl.Workbook()
        ws = wb.active
        first_row = ['키워드', '분류', '상호명', "주소", "연락처"]
        ws.append(first_row)

        for index, eachUrl in enumerate(url):
            print(eachUrl)
            browser.get(eachUrl)
            if eachUrl.find("/db") >= 0:
                index = 0
            elif eachUrl.find("/direct") >= 0:
                index = 1
            elif eachUrl.find("/collect") >= 0:
                index = 2
            elif eachUrl.find("/public") >= 0:
                index = 3

            if index == 0:
                tab = "통신사"
            elif index == 1:
                tab = "본인등록"
            elif index == 2:
                tab = "이용자제보"
            elif index == 3:
                tab = "공공데이터"
            print("인덱스는:", index)
            browser.implicitly_wait(5)
            time.sleep(3)

            # 마지막 페이지를 눌러서 최종 페이지로 이동한다.
            # contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.user > div.regi-info > div > div > a.page.last
            try:
                if index == 0:
                    btnLast = browser.find_element(By.CSS_SELECTOR,
                                                   '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.agency > div.regi-info > div > div > a.page.last')
                elif index == 1:
                    btnLast = browser.find_element(By.CSS_SELECTOR,
                                                   '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.me > div.regi-info > div > div > a.page.last')
                elif index == 2:
                    btnLast = browser.find_element(By.CSS_SELECTOR,
                                                   '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.user > div.regi-info > div > div > a.page.last')
                elif index == 3:
                    btnLast = browser.find_element(By.CSS_SELECTOR,
                                                   '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.public > div.regi-info > div > div > a.page.last')
                btnLast.click()
                time.sleep(1)

                # 마지막 페이지로 이동을 해서 최종 페이지 번호를 읽어온다.
                if index == 0:
                    listPages = browser.find_element(By.CSS_SELECTOR,
                                                     '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.agency > div.regi-info > div > div')
                elif index == 1:
                    listPages = browser.find_element(By.CSS_SELECTOR,
                                                     '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.me > div.regi-info > div > div')
                elif index == 2:
                    listPages = browser.find_element(By.CSS_SELECTOR,
                                                     '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.user > div.regi-info > div > div')
                elif index == 3:
                    listPages = browser.find_element(By.CSS_SELECTOR,
                                                     '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.public > div.regi-info > div > div')
                lastPage = int(listPages.find_elements(By.TAG_NAME, 'a')[-3].text)
                print("마지막페이지는:", lastPage)

                # 첫번째 페이지를 눌러서 처음으로 돌아간다.
                if index == 0:
                    btnFirst = browser.find_element(By.CSS_SELECTOR,
                                                    '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.agency > div.regi-info > div > div > a.page.first')
                elif index == 1:
                    btnFirst = browser.find_element(By.CSS_SELECTOR,
                                                    '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.me > div.regi-info > div > div > a.page.first')
                elif index == 2:
                    btnFirst = browser.find_element(By.CSS_SELECTOR,
                                                    '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.user > div.regi-info > div > div > a.page.first')
                elif index == 3:
                    btnFirst = browser.find_element(By.CSS_SELECTOR,
                                                    '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.public > div.regi-info > div > div > a.page.first')

                browser.execute_script("arguments[0].click();", btnFirst)
                time.sleep(1)
            except:
                lastPage = 1

            # browser.find_element(By.CSS_SELECTOR,'#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.agency > div.regi-info > div > div > a:nth-child(6)').click()

            for page in range(1, lastPage + 1):
                if lastPage == 1:
                    print("페이지 1개뿐임")
                else:
                    print("항목:", tab, "페이지번호:", page, "/", lastPage)
                    # 페이지번호가 10으로 나눴을 때 1이면 넥스트 페이지 버튼을 누른다.
                    if page > 9 and page % 10 == 1:
                        if index == 0:
                            nextPage = browser.find_element(By.CSS_SELECTOR,
                                                            '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.agency > div.regi-info > div > div > a.page.next').click()
                        elif index == 1:
                            nextPage = browser.find_element(By.CSS_SELECTOR,
                                                            '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.me > div.regi-info > div > div > a.page.next').click()
                        elif index == 2:
                            nextPage = browser.find_element(By.CSS_SELECTOR,
                                                            '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.user > div.regi-info > div > div > a.page.next').click()
                        elif index == 3:
                            nextPage = browser.find_element(By.CSS_SELECTOR,
                                                            '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.public > div.regi-info > div > div > a.page.next').click()

                        time.sleep(1)
                    # 10으로 나눈 나머지로 누를 태그의 순서를 본다.
                    convertedPage = page % 10
                    # 만약 나머지가0이면 숫자 10으로 치환한다.
                    if convertedPage == 0:
                        convertedPage = 10
                    if index == 0:
                        browser.find_element(By.CSS_SELECTOR,
                                             '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.agency > div.regi-info > div > div > a:nth-child({})'.format(
                                                 convertedPage + 2)).click()
                    elif index == 1:
                        browser.find_element(By.CSS_SELECTOR,
                                             '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.me > div.regi-info > div > div > a:nth-child({})'.format(
                                                 convertedPage + 2)).click()
                    elif index == 2:
                        browser.find_element(By.CSS_SELECTOR,
                                             '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.user > div.regi-info > div > div > a:nth-child({})'.format(
                                                 convertedPage + 2)).click()
                    elif index == 3:
                        browser.find_element(By.CSS_SELECTOR,
                                             '#contain > div > div.srch-cont > div:nth-child(1) > div.regi-info-wp.public > div.regi-info > div > div > a:nth-child({})'.format(
                                                 convertedPage + 2)).click()
                    time.sleep(1)

                # 페이지 내의 정보를 BeautifulSoup으르 저장한다.
                soup = BeautifulSoup(browser.page_source, 'lxml')
                eachItems = soup.find_all('div', attrs={'class': 'regi-item'})

                # 각 열마다 상호,주소,전화번호 정보를 가져온다.
                for eachItem in eachItems:
                    try:
                        name = eachItem.find('strong', attrs={'class': 'tit'}).get_text()
                    except:
                        name = "없음"
                    try:
                        address = eachItem.find('li', attrs={'class': 'add'}).get_text()
                        addressPosition = address.find("지번")
                        address = address[addressPosition + 2:].strip()
                    except:
                        address = "없음"
                    try:
                        phoneNumber = eachItem.find('li', attrs={'class': 'tel'}).get_text()
                    except:
                        phoneNumber = "없음"
                    print("상호명:", name, "주소:", address, "연락처:", phoneNumber)
                    data = [keyword, tab, name, address, phoneNumber]
                    ws.append(data)
                print("----------------------------------------------------")
                time.sleep(1)
            wb.save('C:/auto_search/result.xlsx')
            self.num = int((index+1)/4*100)
            self.progressBar.setValue(self.num)
            QApplication.processEvents()
        print("작업완료")

        QMessageBox.information(self, "완료창", "작업이 완료 되었습니다.")
        QCoreApplication.instance().quit()
        print("작업완료")



    def setSlot(self):
        pass

    def setIndex(self,index):
        pass
    def quit(self):
        QCoreApplication.instance().quit()

app=QApplication([])
ex=Example()
sys.exit(app.exec_())



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())