import time
import json
import csv

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

listMonHoc = []

#Mở Edge và vào trang qldt
driver = webdriver.Edge(executable_path='C:/Users/thang/OneDrive/Desktop/Kì 20211/Project 1/auto_web/msedgedriver.exe')
driver.get("https://qldt.hust.edu.vn/")
driver.maximize_window()

#Xử lý lịch học
def xuLyLichHoc(s):

    #Xử lý những môn tự học (VD: project1...)
    if(len(s) < 10):
        return 0, 0, 0

    arr = s.split(',', 3)
    thoiGian = arr[1]
    tuanHoc = arr[3]
    return arr[0], thoiGian, tuanHoc

#Xử lý thời gian
def xuLyThoiGian(s):
    if(s == 0):
         return 0
    if (s[3] == ' '): 
        if (s[4] == '1'):
            return int(s[4])*10 + int(s[5])
        else:
            return int(s[4])
    else:
        return int(s[6])*3/4 + 7 

#Hỗ trợ sắp xếp lịch học
def custom_sort(elem):
    return elem['Độ ưu tiên']

#Tổng hợp lịch học
def tongHop(lichHoc, tenMonHoc):
    thu, thoiGian, tuanHoc = xuLyLichHoc(lichHoc)
    tg = xuLyThoiGian(thoiGian)
    if(thu == 0):
        doUuTien = 0
    else:
        if (thu[0] == 'S'):
            doUuTien = int(thu[-1])*100 + tg
        else:
            doUuTien = int(thu[-1])*100 + tg + 5
    thongTinMonHoc = {
        'Thứ': thu,
        'Thời gian': thoiGian,
        'Tuần học': tuanHoc,
        'Môn học': tenMonHoc,
        'Độ ưu tiên': doUuTien    
    }
    listMonHoc.append(thongTinMonHoc)

def process(us):
    #Đăng Nhập vào trang qldt
    time.sleep(1)
    driver.find_element_by_css_selector(".btn-login-main-style").click()
    time.sleep(1)
    driver.find_element_by_css_selector(".orLoginByAccount b a").click()

    #Nhập tên đăng nhập
    try:
        time.sleep(1)
        email = driver.find_element_by_css_selector(".texBoxLogin:first-child")
        email.clear()
        email.send_keys(us['username'])

        #Nhập password
        time.sleep(1)
        password = driver.find_element_by_name("password")
        password.clear()
        password.send_keys(us['password'])

        #Nhấn vào nút đăng nhập
        time.sleep(1)
        driver.find_element_by_css_selector(".gwt-HTML.buttonLoading-status").click()

        #Nhấn vào nút lịch học
        time.sleep(5)
        hoTen = driver.find_element_by_css_selector(".GHB .gwt-HTML").text
        time.sleep(1)
        driver.find_element_by_css_selector(".navbar-nav-header-panel > li:nth-child(4)").click()

        #Chọn kỳ học
        try:
            time.sleep(1)
            kyHoc = Select(driver.find_element_by_css_selector(".top-function-panel-in-data-table.top-function-panel .form-control"))
            kyHoc.select_by_value(us['kyHoc'])
            time.sleep(2)

            #xử lý thời khóa biểu
            soMonHoc = len(driver.find_elements_by_xpath("/html/body/div[3]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div/table/tbody[1]/tr"))
            tenMonHoc = ""
            lichHoc = ""

            cotLichHoc = 4
            if (us['kyHoc'] == '20211'):
                cotLichHoc = 6
            for i in range(1, soMonHoc + 1):
                tenMonHoc = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div/table/tbody[1]/tr[" + str(i) + "]/td[1]/div/p").text
                lichHoc = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div/table/tbody[1]/tr[" + str(i) + "]/td[" + str(cotLichHoc) + "]/div").text

                #Xử lý những môn có 2 buổi / tuần
                tmp = lichHoc.find('Tiết',45)
                if(tmp > 0):                            
                    tmp1 = lichHoc.find(lichHoc[0:4],1)
                    lichHoc1 = lichHoc[0:(tmp1 - 1)]
                    lichHoc2 = lichHoc[tmp1:]
                    tongHop(lichHoc1, tenMonHoc)
                    tongHop(lichHoc2, tenMonHoc)
                else:
                    tongHop(lichHoc, tenMonHoc)


            listMonHoc.sort(key=custom_sort)
            with open('TKB.csv', mode='w', encoding='utf-8') as file:
                writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(['Kỳ: ' + us['kyHoc']])
                writer.writerow(['Họ và tên: ' + hoTen])
                writer.writerow(['Thứ', 'Thời gian', 'Tuần học', 'Môn học'])
                for j in range (0, soMonHoc):
                    writer.writerow([listMonHoc[j]['Thứ'],listMonHoc[j]['Thời gian'],listMonHoc[j]['Tuần học'],listMonHoc[j]['Môn học']])
        except Exception:
            print('Ky hoc khong duoc tim thay!')
    except Exception:
        print('Ten dang nhap hoac mat khau khong chinh xac!')   

def main():
    #Đọc dữ liệu từ file text
    file = open('info.txt', 'r')
    data = file.read()
    file.close()
    us = json.loads(data)
    #Xử lý
    process(us)

if __name__ == "__main__":
    main()
