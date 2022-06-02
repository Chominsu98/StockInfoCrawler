from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#검색창에서 엔터를 입력하기 위해 selenium에서 keys를 가지고옴
import time
#웹브라우저 원격조종 모듈
from bs4 import BeautifulSoup
from tkinter import *
from tkinter import messagebox
#알림창을 띄우기위해서
from tkinter import filedialog
#파일탐색기를 위해서
import openpyxl as xl
#엑셀에 접근하기 위한 모듈

# headers={'User-Agent' :'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36'}
#이것은 네이버 같은데는 agent 정보가 있으면 request 있을 때 뚫고 들어갈 수 있다.

all_data=[]
sub_data=[]
quarter_data=[]
def open_login(url_name):
    global browser
    browser=webdriver.Chrome('chromedriver')

    browser.implicitly_wait(3)

    browser.get(url_name)
    print("로그인 페이지에 접근합니다.")
    time.sleep(3)
    #밑에 계정이랑 패스워드 넣어줄 것
    USER=""
    PASS=""
    e=browser.find_element_by_id("login-user-id")
    e.clear()
    e.send_keys(USER)

    e=browser.find_element_by_xpath("/html/body/div[1]/div/div/div/ul/li[1]/div[2]/form/input[2]") #xml로 해당위치를 정확히 알려줌
    e.clear()
    e.send_keys(PASS)
    browser.find_element_by_xpath('/html/body/div[1]/div/div/div/ul/li[1]/div[2]/form/input[3]').click()#최선은 id나 name으로 찾는게 좋다 하지만 없다면 xml로..
    print("로그인 버튼을 클릭합니다.")

    time.sleep(3)
    url=browser.current_url
    browser.get(url)
    time.sleep(2)
    print("주식정보창에 진입성공!")

def get_info(class_name,arr,state="td"):
    html=browser.page_source
    soup=BeautifulSoup(html,'html.parser')
    table = soup.find('table', { 'class': class_name })

    data=[]

    for tr in table.find_all('tr'):
        if state=="th":
            ths=list(tr.find_all('th'))
            if ths!=[]:
                for th in ths:
                    data.append(th.text.strip())
            else:
                print("th 없다~")
                data.append("걍 없다.")
            quarter_data.append(data)
            data = []
            state="td"
            continue

        tds=list(tr.find_all('td'))
        if tds!=[]:
            for td in tds:
                data.append(td.text.strip())
        else:
            print("td 없다~")
            data.append("날짜")
        if arr=="sub":
            sub_data.append(data)
        elif arr=="quart":
            quarter_data.append(data)
        data = []

    print(quarter_data)
    #print(sub_data)

def integrate(arr1,arr2):
    for i in range(0,len(arr1)):
        all_data.append(arr1[i]+arr2[i])

def clickme():
    messagebox.showinfo("시작합니다","확인 버튼을 누른 이후로는 팝업창이 \n 나오며 작업이 진행되오니 기다리시기 바랍니다.")
    root.destroy()
def path():
    root.dirname=filedialog.askdirectory()
    return root.dirname
def open_gui():
    global root
    root=Tk()
    root.title("Searching your company")

    global str_
    str_=StringVar()
    textbox=Entry(root,width=100,bd=5,textvariable=str_)
    textbox.grid(column=0,row=0)

    action=Button(root,text="엑셀만들기 시작",command=clickme)
    action.grid(column=0,row=2)
    set_path=Button(root,text="엑셀저장경로 지정",command=path)
    set_path.grid(column=0,row=1)
    label=Label(root,text="[사용방법]\n검색창에 스크랩을 원하는 회사의 이름을 \n정확히 입력할것!")
    label.grid(column=0,row=3)
    root.mainloop()
    return str_.get()

def find_company(company_name):
    COMPANY=company_name
    search_engine=browser.find_element_by_id("companyInput")
    search_engine.clear()
    search_engine.send_keys(COMPANY)
    search_engine.send_keys(Keys.ENTER)
    time.sleep(2)

'''def save_xl(company_name,save_path):
    new_xl=xl.Workbook()
    working_x=new_xl.active
    working_x.title=company_name+"정보"

    for row in range(1,len(all_data)+1):
        for column in range(1,len(all_data[row-1])+1):
            working_x.cell(row=row,column=column).value=all_data[row-1][column-1]

    new_xl.save(save_path+"/"+company_name+".xlsx")
'''
Input_v=open_gui()
path=root.dirname
print(path)

open_login("https://bigfinance.co.kr/login" )
find_company(Input_v)
get_info('subject','sub')
get_info('quarter','quart','th')
integrate(sub_data, quarter_data)

# print(all_data)
save_xl(Input_v,root.dirname)

browser.quit()

