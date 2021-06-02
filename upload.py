
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time, sys, argparse


def login(user, passwd, overtime):
    # 打开网页
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    # option.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
    driver = webdriver.Chrome()
    driver.get('http://sw.zyzypc.com.cn/login')
    time.sleep(1)
    # 使用密码登录
    if(user and passwd):
        print('login......')
        driver.find_element_by_xpath('/html/body/div[1]/form/input[1]').send_keys(user)
        driver.find_element_by_xpath('/html/body/div[1]/form/input[2]').send_keys(passwd)
        driver.find_element_by_xpath('/html/body/div[1]/form/button[1]').click()
        time.sleep(2)
    else:   # 等待用户登录
        print('Please input your username and password.')
        if(overtime == 0): # 永久等待
            while(driver.current_url != 'http://sw.zyzypc.com.cn/'):
                time.sleep(0)
        else:
            time.sleep(overtime)
    return driver

def upload_yaocai(file, sheet, start=0, end=-1):
    ''' Upload yaocai specimens
    :param file: file_path
    :param sheet: excel_sheet
    :desciption: The first line of the sheet must be table's header.
                The header of the sheet must contains 采集号，重量，箱号
    :return:
    '''
    # 读取文件
    data = pd.read_excel(file, sheet, header=0)
    print('Read file {}, columns: {}'.format(file, data.columns))
    # 上传药材标本
    context.find_element_by_xpath('//*[@id="sidebar-collapse"]/ul/li[3]/a').click()
    time.sleep(1)
    if(end == -1):
        end = len(data)-1
    for i in range(start, end+1):
        # 点击添加
        while (1):
            if context.find_element_by_xpath('//*[@id="btn_add"]').is_displayed():
                context.find_element_by_xpath('//*[@id="btn_add"]').click()
                break
                # 填写信息
        time.sleep(1)
        context.find_element_by_xpath('//*[@id="tiaomahao"]').clear()
        context.find_element_by_xpath('//*[@id="tiaomahao"]').send_keys(str(data.iloc[i]['采集号']))
        context.find_element_by_xpath('//*[@id="caijihao"]').click()
        time.sleep(1)

        context.find_element_by_xpath('//*[@id="ycweight"]').clear()
        context.find_element_by_xpath('//*[@id="parid"]').clear()
        time.sleep(1)
        context.find_element_by_xpath('//*[@id="ycweight"]').send_keys(str(data.iloc[i]['重量']))
        context.find_element_by_xpath('//*[@id="parid"]').send_keys(int(data.iloc[i]['箱号']))
        time.sleep(1)
        context.find_element_by_xpath('//*[@id="btnedit"]').click()
        # 确认完成
        while (1):
            if EC.alert_is_present()(context):
                test_alert = context.switch_to.alert
                test_alert.accept()
                print(data.iloc[i]['采集号'], '完成')
                break


def upload_laye(file, sheet, start=0, end=-1):
    ''' Upload yaocai specimens
    :param file: file_path
    :param sheet: excel_sheet
    :desciption: The first line of the sheet must be table's header.
                The header of the sheet must contains 采集号，标本状态(花), 标本状态(果), 数量，箱号
    :return:
    '''
    # 读取文件
    data = pd.read_excel(file, sheet, header=0)
    print('Read file {}, columns: {}'.format(file, data.columns))
    # 上传
    context.find_element_by_xpath('//*[@id="sidebar-collapse"]/ul/li[2]/a').click()
    time.sleep(1)
    if (end == -1):
        end = len(data) - 1
    for i in range(start, end + 1):
        # 点击添加
        while (1):
            if context.find_element_by_xpath('//*[@id="btn_add"]').is_displayed():
                context.find_element_by_xpath('//*[@id="btn_add"]').click()
                break
                # 填写信息
        time.sleep(1)
        context.find_element_by_xpath('//*[@id="tiaomahao"]').clear()
        context.find_element_by_xpath('//*[@id="tiaomahao"]').send_keys(str(data.iloc[i]['采集号']))
        context.find_element_by_xpath('//*[@id="caijihao"]').click()
        time.sleep(1)
        if data.iloc[i]['标本状态(花)'] == '√':
            context.find_element_by_xpath('//*[@id="hua_1"]').click()
        else:
            context.find_element_by_xpath('//*[@id="hua_0"]').click()
        if data.iloc[i]['标本状态(果)'] == '√':
            context.find_element_by_xpath('//*[@id="guo_1"]').click()
        else:
            context.find_element_by_xpath('//*[@id="guo_0"]').click()
        context.find_element_by_xpath('//*[@id="number"]').clear()
        context.find_element_by_xpath('//*[@id="parid"]').clear()
        time.sleep(1)
        context.find_element_by_xpath('//*[@id="number"]').send_keys(int(data.iloc[i]['数量']))
        context.find_element_by_xpath('//*[@id="parid"]').send_keys(int(data.iloc[i]['箱号']))
        time.sleep(1)
        context.find_element_by_xpath('//*[@id="btnedit"]').click()
        # 确认完成
        while (1):
            if EC.alert_is_present()(context):
                test_alert = context.switch_to.alert
                test_alert.accept()
                print(data.iloc[i]['采集号'], '完成')
                break

def get_params():
    parser = argparse.ArgumentParser(
        # usage='',
        description='Upload specimens data to http://sw.zyzypc.com.cn/ycypsubedit.aspx')
    parser.add_argument('-u', '--user', type=str, help='username')
    parser.add_argument('-p', '--password', type=str, help='password')
    parser.add_argument('-t', '--specimen_type', required=True, type=int, choices=[2, 3],
        help='''Type of specimen. Select 2 to upload laye specimen, 3 to upload yaocai specimen.''')
    parser.add_argument('-f', '--specimen_file', type=str, required=True,
        help='File path of specimens, just support ".xls" format.')
    parser.add_argument('--sheet', default=0, type=int, help='Sheet of file')
    parser.add_argument('--start', default=0, type=int, help='Start row')
    parser.add_argument('--end', default=-1, type=int, help='End row')
    parser.add_argument('-o', '--overtime', default=0, type=int, help='Login overtime')

    args = parser.parse_args()
    return args

if __name__ == '__main__':
    params = get_params()
    print(params.user, params.password, params.specimen_type, params.specimen_file, params.overtime)
    # 登录
    context = login(params.user, params.password, params.overtime)
    time.sleep(1)
    try:
        context.find_element_by_xpath('//*[@id="sidebar-collapse"]')
    except:  # 没有登录成功
        print('login overtime.')
        context.close()
        sys.exit(-1)
    print('login ok.')
    if(params.specimen_type == 2):
        upload_laye(file=params.specimen_file, sheet=params.sheet, start=params.start, end=params.end)
    elif(params.specimen_type == 3):
        upload_yaocai(file=params.specimen_file, sheet=params.sheet, start=params.start, end=params.end)
    else:
        print("specimen_type error. Select 2 to upload laye specimen, 3 to upload yaocai specimen .")
    context.close()


