#!/usr/bin/env python
# coding=utf-8

###############################################################################
# 腳本：API test with xlsx
# 功能：通過 xlsx 文件上的用例執行接口測試
# 作者：
#        ____ __   __  __ _____ ___  __  __
#       /_  /  _ \/ / / / ____/ __ \/ / / /
#        / / /_/ / /|/ /_/_  / / / / /|/ /
#     __/ / /-/ / / | /___/ / /_/ / / | /
#    /___/_/ /_/_/|_|/_____/\____/_/|_|/
#
# 日期：19 May 2016
# 版本：1.0
# 更新日誌:
#     19 May 2016
#         + 第一版
#
###############################################################################


import re
import smtplib
from email.mime.text import MIMEText
import json
import logging
import random
import os
import sys
# Excel 文件處理
try:
    import openpyxl
except ImportError as e:
    sys.exit('>>>>> 請安裝「openpyxl」：pip install openpyxl <<<<<\n')
#    os.system('pip install openpyxl')
#    import openpyxl

try:
    import requests
except ImportError as e:
    sys.exit('>>>>> 請安裝「requests」：pip install requests <<<<<\n')
#    os.system('pip install requests')
#    import requests

# 設置 requests 只顯示 WARNING 級別日誌（默認還會顯示 INOF 和 DEBUG 日誌）
logging.getLogger('requests').setLevel(logging.WARNING)

# 測試用例文件名
test_case_file = 'api_test_with_xlsx.xlsx'
# 第一個表格名稱
sheet1 = 'Basic Data'
# 第二個表格名稱
sheet2 = 'Test Case'

# 日誌文件保存路徑
log_file = os.path.join(os.getcwd(), 'log/api_test_with_xlsx.log')
if not os.path.exists('log'):
    os.makedirs('log')
    f = open(log_file, 'w')
    f.close()
log_format = '[%(asctime)s] [%(levelname)s] %(message)s'
# 使用「filename」參數後會自動增加 FileHandler
# filemode 文件打開模式：a 追加; w 寫入
logging.basicConfig(format=log_format, filename=log_file, filemode='w', level=logging.DEBUG)
console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
formatter = logging.Formatter(log_format)
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)



def send_mail(mail_host, mail_from, mail_pwd, mail_to, mail_sub, content):
    msg = MIMEText(content, 'html', _charset='utf-8')
    msg['From'] = mail_from
    msg['To'] = ';'.join(mail_to)
    msg['Subject'] = mail_sub
    try:
        s = smtplib.SMTP_SSL(mail_host, timeout=30)
        s.login(mail_from, mail_pwd)
        s.sendmail(mail_from, mail_to, msg.as_string())
        s.quit()
    except Exception as e:
        print('')
        logging.error('>>>>> 郵件發送失敗 <<<<<\n>> 異常：%s %s\n' % (type(e), e.args))
    else:
        logging.info('>>>>> 接口測試完成，郵件發送成功！ <<<<<')



# 作用：獲取 Excel 表中所有測試數據
# 參數：test_case_file 為測試數據所在 Excel 表文件路徑
#       sheet1 第一個表格名稱
#       sheet2 第二個表格名稱
def get_test_case(test_case_file, sheet1, sheet2):
    # 邮件正文，需要在多個函數中引用
    mail_content = ''
    # 以字典形式存放數據，便於後續操作
    res = {}    # 接口返回數據
    basic_data = {}     # Excel 中基礎數據

    # 獲取第一個表格數據
    # data_only=True 可避免讀取到單元格的公式
    #wb = openpyxl.load_workbook(test_case_file, data_only=True)
    wb = openpyxl.load_workbook(test_case_file)
    ws1 = wb.get_sheet_by_name(sheet1)

    # 把表格數據讀取到字典
    # 從第三行開始，讀取每一行數據
    for r in range(3, ws1.max_row + 1):
        r_key = ws1.cell(row=r, column=1).value
        r_value = ws1.cell(row=r, column=2).value
        # .strip 去除字符串前後空格（包括 Tab 和換行）
        # .replace(' ', '') 去除字符串所有空格
        try:
            basic_data[r_key.strip()] = r_value.strip()
        # 可能存在 int 類型值，無須處理
        except AttributeError as e:
            basic_data[r_key.strip()] = r_value

    # .split(',') 通過「,」劃分把 mail_to_* 轉為列表
    mail_to_all = basic_data['mail_to_all'].strip().replace(' ', '').split(',')
    mail_to_me = basic_data['mail_to_me'].strip().replace(' ', '').split(',')

    # 獲取第二個表格數據
    wb = openpyxl.load_workbook(test_case_file)
    ws2 = wb.get_sheet_by_name(sheet2)
    for r in range(3, ws2.max_row + 1):
        # 每一行為一條獨立測試用例，執行完才會執行下一條用例
        test_case = {}
        for c in range(1, ws2.max_column + 1):
            c_key = ws2.cell(row=1, column=c).value
            c_value = ws2.cell(row=r, column=c).value
            test_case[c_key] = c_value

        # is_active 等於 no 表示不執行該用例，直接執行下一次循環
        if test_case['is_active'] == 'yes':
            pass
        else:
            continue

        test_case['api_url'] = 'http://%s%s' % (test_case['api_host'], test_case['req_url'])

        # 存放接口執行結果
        res[test_case['api_id']] = {}

        # req_data 接口請求數據不為 None 時，把數據轉為字典
        if test_case['req_data']:
            try:
                # eval 將 excel 表裏的參數轉為正確的值
                #     eval 方法存在風險，鑑於此腳本不與外界交互，暫不考慮安全性
                test_case['req_data'] = eval(test_case['req_data'])
            except (NameError, KeyError, SyntaxError) as e:
                logging.error('API: %s >> 執行失敗 >>\n>> 異常：%s %s\n' % (test_case['api_title'], type(e), e.args))
                mail_content = '%sAPI: %s >> 執行失敗 >><br>>> URL: %s<br>>> 異常：%s %s<br><br>' % (mail_content, test_case['api_title'], test_case['api_url'], type(e), e.args)
                continue

            if not isinstance(test_case['req_data'], dict):
                logging.error('API: %s >> 執行失敗 >>\n>> 原因：「req_data」要求為字典類型 - %s' % (test_case['api_title'], test_case['req_data']))
                mail_content = '%sAPI: %s >> 執行失敗 >><br>>> URL: %s<br>>> 原因：「req_data」要求為字典類型 - %s<br><br>' % (mail_content, test_case['api_title'], test_case['api_url'], test_case['req_data'])
        else:
            test_case['req_data'] = ''

        # check_point 檢查點為 None 時該用例不再執行
        if not test_case['check_point']:
            logging.error('API: %s >> 執行失敗 >> 「check_point」不可為空' % (test_case['api_title'],))
            mail_content = '%sAPI: %s >> 執行失敗 >><br>%s<br>「check_point」不可為空<br><br>' % (mail_content, test_case['api_title'], test_case['api_url'])
            continue

        # 執行接口測試，把接口返回值保存在 res 字典中
        res[test_case['api_id']], mail_content = run_api(test_case['api_url'], test_case['req_method'], test_case['req_data'], test_case['api_title'], test_case['check_point'], basic_data['session_id'], mail_content)
        # run_api 函數裏會把登入接口的 session_id（如在 Excel 表中未設置）保存到接口返回值中
        if 'session_id' in res[test_case['api_id']]:
            basic_data['session_id'] = res[test_case['api_id']]['session_id']
        else:
            pass

    if res == {}:
        logging.error('未執行任何接口測試\n')
        mail_content = '未執行任何接口測試<br>'
    else:
        pass

    # if_mail 是否郵件通知測試結果
    #    0 不下發郵件；1 每次都下發郵件；2 僅接口出錯時下發郵件
    if basic_data['if_mail'] == 1:
        if mail_content != '':
            mail_content = '%s<br><b>如有問題，請致電 <font color="red">%s</font>（%s）。</b>' % (mail_content, basic_data['contact_phone'], basic_data['contact_name'])
            mail_to = mail_to_all
        else:
            # mail_content_random 接口正常時的隨機郵件正文
            # .split(';') 通過「;」劃分把 mail_content_random 轉為列表
            contents = basic_data['mail_content_random'].split(';')
            mail_content = random.choice(contents)
            basic_data['mail_sub'] = '%s〔正常〕' % (basic_data['mail_sub'],)
            mail_to = mail_to_me
        send_mail(basic_data['mail_host'], basic_data['mail_from'], basic_data['mail_pwd'], mail_to, basic_data['mail_sub'], mail_content)
    elif basic_data['if_mail'] == 2:
        if mail_content != '':
            mail_content = '%s<br><b>如有問題，請致電 <font color="red">%s</font>（%s）。</b>' % (mail_content, basic_data['contact_phone'], basic_data['contact_name'])
            mail_to = mail_to_all
            send_mail(basic_data['mail_host'], basic_data['mail_from'], basic_data['mail_pwd'], mail_to, basic_data['mail_sub'], mail_content)
        else:
            pass
    else:
        pass
    #print(mail_content)


def run_api(url, req_method, req_data, api_title, check_point, session_id, mail_content):
    headers = {
            'Content-Type':'application/x-www-form-urlencoded; charset=UTF-8',
            'X-Requested-With':'XMLHttpRequest',
            'Connection':'keep-alive',
            'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36'
            }
    # session_id 不為 None 時
    if session_id:
        headers['Cookie'] = 'session_id=%s' % (session_id,)

    try:
        if req_method == 'post':
            r = requests.post(url, data=req_data, headers=headers)
        elif req_method == 'get':
            r = requests.get(url, params=req_data, headers=headers) if req_data else requests.get(url, headers=headers)
        else:
            logging.error('API: %s >> 執行失敗 >>\n>> 原因：「req_method」參數不正確。\n' % (api_title,))
            mail_content = '%sAPI: %s >> 執行失敗 >><br>>> 原因：「req_method」參數不正確。<br><br>' % (mail_content, api_title)
            return {'msg': '執行失敗'}, mail_content

    # 後續優化：斷網時保存信息，下次執行判斷到信息再發送出來
    except requests.exceptions.RequestException as e:
        logging.error('API: %s >> 執行失敗 >>\n>> 異常：%s %s\n' % (api_title, type(e), e.args))
        mail_content = '%sAPI: %s >> 執行失敗 >><br>>> 異常：%s %s<br><br>' % (mail_content, api_title, type(e), e.args)
        return {'msg': '執行失敗'}, mail_content

    # 判斷接口返回結果是否為類 json 格式 { : }
    if re.match(r'^{[^:]*:.*}$', r.text):
        resp = json.loads(r.text)
        #print('返回結果：%s' % (resp,))
    else:
        resp = r.text

    try:
        # eval 將 excel 表裏的參數轉為正確的值
        #     eval 方法存在風險，鑑於此腳本不與外界交互，暫不考慮安全性
        is_check_point = eval(check_point)
    except (AttributeError, NameError, KeyError, SyntaxError, TypeError) as e:
        logging.error('API: %s >> 執行失敗 >>\n>> 異常：%s %s\n' % (api_title, type(e), e.args))
        mail_content = '%sAPI: %s >> 執行失敗 >><br>>> 異常：%s %s<br><br>' % (mail_content, api_title, type(e), e.args)
        return {'msg': '執行失敗'}, mail_content

    if is_check_point:
        logging.info('API: %s >> 執行成功' % (api_title,))
        # 如在 Excel 表中未設置 session_id 則把登入接口的 session_id 保留下來
        if not session_id and re.match(r'^.*/user/login$', url):
            resp['session_id'] = r.cookies.values()[0]
        else:
            pass
        return resp, mail_content
    else:
        logging.error('API: %s >> 執行失敗 >>\n>> Status Code: %d\n>> URL: %s\n>> Response: %s\n' % (api_title, r.status_code, url, resp))
        mail_content = '%sAPI: %s >> 執行失敗 >><br>>> Status Code: %d<br>>> URL: %s<br>>> Response: %s<br><br>' % (mail_content, api_title, r.status_code, url, resp)
        return {'msg': '執行失敗'}, mail_content



def main():
    get_test_case(test_case_file, sheet1, sheet2)


if __name__ == '__main__':
    main()

