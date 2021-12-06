from importlib.abc import ResourceReader
import gspread
import datetime
import time
import os
import re
from gspread.exceptions import APIError, SpreadsheetNotFound
from gspread.models import Worksheet
from pyasn1_modules.rfc2459 import Name
import requests
from datetime import datetime, timedelta
from datetime import date

### def list

# 입항일, 요청건수, 신고서 작성, 수리/대기 항목만 추출 후, 같은 입항일 기준으로 나머지 항목 합산 
def extractData(sample_list):
    # 변수 설정
    date = ''
    require = ''
    register = ''
    repair = ''
    extract_list = []

    # 전체 데이터에서 필요한 항목 설정
    for list_line in sample_list:
        date = list_line[2]
        require = list_line[6]
        register = list_line[7]
        repair = list_line[13]
        extract_list.append([date, require, register, repair])

    # 입항일, 요청건수, 신고서 작성, 수리/대기 제목 제거
    extract_list = extract_list[1:]

    # 마지막 탈출 조건
    extract_list.append([1111,1111,1111,1111])

    # 변수 설정
    base_date = int(extract_list[0][0])
    sum_require = 0
    sum_register = 0
    sum_repair = 0
    su_list = []
    
    for i in extract_list:       
        
        #YYYYMMDD
        date = int(i[0])

        # 같은 날짜일 경우 요청건수, 신고서 작성, 수리/대기 계속 add
        if date == base_date:
            sum_require += int(i[1])
            sum_register += int(i[2])
            sum_repair += int(i[3])

        # 날짜가 달라질 경우
        else:
            su_list.append([base_date,sum_require,sum_register,sum_repair])
            base_date = int(i[0])
            sum_require = int(i[1])
            sum_register = int(i[2])
            sum_repair = int(i[3])
        
    return su_list

# 월별 총합 생성하는 함수
def sum_month_data(su_list):

    # 변수 설정
    base_month_date = int(str(su_list[0][0])[0:6])
    sum_month_require = 0
    sum_month_register = 0
    sum_month_repair = 0
    sum_month_list = []
    
    for i in su_list:       
        
        # YYYYMM
        date_month = int(str(i[0])[0:6])

        # 같은 날짜일 경우 요청건수, 신고서 작성, 수리/대기 계속 add
        if date_month == base_month_date:
            sum_month_require += int(i[1])
            sum_month_register += int(i[2])
            sum_month_repair += int(i[3])

        # 날짜가 달라질 경우
        else:
            sum_month_list.append([base_month_date,sum_month_require,sum_month_register,sum_month_repair])
            base_month_date = int(str(i[0])[0:6])
            sum_month_require = int(i[1])
            sum_month_register = int(i[2])
            sum_month_repair = int(i[3])
        
    return sum_month_list

# YYYYTotal 시트 생성 후 월별 총합 기입
def div_year(list):

    # YYYY
    default_year = int(str(list[0][0])[0:4])
    sum_year = []

    for i in list:
        if int(default_year) == 1111:
            return
        # 처음 시작은 YYYY
        d_year = int(str(i[0])[0:4])

        # 년도가 같은 경우
        if default_year == d_year:
            sum_year.append(i)

            try:
                #sec_sheet("통관실적").worksheet(str(i[0])[0:4] + 'Total')
                new_sheet = sec_sheet("통관실적").worksheet(str(i[0])[0:4] + 'Total')
            except: #gspread.exceptions.WorksheetNotFound:
                sec_sheet("통관실적").add_worksheet(title=str(i[0])[0:4] + "Total", rows="100", cols="20")
                new_sheet = sec_sheet("통관실적").worksheet(str(i[0])[0:4] + 'Total')

        # 년도가 다를 경우
        else:
            new_sheet.batch_update([{
                'range': 'A2:D32',
                'values': sum_year
            }, {
                'range': 'A1',
                'values': ne_list_name()[0],
            }, {
                'range': 'B1',
                'values': ne_list_name()[1],
            }, {
                'range': 'C1',
                'values': ne_list_name()[2],
            }, {
                'range': 'D1',
                'values': ne_list_name()[3],  
            }, {
                'range' : 'E1',
                'values' : [['통관율']],
            }, {
                'range': 'A2:A13',
                'values': [['1월'],['2월'],['3월'],['4월'],['5월'],['6월'],['7월'],['8월'],['9월'],['10월'],['11월'],['12월']]
            }])
            new_sheet.format("B2:D", {"numberFormat" : {'type' : 'number', 'pattern' : '#,###,###,###'}})
            sum_year = [i]
            default_date = int(str(i[0])[0:4])

# 월별로 구분해서 시트 생성 후 월별 항목 기입하는 함수
def div_date(extract):
    
    #YYYYMM
    default_date = int(str(extract[0][0])[0:6])
    sum_ex = []

    for i in extract:
        if int(default_date) == 111111:
            return
        # 처음 시작은 YYYYMM
        d_date = int(str(i[0])[0:6])
        
        # 월이 같을 경우
        if default_date == d_date:
            sum_ex.append(i)

            try:
                sec_sheet("통관실적").worksheet(str(i[0])[0:6])
            except gspread.exceptions.WorksheetNotFound:
                sec_sheet("통관실적").add_worksheet(title=str(i[0])[0:6], rows="32", cols="5")
            else:
                new_sheet = sec_sheet("통관실적").worksheet(str(i[0])[0:6])

        # 월이 다를 경우
        else:
            new_sheet.batch_update([{
                'range': 'A2:D32',
                'values': sum_ex
            }, {
                'range': 'A1',
                'values': ne_list_name()[0],
            }, {
                'range': 'B1',
                'values': ne_list_name()[1],
            }, {
                'range': 'C1',
                'values': ne_list_name()[2],
            }, {
                'range': 'D1',
                'values': ne_list_name()[3],  
            }, {
                'range' : 'E1',
                'values' : [["통관율"]]
            }])
            new_sheet.format("B2:D", {"numberFormat" : {'type' : 'number', 'pattern' : '#,###,###,###'}})
            sum_ex = [i]
            default_date = int(str(i[0])[0:6])

# 시트명이 한글인 시트 삭제하는 함수
def del_kor_ws():
    # 모든 워크시트 변수 지정
    all_ws = sec_sheet("통관실적").worksheets()

    # 한글로 된 시트만 따로 저장할 리스트 생성
    kor_ws = []

    # 영어이거나 숫자 또는 영어, 숫자 혼합 시트는 남기고 한글로 된 시트만 골라서 리스트에 저장
    for i in all_ws:
        if i.title.encode().isalnum():
            None
        else:
            kor_ws.append(i)

    # 한글로 된 시트 삭제
    for i in kor_ws:
        delete_worksheet = sec_sheet("통관실적").worksheet(i.title)
        sec_sheet("통관실적").del_worksheet(delete_worksheet)

# 오래된 날짜 순으로 정렬하는 함수
def rearrange():
    # 모든 워크시트 변수 지정
    all_ws = sec_sheet("통관실적").worksheets()

    # 숫자로 된 워크시트가 아닌 시트를 담을 리스트 생성
    not_num_ws = []

    # 숫자 또는 숫자와 영어 혼합이 아닌 제목의 워크시트일 경우 따로 만든 리스트에 저장
    for i in all_ws:
        if i.title.isdigit() == False:
            not_num_ws.append(i)
        else:
            continue

    # 숫자가 아닌 워크시트를 모든 워크시트에서 삭제
    for i in not_num_ws:
        all_ws.remove(i)    

    # 오래된 날짜 순으로 새로 정렬할 리스트 생성
    re_all_ws = []
    
    count = 0
    
    count_all_ws = len(all_ws)

    # 오래된 날짜 순으로 새로 정렬
    while len(all_ws) > 0:
        Min = 100000000000000
        Min_sheet = None
        for i in all_ws:
            if Min > int(i.title):
                Min = int(i.title)
                Min_sheet = i
            else:
                continue

        # 만든 리스트에 삽입
        re_all_ws.append(Min_sheet)

        # False 조건
        all_ws.remove(Min_sheet)

        count += 1

        # 과도한
        if count == count_all_ws * 3:
            print("rearrange 함수 오류 발생")
            break
    
    # 날짜 순으로 정렬된 리스트로 재정렬
    sec_sheet("통관실적").reorder_worksheets(re_all_ws)

# 새로 만들 스프레드 시트 불러오는 함수
def sec_sheet(sheet_name):
    sh = path_oauth().open(sheet_name)
    return sh

# gspread oauth 함수
def path_oauth():
    # gc = gspread.oauth(auth_file_path=path)
    gc = gspread.oauth(
        #절대 경로와 상대 경로 중요!!!
        credentials_filename='./credentials.json',
        authorized_user_filename='./authorized_user.json'
    )
    return gc

# 원래 자료가 있는 스프레드 시트 불러오는 함수
def ori_sheet():
    # 스프레드 시트 열기
    sh = path_oauth().open("통합업무")
    return sh

# 원래 자료가 있는 스프레드시트에서 필요한 항목의 워크시트 지정하는 함수
def all_info_sheet():
    # 모든 정보가 있는 시트 지정
    worksheet = ori_sheet().worksheet("수입신고처리현황")
    return worksheet

# 필요한 항목의 이름 가져오는 함수
def ne_list_name():
    # ['입항일', '요청건수', '신고서 작성', '수리/대기']
    list_name = all_info_sheet().batch_get(['C1', 'G1', 'H1', 'N1'])
    return list_name

def col_E_pert(sheet_name):
    
    worksheet = sec_sheet("통관실적").worksheet(sheet_name)

    for i in range(2, 33):
        worksheet.update("E{}".format(i), "=TRUNC(IFERROR(D{}/C{}, 0), 4)".format(i, i), raw=False)

    worksheet.format("E2:E", {"numberFormat": {"type" : "percent"}})

def main():
    # 경로 지정
    path = os.path.join(os.getcwd(), 'crendentials', 'joe')

    gc = path_oauth()

    sh = ori_sheet()

    worksheet = all_info_sheet()

    tmp_list = ne_list_name()

    # 모든 정보 가져오기
    sample_list = worksheet.get_all_values()

    # 1) 입항일, 요청건수, 신고서 작성, 수리/대기 항목만 추출
    # ex_list = extractData(sample_list)

    # 2) 같은 입항일 기준 나머지 항목 합산
    # su_list = sum_data(ex_list)

    # 1) + 2)
    su_list = extractData(sample_list)

    # 탈출 조건
    su_list.append([11111111,1111,1111,1111])

    # 월별 총합
    sum_month_list = sum_month_data(su_list)

    # 마지막 요소인 [11111111,1111,1111,1111] 제거
    su_list.pop()

    # ['입항일', '요청건수', '신고서 작성', '수리/대기']
    #list_name = all_info_sheet().batch_get(['C1', 'G1', 'H1', 'N1'])

    # 새로운 스프레드 시트 생성 (존재 여부 조건 삽입)
    try:   
        gc.open("통관실적")
        sh2 = gc.open("통관실적")
    except SpreadsheetNotFound:
        gc.create('통관실적')
        sh2 = gc.open("통관실적")

    sum_month_list.append([1111,1111,1111,1111])

    #print(sum_month_list)

    div_year(sum_month_list)

    worksheet_Total2021 = sec_sheet("통관실적").worksheet("2021Total")

    from_Sep = worksheet_Total2021.get('B2:D5')

    worksheet_Total2021.update('B10:D13', from_Sep, raw=False)

    worksheet_Total2021.batch_clear(["B2:D5"])

    # 탈출 조건
    su_list.append([11111111, 1111, 1111, 1111])

    # quota exceed일 경우 time sleep
    for i in range(0, 5):
        try:
            div_date(su_list)
            break
        except Exception as e:

            except_name = str(e)
            matchOB = re.match(except_name, 'Quota Exceed')
            
            if matchOB != -1:
                time.sleep(120)
            else:
                print(e)
                break

    # 시트명이 한글인 시트 삭제
    del_kor_ws()

    # 오래된 날짜 순으로 정렬
    rearrange()

if __name__ == '__main__' :
    main()