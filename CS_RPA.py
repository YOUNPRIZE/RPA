from typing import final
import gspread
import datetime
import time
import os
import re
from gspread.models import Worksheet

# gspread oauth 함수
def path_oauth():
    # gc = gspread.oauth(auth_file_path=path)
    gc = gspread.oauth(
        # 절대 경로와 상대 경로 중요!!!
        credentials_filename='./credentials.json',
        authorized_user_filename='./authorized_user.json'
    )
    return gc


def main():
    path_oauth()
    
    # 스프레드 시트 열기
    # 데이터를 불러올 원래 시트
    sh = path_oauth().open("CS미결 개인부호오류")

    # 불러온 데이터를 가공해서 저장할 새로운 시트
    sh2 = path_oauth().open("CS실적")

    # 전체 시트 불러오는 과정
    all_sheet = sh.worksheets()

    # quota exceed 방지
    time.sleep(120)
    
    # YYYYMMDD의 형태로 된 시트만 담을 리스트 생성
    digit_sheet = []

    # YYYYMMDD의 형태로 된 시트만 추출
    for i in all_sheet:
        if i.title.isdigit():
            digit_sheet.append(i)
        else:
            None

    # 날짜가 오래된 순으로 정렬
    digit_sheet.reverse()

    # 담당자 데이터가 존재하는 시트가 20211008 부터이므로 그 전 날짜의 시트들은 리스트에 제거
    # 담당자 데이터가 계속 존재하므로 주석 처리
    '''
    while digit_sheet[0].title != "20211008":
        for i in digit_sheet:
            if int(i.title) < 20211008:
                digit_sheet.remove(i)
    '''

    # 해당 경우가 더 이상 존재하지 않으므로 주석 처리
    '''
    unnec_sheet = []

    # 20211008 이후에 담당자 열은 존재하지만 아무 이름도 기입되어 있지 않은 시트 제거
    for i in digit_sheet:
        if sh.worksheet("{0}".format(i.title)).batch_get(["O4:O"]) == [[]]:
            unnec_sheet.append(i)
        else:
            None

    time.sleep(120)

    for i in unnec_sheet:
        digit_sheet.remove(i)

    # 최종적으로 담당자의 이름이 있는 시트만 남음
    '''

    # YYYYMM
    default_month = digit_sheet[0].title[0:6]

    month_list = []

    # 탈출 조건
    digit_sheet.append(Worksheet)

    # 담당자의 총 건수 값을 해당 월 시트를 생성해서 기입하는 함수
    for i in digit_sheet:
        # 더 이상 데이터가 없을 경우 (위에서 추가한 탈출 조건일 경우)
        if i == gspread.models.Worksheet:
            # 총 건수 추출하는 함수
            extract = counting(month_list)
            # 배열 재배치
            re_array = change_array(extract)
            # 새로운 시트에 가공한 데이터 저장
            new_sheet.batch_update([{
                'range': 'B3',
                'values': re_array,
            }, {
                'range': 'B2',
                'values': [['담당자']],
            }, {
                'range': 'C2',
                'values': [['해당 월 총 건수']]
            }], raw=False)
            new_sheet.format("C3:C", {"numberFormat":{'type':'number','pattern':'#,###,###,###'}})
            break
        
        # YYYYMM이 같을 경우
        elif default_month == i.title[0:6]:
            
            month_list.append(i)

            # 예외 처리
            try:
                new_sheet = sh2.worksheet("{0}년{1}월".format(i.title[0:4], i.title[4:6]))
            except:
                sh2.add_worksheet("{0}년{1}월".format(i.title[0:4], i.title[4:6]), rows="31", cols="10")
                new_sheet = sh2.worksheet("{0}년{1}월".format(i.title[0:4], i.title[4:6]))

        # YYYYMM이 달라질 경우
        else:
            # 총 건수 추출하는 함수
            extract = counting(month_list)
            # 배열 재배치
            re_array = change_array(extract)
            # 새로운 시트에 가공한 데이터 저장
            new_sheet.batch_update([{
                'range': 'B3',
                'values': re_array,
            }, {
                'range': 'B2',
                'values': [['담당자']],
            }, {
                'range': 'C2',
                'values': [['해당 월 총 건수']]
            }], raw=False)
            new_sheet.format("C3:C", {"numberFormat":{'type':'number','pattern':'#,###,###,###'}})
            month_list = []
            month_list.append(i)
            default_month = i.title[0:6]

# 총 건수 추출하는 함수
def counting(any_list):
    
    # 담당자의 이름만 담을 리스트 생성
    name_list = []

    time.sleep(120)

    # 담당자의 이름을 제외하고 '' 및 '담당자' 제거
    for i in any_list:
        sheet_value = path_oauth().open("CS미결 개인부호오류").worksheet("{0}".format(i.title)).get_all_values()
        for i in sheet_value:
            name = i[14]
            # 공백 제외
            if name == '':
                None
            # '담당자' 제외
            elif name == '담당자':
                None
            # 담당자의 이름만 새 리스트에 추가
            else:
                name_list.append(name)

    time.sleep(120)

    # 중복 제외
    ex_name_list = list(set(name_list))

    # 정렬
    ex_name_list.sort()

    # 이름 카운트 된 숫자를 담을 리스트 생성
    num_name_list = []

    # 이름 카운트
    for i in ex_name_list:
        numb = name_list.count(i)
        num_name_list.append(str(numb))

    #print(num_name_list)

    final_list = []

    final_list.append(ex_name_list)

    final_list.append(num_name_list)

    # [['홍길동', '김민수'], ['234', '216']]의 형태로 출력
    return final_list

# quota exceed일 경우 time sleep
def quota_exceed(any_def):
    for i in range(0, 20):
        try:
            any_def
            break
        except Exception as e:

            except_name = str(e)
            matchOB = re.match(except_name, 'Quota Exceed')
            
            if matchOB != -1:
                time.sleep(120)
            else:
                print(e)
                break

# 배열 재배치하는 함수
def change_array(any_list):

    change_list = []

    for i in any_list[0]:
        change_list.append([])

    for i in range(0, len(any_list[0])):
        change_list[i].append(any_list[0][i])
        change_list[i].append(any_list[1][i])

    # [['홍길동', '234'], ['김민수', '216']]의 형태로 출력
    return change_list

if __name__ == '__main__' :
    main()
