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
from datetime import datetime
from datetime import date
import CNI_RPA

# 잔디 알람 설정
def jandiHeartBeat(location, url_='https://wh.jandi.com/connect-api/webhook/25736384/08743a26646f6f7487606607087dac71') :
    #executed_dttm = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today = date.today()

    daily_info = '요청건수 / 신고서작성 / 수리.대기'

    # 해당 날짜의 연도별 총합 시트 지정
    year_total_sheet = CNI_RPA.sec_sheet("통관실적").worksheet("{0}Total".format(today.strftime("%Y")))

    # 해당 날짜의 '월'에 해당하는 cell 검색
    year_cell = year_total_sheet.find("{0}월".format(today.strftime("%m")))

    # 해당 날짜의 '월'에 해당하는 cell의 행, 열 값
    year_total_sheet.cell(year_cell.row, year_cell.col).value

    # 해당 날짜의 월별 시트 지정
    month_sheet = CNI_RPA.sec_sheet("통관실적").worksheet("{0}{1}".format(today.strftime("%Y"), today.strftime("%m")))

    # 해당 날짜의 'YYYYMMDD'에 해당하는 cell 검색
    month_cell = month_sheet.find("{0}".format(today.strftime("%Y%m%d")))

    # 해당 날짜의 'YYYYMMDD'에 해당하는 cell의 행, 열 값
    month_sheet.cell(month_cell.row, month_cell.col).value

    requests.post(
        url=url_,
        json={
            "body": f"[{location}]\n{daily_info}\n{today.month}월 {today.day}일 : {month_sheet.cell(month_cell.row, month_cell.col+1).value} / {month_sheet.cell(month_cell.row, month_cell.col+2).value} / {month_sheet.cell(month_cell.row, month_cell.col+3).value} \n{today.month}월 총합 : {year_total_sheet.cell(year_cell.row, year_cell.col+1).value} / {year_total_sheet.cell(year_cell.row, year_cell.col+2).value} / {year_total_sheet.cell(year_cell.row, year_cell.col+3).value} ",
            "connectColor": "#FAC11B",
            "connectInfo": [
            ]
        }
    )

if __name__ == '__main__' :
    jandiHeartBeat('통관실적')