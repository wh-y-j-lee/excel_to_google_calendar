from __future__ import print_function
import pandas as pd
from datetime import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
# 아래의 scope는 read/write 권한으로 readonly를 쓰면 insert가 오류
SCOPES = ['https://www.googleapis.com/auth/calendar','https://www.googleapis.com/auth/calendar.events']

## make event 부분에 datetime(시분초까지)으로 할 지 date로 할 지 원하는거로 가져다 쓰면 됨
fmt_datetime = "%Y-%m-%dT%H:%M:%S"
fmt_date = "%Y-%m-%d"

## excel 위치
excel_path = '배포용 일정표.xlsx'

## excel timestamp를 파이썬용으로 바꿔주기
def excel_time_to_timestamp(time_from_excel):
    time_from_excel -= 2    # 내 엑셀의 시간이 2일 더 늦게 나와서 -2 해준다. (my excel time value is two days slower than day)
    return datetime.fromtimestamp((time_from_excel-9/24-17-365*70)*86400)

## 위에거 역변환
def timestamp_to_excel_time(time_from_timestamp):
    return int(time_from_timestamp/86400+9/24+17+365*70)

## google용 time str형식으로 바꿔주기 그냥 엑셀 timestamp 입력해도 되고 datetime으로 변환한 time stamp 입력해도 된다.
def set_google_time_str(timestamp,fmt):
    if type(timestamp) == int:
        timestamp = excel_time_to_timestamp(timestamp)
    return timestamp.strftime(fmt)

## 퍼온거  해당 날짜의 event 긁어오기
def get_date_events(date, events):
    lst = []
    date = date
    for event in events:
        if event.get('start').get('date'):
            d1 = event['start']['date']
            if d1 == date:
                lst.append(event)
    return lst
## event 만드는 함수; summary = title, start, end는 시간 정해주는 거 여기서 date 말고 datetime으로 쓰고싶으면 바꿔서 사용
## email 알람보내기, 몇분 전, 몇일 전 알람 등등 설정할 수 있다.
def make_event(title,date):
    new_event = {
        'summary': title,
        # 'location': 'My home',
        # 'description': 'where is description?',
        # 'start': {
        #   'dateTime': '2019-04-05T07:00:00',
        #   'timeZone': 'Asia/Seoul',
        # },
        # 'end': {
        #   'dateTime': '2019-04-05T09:00:00',
        #   'timeZone': 'Asia/Seoul',
        # },

        'start': {
            'date': date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'date': date,
            'timeZone': 'Asia/Seoul',
        },
        # 'recurrence': [ # 아마 반복
        #   'RRULE:FREQ=DAILY;COUNT=2'
        # ],
        # 'attendees': [
        #   {'email': 'paulyongju@gmail.com'},
        #   {'email': 'paulyongju@yonsei.ac.kr'},
        # ],
        # 'reminders': {
        #     'useDefault': False,
        #     'overrides': [
        #         {'method': 'email', 'minutes': 24 * 60},
        #         {'method': 'popup', 'minutes': 10},
        #     ],
        # },
    }
    return new_event

## read Excel
def make_date_keyword_list(excel_path):
    my_excel = pd.read_excel(excel_path, sheet_name='Sheet1') # pandas로 읽는데 원하는 Sheet 정해서 가져올 수 있다.

    # my_excel.keys() 출력해보면 A, B, C, D 열의 index대로 출력이 되는데 제일 윗 행으로 dictionary 형태로 만들어져 있다.

    # indexing할 행 미리 저장
    excel_date_key = my_excel.keys()[1] #내 엑셀은 2번째('B')행에 날짜가 적혀있다
    keyword_key = my_excel.keys()[3]    # 대표 keyword는 'D'행

    # title과 datetime 저장할 리스트 생성
    my_date_keyword_list = []

    # 날짜말고 keyword로 search 시작(해당 행에 있는 값들만 차례대로 출력이 된다.)
    for n, i in enumerate(my_excel.get(keyword_key)):
        if str(i) == 'nan': # 빈칸은 str 타입의 nan
            continue    # 빈칸 건너뛰기
        else:   # 빈칸이 아닐 경우
            next_datetime = my_excel.get(excel_date_key)[n]
            if str(next_datetime).isdigit():    # 날짜 행에 있는 값도 nan인 부분 있어서 숫자인 경우일 때만 사용
                # 이거는 오늘 지금 이시간 이후에 있는 일정들만 찾는 조건
                if next_datetime>=timestamp_to_excel_time(datetime.now().timestamp()):
                    # [date, title]을 반환할 리스트에 append하여 추가
                    todo_list_in_same_day = str(i).split('\n')
                    for j in todo_list_in_same_day:
                        my_date_keyword_list.append([set_google_time_str(next_datetime, fmt_date), j])
    return my_date_keyword_list

def main():
    # cred이랑 token 검사 =====================================================================
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) # 여기 SCOPE에서 읽고 쓰기 권한 설정
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)    # build 해서 이제부터 calendar와 상호 작용 가능

    now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time #google식 타임

    my_list = make_date_keyword_list(excel_path)    # 내 엑셀 일정표 불러오기

    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          maxResults=500, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', []) # 현재 이후 500개 event들 불러오기
    ##########################
    ##    matching events   ##
    ##########################
    for date, title in my_list: # 내 리스트에서 date, title 값 하나씩 받아오고
        new_event = make_event(title, date) # 위에 받아온거로 event 새로 생성
        # 새로 생성한 이벤트의 날짜와 기존 구글 캘린더에 있는 날짜가 같을 때의 event들을 date_events에 저장
        date_events = get_date_events(new_event['start']['date'], events)   #만약 datetime까지 하고싶으면 event 만드는 곳이랑 여러함수들에서 date->datetime으로 변경해주면 된다.
        lst = [x['summary'] for x in date_events]   #해당 날짜의 event들의 summary(title)만 리스트로 저장
        if title in lst:    # 엑셀에서 읽어온 title이 그날 list안에 있으면 continue해서 넘김(또 추가할 필요는 없으니까)
            continue
        else:
            #아닐 경우 새로 event 추가
            service.events().insert(calendarId='primary', body=new_event).execute()
            print(title,date,'executed')

if __name__ == '__main__':
    main()


