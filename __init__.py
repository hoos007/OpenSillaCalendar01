import csv
import datetime
import json
import logging
import sys
import time
from encodings import utf_8_sig

import azure.functions as func
import requests
import os.path
from datetime import datetime as dt

from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

sys.path.append('./')

def main(mytimer: func.TimerRequest, tablePath:func.Out[str]) -> None:
    try:
        utc_timestamp = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc).isoformat()

        if mytimer.past_due:
            logging.info('The timer is past due!')

        logging.info('Python timer trigger function ran at %s', utc_timestamp)
        tags1 = schedule()
        
        data__schedule = {
            "data": tags1,
            "PartitionKey": time.time(),
            "RowKey": time.time()
        }
        print("data__schedule")
        
        tablePath.set(json.dumps(data__schedule))
        
        setSchedule(tags1)
    except Exception as e:
        raise Exception(e)
    


def schedule():
    
    url = "https://www.silla.ac.kr/ko/index.php?pCode=sillacalendar&mode=list.year"
    
    #/tmp/ : AzureFunction에서 파일 쓰기를 하려면 이 위치 아래에서만 가능함.
    #본인 컴퓨터에서 실행하려면 모든 주소앞에 /tmp/제거 후 실행
    filename = "/tmp/schedule_index.csv"
    f = open(filename, "w", encoding="utf_8_sig",newline= "")
    writer = csv.writer(f)
    title = "year,month,date,title".split(",")
    writer.writerow(title)

    res = requests.get(url, verify=False)
    res.raise_for_status
    soup = BeautifulSoup(res.text, "lxml")
    
    data_results = []
    try:
        data_rows_a = soup.find_all("div", attrs={"class": "sch-datalist"})
        for a in data_rows_a:
            data_rows_b = a.find("ol").find_all("li")
            for b in data_rows_b:
                data_month = b.find_all("span", attrs={"class": "mtxt"})
                data_day = b.find_all("span", attrs={"class": "dtxt"})
                for c, d in zip(data_month,data_day):
                    yy,mm = c.get_text().strip().split(".")
                    dd = d.get_text().strip()
                    cont = b.find("div", attrs={"class":"pcont"}).get_text().strip()
                    data = [yy, mm, dd, cont]
                    data_results.append(data)
                    writer.writerow(data)
        print(data_results)
    except IndexError:
        pass
    return data_results
        
def setSchedule(array):

    #Function app 에 deploy 하기전에 먼저 본인 컴퓨터에서 실행해 구글 로그인 창에서 로그인하고
    #토큰을 만들어 같이 올려야 동작함
    #function app 에서는 로그인 할 수 없음.
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    elif not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    print("구글 계정 인증 완료")
    #본인 캘린더 아이디 입력
    #기본캘린더 : primary
    cId = ''

    service = build('calendar', 'v3', credentials=creds)

    events_result = service.events().list(calendarId=cId,
                                            timeMin= str(dt.today().year-1)+'-12-01T00:00:00Z',
                                            maxResults=400, singleEvents=True,
                                            orderBy='startTime').execute()
    events = events_result.get('items', [])
    
    deleteItem = []
    for event in events:
        flag1 = 0
        for a in array:
            if(event['summary'] == a[3]):
                flag1 = 1
                break
        if(flag1 == 0):
            deleteItem.append(event['id'])
    
    addItem = []
    for idx ,a in enumerate(array):
        flag1 = 0
        for event in events:
            if(event['summary'] == a[3]):
                flag1 = 1
                break
        if(flag1 == 0):
            addItem.append(idx)
    
    for deleteId in deleteItem:
        service.events().delete(calendarId=cId, eventId=deleteId).execute()
    
    flag = 0
    for i in addItem:
        if flag == 0:
            if array[i+1][3] == array[i][3]:
                flag = 1
            calendarData = {
                'summary': array[i][3],
                'description': array[i][3],
                'start': {
                    'timeZone': 'Asia/Seoul',
                    'date': array[i][0]+"-"+array[i][1]+"-"+array[i][2]
                },
                'end': {
                    'timeZone': 'Asia/Seoul',
                    'date': array[i+1][0]+"-"+array[i+1][1]+"-"+array[i+1][2] if flag == 1 else array[i][0]+"-"+array[i][1]+"-"+array[i][2]
                },
                'transparency': 'transparent',
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 15 * 60},
                    ],
                },
            }
            event = service.events().insert(calendarId=cId, body=calendarData).execute()

            print("Calendar Name : " + event['summary'] + ", Calendar Link : " + event['htmlLink'])

            print("캘린더 추가 완료")
        else:
            flag = 0
    print("업데이트 완료")