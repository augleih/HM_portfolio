# -*- coding:utf-8 -*-
import win32com.client as win32
import gspread
import os
from time import sleep
import datetime
import requests
import json
import slack
import zipfile

today = datetime.date.today()
nextMon = today + datetime.timedelta(days=-today.weekday(), weeks=1)
nextFri = nextMon + datetime.timedelta(4)

names = ['이름']
dob = ['생년월일(int)']
start_mon = []
end_mon = []
start_tue = []
end_tue = []
start_wed = []
end_wed = []
start_thu = []
end_thu = []
start_fri = []
end_fri = []


file_path = os.getcwd()
os.chdir(file_path)
file_name = f'[{nextMon}~{nextFri}] 유연근무신청모음'

def getDuty():
    gc = gspread.service_account(filename="[파일이름]")
    targetSheet = gc.open("유연근무 시간 관리 (산업기능요원)").worksheet("유연근무 시간 관리표")

    for startTimeData in targetSheet.get_values('B2:B'):
        start_mon.append(str(startTimeData))

    for endTimeData in targetSheet.get_values('C2:C'):
        end_mon.append(str(endTimeData))
    
    for startTimeData in targetSheet.get_values('D2:D'):
        start_tue.append(str(startTimeData))

    for endTimeData in targetSheet.get_values('E2:E'):
        end_tue.append(str(endTimeData))

    for startTimeData in targetSheet.get_values('F2:F'):
        start_wed.append(str(startTimeData))

    for endTimeData in targetSheet.get_values('G2:G'):
        end_wed.append(str(endTimeData))

    for startTimeData in targetSheet.get_values('H2:H'):
        start_thu.append(str(startTimeData))

    for endTimeData in targetSheet.get_values('I2:I'):
        end_thu.append(str(endTimeData))

    for startTimeData in targetSheet.get_values('J2:J'):
        start_fri.append(str(startTimeData))

    for endTimeData in targetSheet.get_values('K2:K'):
        end_fri.append(str(endTimeData))

def setDutyReq(name, dob, start_mon, end_mon, start_tue, end_tue, start_wed, end_wed, start_thu, end_thu, start_fri, end_fri):
    try:
        os.mkdir(file_name.rsplit(".", maxsplit=1)[0])
    except FileExistsError:
        pass

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    if name == '전문연인 경우':
        hwp.Open(f'{file_path}/[원본] 유연근무제 신청(변경)서 - 전문연.hwp')  
        option=hwp.HParameterSet.HFindReplace
        hwp.XHwpWindows.Item(0).Visible = True
    else:
        hwp.Open(f'{file_path}/[원본] 유연근무제 신청(변경)서 - 산기요.hwp')  
        option=hwp.HParameterSet.HFindReplace
        hwp.XHwpWindows.Item(0).Visible = True

    option.FindString = "[DateOfBrith]"
    option.ReplaceString = dob
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[ReqName]"
    option.ReplaceString = name
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[StartMon]"
    option.ReplaceString = start_mon
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[EndMon]"
    option.ReplaceString = end_mon
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[StartTue]"
    option.ReplaceString = start_tue
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[EndTue]"
    option.ReplaceString = end_tue
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[StartWed]"
    option.ReplaceString = start_wed
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[EndWed]"
    option.ReplaceString = end_wed
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[StartThu]"
    option.ReplaceString = start_thu
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[EndThu]"
    option.ReplaceString = end_thu
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[StartFri]"
    option.ReplaceString = start_fri
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[EndFri]"
    option.ReplaceString = end_fri
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[DutyStartDate]"
    option.ReplaceString = datetime.datetime.strftime(nextMon, '%y년 %m월 %d일')
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[DutyEndDate]"
    option.ReplaceString = datetime.datetime.strftime(nextFri, '%y년 %m월 %d일')
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    option.FindString = "[ReqDate]"
    option.ReplaceString = datetime.datetime.strftime(today, '%Y 년  %m 월  %d 일')
    option.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    # 이미지 삽입할 위치 찾기
    while(True):
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        hwp.HParameterSet.HFindReplace.FindString = "서명"
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

        if (result == False):
            break

    imgCtrl = hwp.InsertPicture(f'{file_path}/signs/{name.replace(" ","")}_서명.png', Embedded=True)
    hwp.SetPosBySet(imgCtrl.GetAnchorPos(0))
    hwp.FindCtrl()

    hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)  # 액션 초기화
    hwp.HParameterSet.HShapeObject.TextWrap = hwp.TextWrapType("BehindText")  # 글 뒤로 배치
    hwp.HParameterSet.HShapeObject.TreatAsChar = 0  # 글자처럼 취급 해제
    hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)  # 실행

    hwp.SaveAs(os.path.join(f'{file_path}/{file_name}/', f'[{name.replace(" ", "")}] 유연근무 신청(변경)서.hwp'))
    hwp.Clear(3)

    hwp.XHwpDocuments.Item(0).Close(isDirty=False)
    hwp.Quit()
    sleep(2)

def zipResult():
    target_path = f'{file_path}/{file_name}/'
    owd = os.getcwd()
    os.chdir(target_path)

    zip_file = zipfile.ZipFile(f'[{nextMon.strftime("%Y.%m.%d")}~{nextFri.strftime("%Y.%m.%d")}] 산기요 유연근무신청.zip', 'w')
    for (path, dir, files) in os.walk(target_path):
        for file in files:
            if file.endswith('.hwp'):
                zip_file.write(os.path.join(os.path.relpath(path, target_path), file),
                compress_type=zipfile.ZIP_DEFLATED)
    
    zip_file.close()

def sendReqMsg():
    bot_token = "슬랙봇토큰"
    client = slack.WebClient(token=bot_token)

    response_sendmessage = client.chat_postMessage(
        channel = '#병역특례',
        text = f'<유저아이디>님, [{nextMon.strftime("%Y.%m.%d")} ~ {nextFri.strftime("%Y.%m.%d")}] 기간 내\n산업기능요원 유연 근무 신청서를 아래 첨부하여 공유 드립니다.'
    )
    print(f'메시지 전송 결과: {response_sendmessage.status_code}')

def sendFile():
    headers = {
        'Authorization': 'Bearer xoxb-5135182182-4557409151671-3xzzNBvgaQainYnlnhC6S38m',
    }
    files = {
        'file': open(f'{file_path}/{file_name}/[{nextMon.strftime("%Y.%m.%d")}~{nextFri.strftime("%Y.%m.%d")}] 산기요 유연근무신청.zip', 'rb'),
        'Authorization': 'Bearer ',
        'channels': (None, 'C01V61HD4GH')
        # 테스트 채널 : C04H58DKH98
    }
    response = requests.post('https://slack.com/api/files.upload', headers=headers, files=files)
    print(f'파일 전송 결과: {response.status_code}')

getDuty()
i = 0
for data in names:
    setDutyReq(data, dob[i], start_mon[i][2:-2], end_mon[i][2:-2], start_tue[i][2:-2], end_tue[i][2:-2], start_wed[i][2:-2], end_wed[i][2:-2], start_thu[i][2:-2], end_thu[i][2:-2], start_fri[i][2:-2], end_fri[i][2:-2])
    i += 1

zipResult()
sendReqMsg()
sendFile()