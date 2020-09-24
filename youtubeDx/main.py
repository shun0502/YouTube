from win32com.client import Dispatch
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pandas as pd
from gspread_dataframe import get_as_dataframe, set_with_dataframe

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")

all_inbox = inbox.Items
email = []
for msg in all_inbox:
    for att in msg.Attachments:
        email.append(att)

file = email[-1]
print(file)
new_path = 'C:/Users'
os.makedirs('C:/youtuber', exist_ok=True)

file.SaveAsFile('C:/youtuber' + '\\youtuber.csv')
print('成功')
print('おめ')


df_youtuber = pd.read_csv('C:/youtuber/youtuber.csv', encoding="shift-jis")
df_youtuber = df_youtuber.reindex(columns=['チャンネルurl','調査日時','調査名','キーワード','ページ','ページ内順位',' 動画タイトル','動画url','視聴回数','チャンネル名','登録者数'])
print(df_youtuber)
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('youtube-project-288901-7f08a088f1ea.json', scope)
gc = gspread.authorize(credentials)
SPREADSHEET_KEY = '1SPtc2Qj836jGdyY3EKmEhCPU42H7umWg5HtiuQqq4m4'
workbook = gc.open_by_key(SPREADSHEET_KEY)
worksheet = workbook.worksheet('list')





li_youtuber = df_youtuber.values.tolist()
print(li_youtuber)
worksheet.append_rows(li_youtuber)


# set_with_dataframe(worksheet, df_youtuber)
# df_youtuber2 = get_as_dataframe(worksheet)
