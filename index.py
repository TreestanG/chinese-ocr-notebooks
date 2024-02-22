from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from opencc import OpenCC
import os.path
import pandas as pd

# Copy and paste vocab here from online pdf.
# Leave first line as is.
vocab = """ 
1. 吵 chǎo v/adj to quarrel; noisy
2. 连 lián prep even
3. 做饭 zuò fàn vo to cook; to prepare a meal
4. 报纸 bàozhǐ n newspaper
5. 广告 guǎnggào n advertisement
6. 附近 fùjìn n vicinity; neighborhood; nearby area
7. 套 tào m (measure word for suite or set)
8. 公寓 gōngyù n apartment
9. 出租 chūzū v to rent out
10. 走路 zǒu lù vo to walk
11. 分钟 fēnzhōng n minute
12. 卧室 wòshì n bedroom
13. 厨房 chúfáng n kitchen
14. 卫生间 wèishēngjiān n bathroom"""

cc = OpenCC('s2t')
vocab = vocab.split('\n')[1:]

final = []

if len(vocab) > 17: 
    print("Too many lines. Please limit total lines of string to 17.")
    exit()

for line in vocab:
    line = line.split(' ')
    if line[0].endswith('.'):
        line[0] = line[0][:-1]

    
    try: 
        text = line[5:] if line[4] in ['adj', 'n', 'v', 'pr', 'adv', 'vc', 'vo'] else line[4:]
        final.append([int(line[0]), cc.convert(line[1]), " ".join(text)])
    except: 
        text = line[4:] if line[3] in ['adj', 'n', 'v', 'pr', 'adv', 'vc', 'vo'] else line[3:]
        final.append(['', cc.convert(line[0]), " ".join(line[3:])])     

print(final)

SCOPES = ['https://www.googleapis.com/auth/presentations']

def updateTableObject(tableId, row, column, content):
    return [
        {
            "insertText": {
                "objectId": tableId,
                "cellLocation":{
                    "rowIndex": row,
                    "columnIndex": column
                },
                "text": content
            }
        },
        {
            "updateTextStyle": {
                "objectId": tableId,
                "cellLocation": {
                    "rowIndex": row,
                    "columnIndex": column
                },
                "style": {
                    "fontSize": {
                        "magnitude": 9,
                        "unit": "PT"
                    },
                    "fontFamily": "Calibri"
                },
                "textRange": {
                    "type": "ALL"
                },
                "fields": "fontSize,fontFamily"
            }
        }
    ]

df = pd.read_json('config.json', orient='index')

presentationId=df.at['PRESENTATION_ID', 0]
slideNumber = int(input("Which slide number? "))-1

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('slides', 'v1', credentials=creds)
        presentation = service.presentations().get(presentationId=presentationId).execute()

        slide = presentation['slides'][slideNumber] 
        tableObjectId = slide['pageElements'][1]['objectId'] 
        secondTable = slide['pageElements'][2]['objectId']
        
        requests = []
    
        for e in range(len(final[:8])):
            for d in range(len(final[e])):
                requests.extend(updateTableObject(tableObjectId, e, d, str(final[e][d])))

        for e in range(len(final[8:])):
            e += 8
            for d in range(len(final[e])):
                requests.extend(updateTableObject(secondTable, e-8, d, str(final[e][d])))


        body = {'requests': requests}
        service.presentations().batchUpdate(presentationId=presentationId, body=body).execute()

        print('Updated text in table cell.')

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
