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
1. 那么 nàme pr (indicating degree) so, such
2. 好玩儿 hǎowánr adj fun, amusing, interesting
3. 出去 chū qu vc to go out
4. 非常 fēicháng adv very, extremely, exceedingly
5. 糟糕 zāogāo adj in a terrible mess; how terrible
6. 下雨 xià yǔ vo to rain
7. 又 yòu adv again [See Grammar 5.]
8. 面试 miànshì v/n to interview; interview
9. 回去 huí qu vc to go back; to return
10. 冬天 dōngtiān n winter
11. 夏天 xiàtiān n summer
12. 热 rè adj hot
13. 春天 chūntiān n spring
14. 秋天 qiūtiān n autumn; fall
15. 舒服 shūfu adj comfortable"""

cc = OpenCC('s2t')
vocab = vocab.split('\n')[1:]

final = []
for line in vocab:
    line = line.split(' ')
    if line[0].endswith('.'):
        line[0] = line[0][:-1]
    try: 
        text = line[5:] if line[4] in ['adj', 'n', 'v', 'pr', 'adv', 'vc', 'vo'] else line[4:]
        final.append([int(line[0]), cc.convert(line[1]), " ".join(text)])
    except: 
        text = line[4:] if line[3] in ['adj', 'n', 'v'] else line[3:]
        final.append(['', cc.convert(line[0]), " ".join(line[3:])])     

SCOPES = ['https://www.googleapis.com/auth/presentations']

def updateTableObject(tableId, row, column, content):
        return {
            "insertText": {
                "objectId": tableId,
                "cellLocation":{
                    "rowIndex": row,
                    "columnIndex": column
                },
                "text":content
            }
        }

df = pd.read_json('config.json', orient='index')

presentationId=df.at['PRESENTATION_ID', 0]
slideNumber=int(df.at['SLIDE_NUMBER', 0])-1

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
                requests.append(updateTableObject(tableObjectId, e, d, str(final[e][d])))

        for e in range(len(final[8:])):
            e += 8
            for d in range(len(final[e])):
                requests.append(updateTableObject(secondTable, e-8, d, str(final[e][d])))


        body = {'requests': requests}
        service.presentations().batchUpdate(presentationId=presentationId, body=body).execute()

        print('Updated text in table cell.')

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
