import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pandas as pd


class GoogleSheetToCSV:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.creds = None


    def authenticate(self):
        """Google Sheets API 인증을 처리합니다."""
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                self.creds = flow.run_local_server(port=0)

            # token.json 파일에 인증 정보 저장
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())

    def export_to_csv(self, spreadsheet_id, range_name, output_file):
        """
        구글 시트의 내용을 CSV 파일로 저장합니다.

        Parameters:
        spreadsheet_id (str): 구글 시트의 ID
        range_name (str): 데이터를 가져올 범위 (예: 'Sheet1!A1:Z1000')
        output_file (str): 저장할 CSV 파일 경로
        """
        try:
            # API 서비스 생성
            service = build('sheets', 'v4', credentials=self.creds)

            # 시트 데이터 가져오기
            sheet = service.spreadsheets()
            result = sheet.values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])

            if not values:
                print('데이터가 없습니다.')
                return

            # DataFrame으로 변환
            df = pd.DataFrame(values[1:], columns=values[0])

            # CSV 파일로 저장
            df.to_csv(output_file, index=False, encoding='utf-8-sig')
            print(f'CSV 파일이 성공적으로 저장되었습니다: {output_file}')

        except Exception as e:
            print(f'오류가 발생했습니다: {str(e)}')


def main():
    # 스프레드시트 ID는 URL에서 찾을 수 있습니다
    # 예: https://docs.google.com/spreadsheets/d/[spreadsheet_id]/edit
    SPREADSHEET_ID = '1soq_jj1eJvtpJs1pTaAF7Qn1pUXNynSXRW-ByuETkMc'
    RANGE_NAME = '시트1!A1:Z1000'
    OUTPUT_FILE = 'output.csv'

    converter = GoogleSheetToCSV()
    converter.authenticate()
    converter.export_to_csv(SPREADSHEET_ID, RANGE_NAME, OUTPUT_FILE)


if __name__ == '__main__':
    main()