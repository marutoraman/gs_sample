import gspread

from oauth2client.service_account import ServiceAccountCredentials

import pandas as pd


#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
SCOPE = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
JSONKEY = './secrets/my-project-8844405-05992ddf0789.json' # ここに個人のSECRET JSONを指定する

class GspreadHandler():
    """グーグルスプレッドシートを操作するクラス"""
    
    def __init__(self, url:str=None):

        #認証情報設定
        #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
        credentials = ServiceAccountCredentials.from_json_keyfile_name(JSONKEY, SCOPE)
        #OAuth2の資格情報を使用してGoogle APIにログインします。
        self.gc = gspread.authorize(credentials)
        self.url = url # スプレッドシートのURL
        self.header = self.init_fetch_sheet_header() # ２行目をシステム用がカラムを判別するのに使用する運用のため


    def get_worksheet_as_df(self, sheet_name=None):
        """指定したシートの情報をpandasのDataFrameとして取得する"""
        # シート無指定の場合は1番目を指定する
        if sheet_name == None:
            worksheet = self.gc.open_by_url(self.url).get_worksheet(0)
        # name指定の場合
        else:
            try:
                worksheet = self.gc.open_by_url(self.url).worksheet(sheet_name)
            except Exception as e:
                print(f"worksheetが開けません:{e}")
                return False
        all_data = worksheet.get_all_values()
        df = pd.DataFrame(all_data[1:],columns=all_data[0])
        #df.columns = list(df.loc[0, :]) # valuesの1行目をcolumnsとして指定
        #df.drop(0, inplace=True) # valuesの1行目(0番目のデータ)を削除
        #df.reset_index(inplace=True) # 
        #df.drop('index', axis=1, inplace=True)
        return df


    def init_fetch_sheet_header(self):
        '''
        ２行目をシステム用がカラムを判別するのに使用する運用のため
        '''
        worksheet = self.gc.open_by_url(self.url).get_worksheet(0)
        df = pd.DataFrame(worksheet.get_all_values())
        return list(df.loc[1,:]) # 2行目をシステム用のヘッダとする
        
        
    def bulk_update_row(self, datas:list, row:int):
        '''
        listを指定してスプレッドシートを一括更新
        '''
        worksheet = self.gc.open_by_url(self.url).get_worksheet(0)
        cells = worksheet.range(row, 1, len(datas) + row -1 , len(self.header))
        for row,data in enumerate(datas):
            for k,v in data.items():
                try:
                    col = self.header.index(k)
                    num = row*(len(self.header)) + col # 複数行にまたがるデータの場合でも１次元配列に格納されているため２次元→１次元に変換する
                    cells[num].value = v
                except Exception as e:
                    print(e)
                    pass

        worksheet.update_cells(cells)
        return True 
    
    def get_last_row(self):
        '''
        最終行の取得
        '''
        worksheet = self.gc.open_by_url(self.url).get_worksheet(0)
        return len(worksheet.get_all_values())
           
           
if __name__ == "__main__":
    pass

