# スプレッドシート接続用のサンプル
事前にGCPアカウントの作成およびシークレットJSONの取得が必要です。<BR>
参考：https://japan.appeon.com/technical/techblog/technicalblog019/

## 使い方
- シークレットJSONを任意の場所に配置して、ソースコードのJSONKEYに相対パスを記述する
- スプレッドシートを準備して公開設定（編集モード）にしてURLを取得する
- 以下を実行して依存モジュールをインストールする<BR>
pip install gspread pandas oauth2client
