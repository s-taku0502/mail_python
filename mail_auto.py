'''
以下の画面をコマンドプロンプトに入力する。
pip と setuptools の更新
 python -m pip install -U pip setuptools
pywin32のインストール
 pip install -U wheel pywin32
最終確認
 pip show pywin32
エラーが出なければ以下のコードを

'''

import win32com.client

# Outlookアプリケーションをインスタンス化
outlook = win32com.client.Dispatch("Outlook.Application")

# メールオブジェクトの作成
mail = outlook.CreateItem(0) # 0:メール

mail.to = 'sample@example.com'
#mail.cc = 'cc.xxx@example.com'
#mail.bcc = 'bcc.xxx@example.com'
mail.subject = '件名：テストメール'
mail.bodyFormat = 1 # 1:テキスト
mail.body = '''おはようございます'''

# 送信前に確認（Outlookが起動）
#mail.display(True)

# メール送信
mail.Send()