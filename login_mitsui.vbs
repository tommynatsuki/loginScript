Option Explicit

    Dim objIE

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IEを開く
    objIE.navigate "https://direct.smbc.co.jp/aib/aibgsjsw5001.jsp"

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    'ログイン情報を入力する
    objIE.document.getElementsByName("S_BRANCH_CD")(0).Value = "※店番号"
    objIE.document.getElementsByName("S_ACCNT_NO")(0).Value = "※口座番号"
    objIE.document.getElementsByName("PASSWORD")(0).Value = "※パスワード"

    'ログインボタンをクリックする
    objIE.document.getElementsByName("bLogon.y")(0).Click

