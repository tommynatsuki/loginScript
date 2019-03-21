Option Explicit

    Dim objIE
    Dim objINPUT
    Dim n

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IEを開く
    objIE.navigate "https://www.rakuten-sec.co.jp/web/direct_login.html"

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    'ＩＤとパスワードを入力する
    objIE.document.getElementById("form-login-id").Value = "※ＩＤ"
    objIE.document.getElementById("form-login-pass").Value = "※パスワード"

    'ログインボタンをクリックする
    objIE.document.getElementById("login-btn").Click