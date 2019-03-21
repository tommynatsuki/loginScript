Option Explicit

    Dim objIE
    Dim objINPUT
    Dim n

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IEを開く
    objIE.navigate "https://www2.cr.mufg.jp/newsplus/?cardBrand=0011&lid=news_mufg"

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    'カードブランドを選択（一応）
    objIE.document.getElementsByName("cardBrand")(0).Value = "0011"

    'ＩＤとパスワードを入力する
    objIE.document.getElementsByName("webId")(0).Value = "※ＩＤ"
    objIE.document.getElementsByName("webPassword")(0).Value = "※パスワード"

    'ログインボタンをクリックする
    objIE.document.getElementById("submit1").Click

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '生年月を入力する
    objIE.document.getElementById("addAuthSelect2").checked = true
    objIE.document.getElementsByName("webBirthDay")(0).Value = "※生まれ年月(YYYYMM)"

    'ログインボタンをクリックする
    objIE.document.getElementById("submit").Click
