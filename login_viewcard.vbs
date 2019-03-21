Option Explicit

    Dim objIE
    Dim objINPUT
    Dim n

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IEを開く
    objIE.navigate "https://viewsnet.jp/default.htm#_ga=2.45984812.940755902.1553138155-108689320.1553138155"

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    'ＩＤとパスワードを入力する
    objIE.document.getElementById("id").Value = "※ＩＤ"
    objIE.document.getElementById("pass").Value = "※パスワード"

    'INPUTのタグを集める
    Set objINPUT = objIE.Document.getElementsByTagName("INPUT")

    'INPUTの中からログインを探してクリックする
    For n = 0 To objINPUT.Length - 1
        If Instr(objINPUT(n).alt,"ログイン") > 0 Then
            objINPUT(n).Click
            Exit For
        End If
    Next
