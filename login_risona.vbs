Option Explicit

    Dim objIE

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IEを開く
    objIE.navigate "https://ib.resonabank.co.jp/IB/0102/SC_N_0102_010.aspx"

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    'キーボード入力を有効にする
    if objIE.document.getElementById("chkUseSoftwareKeyBoard").checked Then
        objIE.document.getElementById("chkUseSoftwareKeyBoard").checked = false
    End If

    'ＩＤを入力する
    objIE.document.getElementById("ctl00_cphBizConf_txtLoginId").Value = "※ＩＤ"

    '次へボタンをクリックする
    objIE.document.getElementById("ctl00_cphBizConf_btnNext").Click

    'ページが読み込まれるまで待つ
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    'キーボード入力を有効にする
    if objIE.document.getElementById("chkUseSoftwareKeyBoard").checked Then
        objIE.document.getElementById("chkUseSoftwareKeyBoard").checked = false
    End If

    'パスワードを入力する
    objIE.document.getElementById("ctl00_cphBizConf_txtLoginPw").Value = "※パスワード"

    '次へボタンをクリックする
    objIE.document.getElementById("ctl00_cphBizConf_btnLogin").Click
