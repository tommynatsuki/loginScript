Option Explicit

    Dim objIE
    Dim objINPUT
    Dim n

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IE���J��
    objIE.navigate "https://www.rakuten-sec.co.jp/web/direct_login.html"

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '�h�c�ƃp�X���[�h����͂���
    objIE.document.getElementById("form-login-id").Value = "���h�c"
    objIE.document.getElementById("form-login-pass").Value = "���p�X���[�h"

    '���O�C���{�^�����N���b�N����
    objIE.document.getElementById("login-btn").Click