Option Explicit

    Dim objIE
    Dim objINPUT
    Dim n

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IE���J��
    objIE.navigate "https://viewsnet.jp/default.htm#_ga=2.45984812.940755902.1553138155-108689320.1553138155"

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '�h�c�ƃp�X���[�h����͂���
    objIE.document.getElementById("id").Value = "���h�c"
    objIE.document.getElementById("pass").Value = "���p�X���[�h"

    'INPUT�̃^�O���W�߂�
    Set objINPUT = objIE.Document.getElementsByTagName("INPUT")

    'INPUT�̒����烍�O�C����T���ăN���b�N����
    For n = 0 To objINPUT.Length - 1
        If Instr(objINPUT(n).alt,"���O�C��") > 0 Then
            objINPUT(n).Click
            Exit For
        End If
    Next
