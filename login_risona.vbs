Option Explicit

    Dim objIE

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IE���J��
    objIE.navigate "https://ib.resonabank.co.jp/IB/0102/SC_N_0102_010.aspx"

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '�L�[�{�[�h���͂�L���ɂ���
    if objIE.document.getElementById("chkUseSoftwareKeyBoard").checked Then
        objIE.document.getElementById("chkUseSoftwareKeyBoard").checked = false
    End If

    '�h�c����͂���
    objIE.document.getElementById("ctl00_cphBizConf_txtLoginId").Value = "���h�c"

    '���փ{�^�����N���b�N����
    objIE.document.getElementById("ctl00_cphBizConf_btnNext").Click

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '�L�[�{�[�h���͂�L���ɂ���
    if objIE.document.getElementById("chkUseSoftwareKeyBoard").checked Then
        objIE.document.getElementById("chkUseSoftwareKeyBoard").checked = false
    End If

    '�p�X���[�h����͂���
    objIE.document.getElementById("ctl00_cphBizConf_txtLoginPw").Value = "���p�X���[�h"

    '���փ{�^�����N���b�N����
    objIE.document.getElementById("ctl00_cphBizConf_btnLogin").Click
