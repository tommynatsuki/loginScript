Option Explicit

    Dim objIE

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IE���J��
    objIE.navigate "https://direct.smbc.co.jp/aib/aibgsjsw5001.jsp"

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '���O�C��������͂���
    objIE.document.getElementsByName("S_BRANCH_CD")(0).Value = "���X�ԍ�"
    objIE.document.getElementsByName("S_ACCNT_NO")(0).Value = "�������ԍ�"
    objIE.document.getElementsByName("PASSWORD")(0).Value = "���p�X���[�h"

    '���O�C���{�^�����N���b�N����
    objIE.document.getElementsByName("bLogon.y")(0).Click

