Option Explicit

    Dim objIE
    Dim objINPUT
    Dim n

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    'IE���J��
    objIE.navigate "https://www2.cr.mufg.jp/newsplus/?cardBrand=0011&lid=news_mufg"

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '�J�[�h�u�����h��I���i�ꉞ�j
    objIE.document.getElementsByName("cardBrand")(0).Value = "0011"

    '�h�c�ƃp�X���[�h����͂���
    objIE.document.getElementsByName("webId")(0).Value = "���h�c"
    objIE.document.getElementsByName("webPassword")(0).Value = "���p�X���[�h"

    '���O�C���{�^�����N���b�N����
    objIE.document.getElementById("submit1").Click

    '�y�[�W���ǂݍ��܂��܂ő҂�
    Do While objIE.Busy = True Or objIE.readyState <> 4
        WScript.Sleep 100
    Loop

    '���N������͂���
    objIE.document.getElementById("addAuthSelect2").checked = true
    objIE.document.getElementsByName("webBirthDay")(0).Value = "�����܂�N��(YYYYMM)"

    '���O�C���{�^�����N���b�N����
    objIE.document.getElementById("submit").Click
