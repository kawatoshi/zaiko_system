VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "���Օi�Ǘ��V�X�e�����C��"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btnBuyForm_Click()
    Call ����
End Sub

Private Sub btnDeliveryForm_Click()
    Call �o��
End Sub

Private Sub btnEnd_Click()
    If Not MsgBox("���Օi�Ǘ��V�X�e�����I���܂�" & vbCrLf & "��낵���ł����H", vbYesNo) = vbYes Then
        MsgBox ("���~���܂���")
        Exit Sub
    End If
    Call SheetProtect
    Call SheetProtect("select")
    ActiveWorkbook.Save
'    ActiveWorkbook.Close
    Application.Quit
End Sub

Private Sub btnSettleForm_Click()
    Call ����
End Sub

Private Sub UserForm_Initialize()
    Call SheetProtect
    Call SheetProtect("select")
    labVer.Caption = getVer
    labOffice.Caption = Range("OFFICE_NAME").Value
End Sub
