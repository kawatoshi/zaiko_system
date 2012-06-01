VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "消耗品管理システムメイン"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btnBuyForm_Click()
    Call 入庫
End Sub

Private Sub btnDeliveryForm_Click()
    Call 出庫
End Sub

Private Sub btnEnd_Click()
    If Not MsgBox("消耗品管理システムを終了ます" & vbCrLf & "よろしいですか？", vbYesNo) = vbYes Then
        MsgBox ("中止しました")
        Exit Sub
    End If
    Call SheetProtect
    Call SheetProtect("select")
    ActiveWorkbook.Save
'    ActiveWorkbook.Close
    Application.Quit
End Sub

Private Sub btnSettleForm_Click()
    Call 決済
End Sub

Private Sub UserForm_Initialize()
    Call SheetProtect
    Call SheetProtect("select")
    labVer.Caption = getVer
    labOffice.Caption = Range("OFFICE_NAME").Value
End Sub
