VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockToLost 
   Caption         =   "���X�g����"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   OleObjectBlob   =   "StockToLost.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "StockToLost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    unload StockToLost
End Sub

Private Sub btnLoss_Click()
    Dim dblNum As Double
    Dim strState As String
    
    dblNum = CDbl(txtNumber.text)
    strState = moveStockToLost(Cells(Selection.Row, 1).Value, dblNum)
    If Not MsgBox(strState & "��j�����܂�" & vbCrLf & "��낵���ł����H", vbYesNo) = vbYes Then
        MsgBox ("�j���𒆎~���܂���"): Exit Sub
    End If
    MsgBox (strState & "��" & txtNumber.text & "���X���܂���")
    Call makeStockList
    unload StockToLost
End Sub

Private Sub UserForm_Initialize()
    Dim lngRow As Long
    
    lngRow = Selection.Row
    labItemName.Caption = Cells(lngRow, StockList_item_name_COL)
End Sub
