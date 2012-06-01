VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockToLost 
   Caption         =   "ロスト処理"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   OleObjectBlob   =   "StockToLost.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    If Not MsgBox(strState & "を破棄します" & vbCrLf & "よろしいですか？", vbYesNo) = vbYes Then
        MsgBox ("破棄を中止しました"): Exit Sub
    End If
    MsgBox (strState & "を" & txtNumber.text & "ロスしました")
    Call makeStockList
    unload StockToLost
End Sub

Private Sub UserForm_Initialize()
    Dim lngRow As Long
    
    lngRow = Selection.Row
    labItemName.Caption = Cells(lngRow, StockList_item_name_COL)
End Sub
