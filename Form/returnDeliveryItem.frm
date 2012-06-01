VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} returnDeliveryItem 
   Caption         =   "売場返品処理"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   OleObjectBlob   =   "returnDeliveryItem.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "returnDeliveryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    unload returnDeliveryItem
End Sub

Private Sub btnOk_Click()
    Dim dblNumber As Double
    dblNumber = CDbl(txtNum.text)
    If MsgBox(labItemName.Caption & " を " & txtNum.text & "個返品します" & vbCrLf & _
              "よろしいですか?", vbYesNo) = vbNo Then MsgBox ("中止しました"): Exit Sub
    MsgBox (returnDeleveryToStock(txtDeliveryId.Caption, dblNumber))
    Call MakeDeliveryList
    unload returnDeliveryItem
End Sub

Private Sub UserForm_Initialize()
'    Dim shtMy As Worksheet
    Dim lngRow As Long
'    Set shtMy = Sheets("出庫リスト")
    
'    If ActiveSheet <> shtMy Then
'        MsgBox ("出庫リストシートでしか実行出来ません")
'        Unload returnDeliveryItem
'    End If
    lngRow = Selection.Row
    txtDeliveryId.Caption = Cells(lngRow, DeliveryList_id_COL)
    labDeliveryNum.Caption = "出庫数 : " & Cells(lngRow, DeliveryList_number_COL)
    labItemName.Caption = Cells(lngRow, DeliveryList_item_name_COL)
    
ending:
'    Set shtMy = Nothing
End Sub
