VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} returnDeliveryItem 
   Caption         =   "����ԕi����"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   OleObjectBlob   =   "returnDeliveryItem.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
    If MsgBox(labItemName.Caption & " �� " & txtNum.text & "�ԕi���܂�" & vbCrLf & _
              "��낵���ł���?", vbYesNo) = vbNo Then MsgBox ("���~���܂���"): Exit Sub
    MsgBox (returnDeleveryToStock(txtDeliveryId.Caption, dblNumber))
    Call MakeDeliveryList
    unload returnDeliveryItem
End Sub

Private Sub UserForm_Initialize()
'    Dim shtMy As Worksheet
    Dim lngRow As Long
'    Set shtMy = Sheets("�o�Ƀ��X�g")
    
'    If ActiveSheet <> shtMy Then
'        MsgBox ("�o�Ƀ��X�g�V�[�g�ł������s�o���܂���")
'        Unload returnDeliveryItem
'    End If
    lngRow = Selection.Row
    txtDeliveryId.Caption = Cells(lngRow, DeliveryList_id_COL)
    labDeliveryNum.Caption = "�o�ɐ� : " & Cells(lngRow, DeliveryList_number_COL)
    labItemName.Caption = Cells(lngRow, DeliveryList_item_name_COL)
    
ending:
'    Set shtMy = Nothing
End Sub
