VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeliveryForm 
   Caption         =   "�o�ɕi"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   OleObjectBlob   =   "DeliveryForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DeliveryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnDeliveryList_Click()
    MsgBox (MakeDeliveryList)
    Sheets("�o�Ƀ��X�g").Activate
    unload DeliveryForm
    unload MainForm
End Sub

Private Sub btnMain_Click()
    unload DeliveryForm
End Sub

Private Sub btnReturnItem_Click()
    MsgBox (MakeDeliveryList)
    Sheets("�o�Ƀ��X�g").Activate
    unload DeliveryForm
    unload MainForm
End Sub

Private Sub btnLostItem_Click()
    MsgBox (makeStockList)
    Sheets("�݌Ƀ��X�g").Activate
    unload DeliveryForm
    unload MainForm
End Sub

Private Sub btnStockList_Click()
    MsgBox (makeStockList)
    Sheets("�݌Ƀ��X�g").Activate
    unload DeliveryForm
    unload MainForm
End Sub

Private Sub btnStockToDelivery_Click()
    StockToDelivery.Show
    unload StockToDelivery
End Sub

Private Sub UserForm_Activate()
    If Not ActiveWorkbook.name Like DataBaseName Then
        MsgBox ("���̃u�b�N��ł͎��s�o���܂���")
        unload DeliveryForm
    End If
End Sub

