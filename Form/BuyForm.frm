VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuyForm 
   Caption         =   "����"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   OleObjectBlob   =   "BuyForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "BuyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuyItem_Click()
    Buyitem.Show
    unload Buyitem
End Sub

Private Sub btnBuyItemList_Click()
    Call makeBuyList
    Sheets("���Ƀ��X�g").Activate
    unload BuyForm
    unload MainForm
End Sub

Private Sub btnMain_Click()
    unload BuyForm
End Sub

Private Sub btnReturnBuyItem_Click()
    MsgBox ("���ݎ�������Ă��܂���")
End Sub
