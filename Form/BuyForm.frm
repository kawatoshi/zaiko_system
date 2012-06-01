VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuyForm 
   Caption         =   "入庫"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   OleObjectBlob   =   "BuyForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    Sheets("入庫リスト").Activate
    unload BuyForm
    unload MainForm
End Sub

Private Sub btnMain_Click()
    unload BuyForm
End Sub

Private Sub btnReturnBuyItem_Click()
    MsgBox ("現在実装されていません")
End Sub
