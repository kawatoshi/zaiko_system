VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettleForm 
   Caption         =   "決済"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   OleObjectBlob   =   "SettleForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SettleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInventoryReport_Click()
'棚卸し報告
    Dim strState As String
    Dim strPathMy As String
    Dim strMode As String
        
    If MsgBox("棚卸し報告を本部へ送ります。" & vbCrLf & _
              "よろしいですか？", vbOKCancel) <> vbOK Then
        MsgBox ("中止しました"): GoTo ending
    End If
    strState = CopyFile(ActiveWorkbook, "\\honbu\営業部報告\売上集計\消耗品\棚卸報告")
    If Not strState Like "ok" Then
        MsgBox ("本部報告処理に異常がありました。本部への報告はキャンセルされています。" & vbCrLf & _
               "システム管理者に連絡してください。" & vbCrLf & _
               strState)
    Else
        MsgBox ("本部への報告を完了しました")
    End If
ending:
End Sub

Private Sub btnMain_Click()
    Sheets("メイン").Select
    unload SettleForm
End Sub

Private Sub btnmakeApprovalList_Click()
    MsgBox (makeApprovalList)
    Range("a1").Select
    unload SettleForm
    unload MainForm
End Sub

Private Sub btnMonthDegreeProcess_Click()
    MonthDegreeProcess.Show
End Sub

Private Sub btnStockReport_Click()
'在庫報告
    Dim strState As String
    Dim strPathMy As String
    Dim strMode As String
        
    If MsgBox("現在の在庫状況を本部へ報告します。" & vbCrLf & _
              "よろしいですか？", vbOKCancel) <> vbOK Then
        MsgBox ("中止しました"): GoTo ending
    End If
    strState = CopyFile(ActiveWorkbook, "\\honbu\営業部報告\売上集計\消耗品\在庫状況報告")
    If Not strState Like "ok" Then
        MsgBox ("本部報告処理に異常がありました。本部への報告はキャンセルされています。" & vbCrLf & _
               "システム管理者に連絡してください。" & vbCrLf & _
               strState)
    Else
        MsgBox ("本部への報告を完了しました")
    End If
ending:

End Sub

Private Sub CommandButton2_Click()
    MsgBox ("この機能は実装されていません")
End Sub


Private Sub UserForm_Activate()
    If Not ActiveWorkbook.name Like DataBaseName Then
        MsgBox ("このブック上では実行出来ません")
        unload SettleForm
    End If
End Sub

