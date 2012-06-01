VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthDegreeProcess 
   Caption         =   "月度処理"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   OleObjectBlob   =   "MonthDegreeProcess.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MonthDegreeProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnMonthDegreeProcess_Click()
'月度処理

    Dim strState As String
    Dim strPathMy As String
    Dim strMode As String
    Dim tmpSht As Worksheet

    strPathMy = ActiveWorkbook.Path
    If Not strPathMy Like "*本部*" Then
        If MsgBox("月末処理を全て自動で行います。" & vbCrLf & _
               "チェックされた書類が印刷されますのでプリンターに用紙を用意してください" & vbCrLf & _
               "準備が出来たらOKボタンを押してください。", vbOKCancel) <> vbOK Then
            MsgBox ("中止しました"): GoTo ending
        End If
        '月度処理作成および印刷
        Call MonthDegreeProcessP2(chk_maruhiro_p.Value, _
                                  chk_tenant_p.Value, _
                                  chk_bill_p.Value, _
                                  chk_uriage_p.Value, _
                                  getBilldateOnStr(getClosingdate))
        If chk_uriage_p.Value = True Then
            '保存
            ActiveWorkbook.Save
            '指定フォルダへの報告
            strState = CopyFile(ActiveWorkbook, "\\honbu\営業部報告\売上集計\消耗品")
            If Not strState Like "ok" Then
                MsgBox ("本部報告処理に異常がありました。本部への報告はキャンセルされています。" & vbCrLf & _
                       "システム管理者に連絡してください。" & vbCrLf & _
                       strState)
            Else
                MsgBox ("本部への報告を完了しました")
            End If
        End If
    Else
        Call HonbuView
    End If
    Call PrintMode("put")
ending:
    unload MonthDegreeProcess
    unload SettleForm
    unload MainForm
End Sub

Private Sub btnMonthDegreeProcess2_Click()
'月度処理作成および印刷
    If cmbBillDate.MatchFound = True Then
        Call MonthDegreeProcessP2(chk_maruhiro_p.Value, _
                                  chk_tenant_p.Value, _
                                  chk_bill_p.Value, _
                                  chk_uriage_p.Value, _
                                  cmbBillDate.text)
        unload MonthDegreeProcess
    Else
        MsgBox "出力したい月度を選択してください"
    End If
End Sub

Private Sub btnReturn_Click()
    unload MonthDegreeProcess
End Sub

Private Sub UserForm_Activate()
    Dim sht As Worksheet
    Dim list() As String
    
    If Not ActiveWorkbook.name Like DataBaseName Then
        MsgBox ("このブック上では実行出来ません")
        unload SettleForm
    Else
        Call PrintMode("get")
        If getBillDateFromSettleItems(list) = True Then
            cmbBillDate.list = list
            cmbBillDate.text = list(UBound(list))
        Else
            btnMonthDegreeProcess2.Enabled = False
        End If
    End If
End Sub

Private Sub PrintMode(mode As String)
    Dim sht As Worksheet
    Dim address(3) As String
    Set sht = ActiveWorkbook.Worksheets("tmp")
    address(0) = "h10"
    address(1) = "h11"
    address(2) = "h12"
    address(3) = "h13"
    
    Select Case mode
    Case "get"
        chk_maruhiro_p.Value = sht.Range(address(0))
        chk_tenant_p.Value = sht.Range(address(1))
        chk_bill_p.Value = sht.Range(address(2))
        chk_uriage_p.Value = sht.Range(address(3))
    Case "put"
        sht.Range(address(0)) = chk_maruhiro_p.Value
        sht.Range(address(0)).Offset(0, 1) = chk_maruhiro_p.Caption
        sht.Range(address(1)) = chk_tenant_p.Value
        sht.Range(address(1)).Offset(0, 1) = chk_tenant_p.Caption
        sht.Range(address(2)) = chk_bill_p.Value
        sht.Range(address(2)).Offset(0, 1) = chk_bill_p.Caption
        sht.Range(address(3)) = chk_uriage_p.Value
        sht.Range(address(3)).Offset(0, 1) = chk_uriage_p.Caption
    End Select
End Sub

