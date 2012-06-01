Attribute VB_Name = "TestModule"
Option Explicit
Function getBillDateFromSettleItems(list() As String) As Boolean
'決済済シートから請求月のリストを抽出する
    Dim sht As Worksheet
    Dim rng As Range
    Dim data As Variant
    Dim i As Long
    Dim bill_date As String
    
    getBillDateFromSettleItems = False
    Set sht = ActiveWorkbook.Sheets("決済済")
    Set rng = getFindRange(sht, SettleArticles_bill_date_COL)
    If rng Is Nothing Then Exit Function
    If rng.rows.Count = 0 Then Exit Function
    ReDim list(rng.rows.Count - 1)
    i = 0
    list(i) = rng(1, 1).Value
    For Each data In rng
        If list(i) <> data Then
            i = i + 1
            list(i) = data
        End If
    Next
    ReDim Preserve list(i)
    getBillDateFromSettleItems = True
End Function
Sub t()
    Dim list() As String
    Call getBillDateFromSettleItems(list)
End Sub
