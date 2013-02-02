Attribute VB_Name = "Tool"
Option Explicit

Sub 集計表表示()
    Dim shtMy As Worksheet
    Dim aryShtName(3) As String
    Dim i As Long
    
    aryShtName(0) = "丸広請求内訳"
    aryShtName(1) = "テナント請求内訳"
    aryShtName(2) = "請求金額一覧表"
    aryShtName(3) = "売上一覧表"
    For i = 0 To UBound(aryShtName)
        Set shtMy = ActiveWorkbook.Sheets(aryShtName(i))
        shtMy.Visible = xlSheetVisible
    Next
End Sub
