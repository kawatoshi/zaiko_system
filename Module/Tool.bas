Attribute VB_Name = "Tool"
Option Explicit

Sub �W�v�\�\��()
    Dim shtMy As Worksheet
    Dim aryShtName(3) As String
    Dim i As Long
    
    aryShtName(0) = "�ۍL��������"
    aryShtName(1) = "�e�i���g��������"
    aryShtName(2) = "�������z�ꗗ�\"
    aryShtName(3) = "����ꗗ�\"
    For i = 0 To UBound(aryShtName)
        Set shtMy = ActiveWorkbook.Sheets(aryShtName(i))
        shtMy.Visible = xlSheetVisible
    Next
End Sub
