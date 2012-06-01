Attribute VB_Name = "chkModule"
Option Explicit

Function chkItemCost(strCost As String) As String
'品目原価のチェックを行い結果を返す
    Dim curCost As Currency
    If IsNumeric(strCost) = False Then _
        chkItemCost = "原価が数字ではありません ERROR": GoTo ending
    curCost = CCur(strCost)
    If curCost <= 0 Then _
        chkItemCost = "原価に0以下を設定することは出来ません ERROR": GoTo ending
    chkItemCost = CStr(curCost)
ending:
End Function

Function chkNumber(strNumber As String) As String
'数量のチェックを行い結果を返す
    Dim dblNumber As Double
    If IsNumeric(strNumber) = False Then _
        chkNumber = "数量が数字ではありません ERROR": GoTo ending
    dblNumber = CDbl(strNumber)
    If dblNumber <= 0 Then _
        chkNumber = "数量を0以下にすることは出来ません ERROR": GoTo ending
    chkNumber = CStr(dblNumber)
ending:
End Function

Function chkSplit(strChecks As String, strCheck() As String) As Boolean
'文字列が" " で配列になるかどうかを判定して結果を返す
'配列になる場合にはstrCheck()に結果を入れる
    strCheck() = Split(strChecks, " ")
    If UBound(strCheck) = -1 Then
        chkSplit = False
    Else
        chkSplit = True
    End If
End Function
Function chkBillType(strBilltype As String) As Boolean
'引数が請求方法に合致しているかを確認する
    Dim varList As Variant
    Dim varCheck As Variant
    chkBillType = False
    varList = Range("bill_type_list")
    For Each varCheck In varList
        If strBilltype = CStr(varCheck) Then _
            chkBillType = True: Exit Function
    Next
End Function
Function chkItemOnForm(strId As String, strItemName As String, _
                        strMakerCallName As String, strProduct As String) As String
'フォームに入力されたIDが他の項目と整合しているかを確認し結果を返す
    Dim item As Articles
    If strMakerCallName <> getMakerFromItemId(strId).call_name Then _
        chkItemOnForm = "メーカー名が整合していません": GoTo ending
    Call getItem(strId, item)
    If strItemName <> item.name Then _
        chkItemOnForm = "品名が整合していません": GoTo ending
    If strProduct <> item.product_number Then _
        chkItemOnForm = "型番が整合していません": GoTo ending
    chkItemOnForm = "OK"
ending:
End Function
Function chkFolder(strPath As String) As String
'引数で渡されたフォルダまたはファイルが存在するかを
'確認し、結果を返す:folder:file:NG

'フォルダのチェック
    Dim lngERROR As Long
        
    On Error GoTo ER
    If Len(Dir(strPath, vbDirectory)) > 0 Then
        If (GetAttr(strPath) And vbDirectory) = vbDirectory Then
            chkFolder = "folder"
        Else
            chkFolder = "file"
        End If
        On Error GoTo 0
    Else
        chkFolder = "NG"
    End If
    Exit Function
ER:
    chkFolder = "ERROR"
End Function
Function chkInputStr(strInput As String) As String
'strInputに与えられたデータが消耗品システムに与えられた規則に準拠しているかを判定し、
'結果を返す
    Dim lngLen As Long
    Dim strCheck As String
    
    strCheck = StrConv(strInput, vbNarrow)
    lngLen = InStr(strCheck, " ")
    If lngLen = 0 Then
        chkInputStr = "ok"
    Else
        chkInputStr = lngLen & "文字目にスペースが含まれています"
    End If
End Function
Function chkJan(strJanCode As String) As String
'janコードがすでに登録済みかを確認して返す
    Dim strItemId As String
    Dim itemData As Articles
    Dim makerData As makers
    
    If strJanCode Like "" Then
        chkJan = "JANコードが入力されていません"
        Exit Function
    End If
    strItemId = getItemIdFromItemJanCode(strJanCode)
    Select Case strItemId
    Case ""
    Case Else
        Call getItem(strItemId, itemData)
        Call getMaker(itemData.maker_id, makerData)
        chkJan = (strJanCode & "は" & makerData.call_name & " の" & itemData.product_number & "として登録済みです")
    End Select
End Function
Function chkItemHasJanCode(strItemId As String, strJan As String) As String
'引数のitem_idに対応するitemにjan codeが設定されているかを確認して返す
'janコードが整合しない場合には"nomatch"を返す
    Dim strJanCode As String
    Dim itemData As Articles
    
    Select Case getItem(strItemId, itemData)
    Case "OK"
        strJanCode = itemData.JAN_CODE
        Select Case strJanCode
        Case ""
            If Not strJan Like "" Then
                Exit Function
            Else
                chkItemHasJanCode = "nomatch"
            End If
        Case Else
            chkItemHasJanCode = strItemId & "には" & strJanCode & "が登録済みです"
        End Select
    Case Else
        chkItemHasJanCode = "ERROR"
    End Select
End Function

Public Function JANCD(argCode As Variant) As Variant
'JANコードのチェックディジットを演算して返す
'JAN桁数が違う、又は数字以外がある場合には何も空白を返す
    Dim strCode As String
    Dim intDigit As Integer
    Dim intPos As Integer
    Dim intCD As Integer
    If IsNull(argCode) Then Exit Function
    If Not IsNumeric(argCode) Then Exit Function
    Select Case Len(argCode)
    Case 7, 8
        strCode = Left(argCode, 7)
        For intPos = 1 To 7 Step 2
            intDigit = intDigit + CInt(Mid(strCode, intPos, 1))
        Next
        intDigit = intDigit * 3
        For intPos = 2 To 6 Step 2
            intDigit = intDigit + CInt(Mid(strCode, intPos, 1))
        Next
        
    Case 12, 13
        strCode = Left(argCode, 12)
        For intPos = 2 To 12 Step 2
            intDigit = intDigit + CInt(Mid(strCode, intPos, 1))
        Next
        intDigit = intDigit * 3
        For intPos = 1 To 11 Step 2
            intDigit = intDigit + CInt(Mid(strCode, intPos, 1))
        Next
    Case Else
        Exit Function
    End Select
    intCD = intDigit Mod 10
    If intCD <> 0 Then
        intCD = 10 - intCD
    End If
    JANCD = strCode & Format(intCD)
End Function
