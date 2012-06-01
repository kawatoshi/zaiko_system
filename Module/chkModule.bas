Attribute VB_Name = "chkModule"
Option Explicit

Function chkItemCost(strCost As String) As String
'�i�ڌ����̃`�F�b�N���s�����ʂ�Ԃ�
    Dim curCost As Currency
    If IsNumeric(strCost) = False Then _
        chkItemCost = "�����������ł͂���܂��� ERROR": GoTo ending
    curCost = CCur(strCost)
    If curCost <= 0 Then _
        chkItemCost = "������0�ȉ���ݒ肷�邱�Ƃ͏o���܂��� ERROR": GoTo ending
    chkItemCost = CStr(curCost)
ending:
End Function

Function chkNumber(strNumber As String) As String
'���ʂ̃`�F�b�N���s�����ʂ�Ԃ�
    Dim dblNumber As Double
    If IsNumeric(strNumber) = False Then _
        chkNumber = "���ʂ������ł͂���܂��� ERROR": GoTo ending
    dblNumber = CDbl(strNumber)
    If dblNumber <= 0 Then _
        chkNumber = "���ʂ�0�ȉ��ɂ��邱�Ƃ͏o���܂��� ERROR": GoTo ending
    chkNumber = CStr(dblNumber)
ending:
End Function

Function chkSplit(strChecks As String, strCheck() As String) As Boolean
'������" " �Ŕz��ɂȂ邩�ǂ����𔻒肵�Č��ʂ�Ԃ�
'�z��ɂȂ�ꍇ�ɂ�strCheck()�Ɍ��ʂ�����
    strCheck() = Split(strChecks, " ")
    If UBound(strCheck) = -1 Then
        chkSplit = False
    Else
        chkSplit = True
    End If
End Function
Function chkBillType(strBilltype As String) As Boolean
'�������������@�ɍ��v���Ă��邩���m�F����
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
'�t�H�[���ɓ��͂��ꂽID�����̍��ڂƐ������Ă��邩���m�F�����ʂ�Ԃ�
    Dim item As Articles
    If strMakerCallName <> getMakerFromItemId(strId).call_name Then _
        chkItemOnForm = "���[�J�[�����������Ă��܂���": GoTo ending
    Call getItem(strId, item)
    If strItemName <> item.name Then _
        chkItemOnForm = "�i�����������Ă��܂���": GoTo ending
    If strProduct <> item.product_number Then _
        chkItemOnForm = "�^�Ԃ��������Ă��܂���": GoTo ending
    chkItemOnForm = "OK"
ending:
End Function
Function chkFolder(strPath As String) As String
'�����œn���ꂽ�t�H���_�܂��̓t�@�C�������݂��邩��
'�m�F���A���ʂ�Ԃ�:folder:file:NG

'�t�H���_�̃`�F�b�N
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
'strInput�ɗ^����ꂽ�f�[�^�����Օi�V�X�e���ɗ^����ꂽ�K���ɏ������Ă��邩�𔻒肵�A
'���ʂ�Ԃ�
    Dim lngLen As Long
    Dim strCheck As String
    
    strCheck = StrConv(strInput, vbNarrow)
    lngLen = InStr(strCheck, " ")
    If lngLen = 0 Then
        chkInputStr = "ok"
    Else
        chkInputStr = lngLen & "�����ڂɃX�y�[�X���܂܂�Ă��܂�"
    End If
End Function
Function chkJan(strJanCode As String) As String
'jan�R�[�h�����łɓo�^�ς݂����m�F���ĕԂ�
    Dim strItemId As String
    Dim itemData As Articles
    Dim makerData As makers
    
    If strJanCode Like "" Then
        chkJan = "JAN�R�[�h�����͂���Ă��܂���"
        Exit Function
    End If
    strItemId = getItemIdFromItemJanCode(strJanCode)
    Select Case strItemId
    Case ""
    Case Else
        Call getItem(strItemId, itemData)
        Call getMaker(itemData.maker_id, makerData)
        chkJan = (strJanCode & "��" & makerData.call_name & " ��" & itemData.product_number & "�Ƃ��ēo�^�ς݂ł�")
    End Select
End Function
Function chkItemHasJanCode(strItemId As String, strJan As String) As String
'������item_id�ɑΉ�����item��jan code���ݒ肳��Ă��邩���m�F���ĕԂ�
'jan�R�[�h���������Ȃ��ꍇ�ɂ�"nomatch"��Ԃ�
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
            chkItemHasJanCode = strItemId & "�ɂ�" & strJanCode & "���o�^�ς݂ł�"
        End Select
    Case Else
        chkItemHasJanCode = "ERROR"
    End Select
End Function

Public Function JANCD(argCode As Variant) As Variant
'JAN�R�[�h�̃`�F�b�N�f�B�W�b�g�����Z���ĕԂ�
'JAN�������Ⴄ�A���͐����ȊO������ꍇ�ɂ͉����󔒂�Ԃ�
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
