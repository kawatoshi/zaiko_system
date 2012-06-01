Attribute VB_Name = "makeModule"
Function MakeDeliveryList() As String
'�o�Ƀ��X�g���쐬����
    Dim Ditem() As DeliveryArticles
    Dim dlist() As DeliveryList
    Dim shtDelivery As Worksheet
    Dim shtDlist As Worksheet
    Dim rngMy As Range
    Dim i As Long
    Dim varId As Variant
    Dim strId() As String
    Dim strState As String
    Dim lngRows As Long
    Set shtDlist = ActiveWorkbook.Sheets("�o�Ƀ��X�g")
    Set shtDelivery = ActiveWorkbook.Sheets("�o��")
    Set rngMy = getFindRange(shtDelivery, DeliveryArticles_id_COL)
    '�o�Ƀ��X�g�V�[�g������
    Call SheetUnprotect(shtDlist)
    Call clearList(shtDlist)
    
    If rngMy Is Nothing Then MakeDeliveryList = "�o�ɂ͂���܂���": GoTo ending
    ReDim Ditem(rngMy.rows.Count)
    ReDim dlist(rngMy.rows.Count)
    'varId = rngMy.Value
    lngRows = rngMy.rows.Count
    '�o�Ƀf�[�^�擾
    For i = 1 To lngRows
        strState = getDeliveryItem(CStr(rngMy.Cells(i).Value), Ditem(i))
    Next
    '�f�[�^�쐬�y�ѓ]�L
    Set rngMy = shtDlist.Cells(getEndRow("a", shtDlist) + 1, 1)
    For i = 1 To lngRows
        strState = postDeliveryList(Ditem(i), dlist(i))
        strState = putDeliveryList(rngMy, dlist(i))
        Set rngMy = rngMy.Offset(1, 0)
    Next
    Set rngMy = shtDlist.Cells(rngMy.Row, DeliveryList_sum_COL)
    rngMy.Value = getSumFunction(getFindRange(shtDelivery, DeliveryList_sum_COL))
    Call FilterOff(shtDlist)
    MakeDeliveryList = CStr(i - 1) & "���̏o�ɂ�����܂���"
    shtDlist.Visible = xlSheetVisible
    shtDlist.Columns(DeliveryList_sum_COL).AutoFit
    Call WindowFreezePanes(shtDlist, shtDlist.Range("a6"))
    Selection.End(xlDown).Select
ending:
    Call SheetProtect("select")
    Set shtDelivery = Nothing
    Set rngMy = Nothing
End Function
Function makeStockList() As String
'�݌Ƀ��X�g���쐬����
    Dim rngMy As Range
    Dim shtMy As Worksheet
    Dim shtStock As Worksheet
    Dim i As Long
    Dim strId() As String
    Dim item As Articles
    Dim varCol As Variant
    
    Set shtStock = Sheets("�݌�")
    Set rngMy = getFindRange(shtStock, StockArticles_item_id_COL)
    Set shtMy = Sheets("�݌Ƀ��X�g")
    Call SheetUnprotect(shtMy)
    Call FilterOff(shtMy)
    Call clearList(shtMy)
    
    If rngMy Is Nothing Then makeStockList = "�݌ɂ͂���܂���": GoTo ending
    '�S�i��id�̎擾
    ReDim strId(rngMy.Count - 1)
    For i = 0 To rngMy.Count - 1
        strId(i) = rngMy.Cells(i + 1)
    Next
    Call DuplicationMerge(strId)
    Set rngMy = shtMy.Cells(getEndRow("a", shtMy) + 1, 1)
    For i = 0 To UBound(strId)
        '�i�ڂ̎擾
        Call getItem(strId(i), item)
        '�݌Ƀ��X�g��������
        Call putStockList(rngMy, postStocklist(item))
        Set rngMy = rngMy.Offset(1, 0)
    Next
    makeStockList = CStr(i) & " �̕i�ڂ�����܂�"
    '���v��
    varCol = Array(StockList_cost_COL, StockList_sum_COL, StockList_item_price_COL, StockList_number_COL)
    For i = 0 To UBound(varCol)
        Set rngMy = shtMy.Cells(rngMy.Row, CLng(varCol(i)))
        rngMy.Value = getSumFunction(getFindRange(shtMy, CLng(varCol(i))))
    Next
    With shtMy
        .Columns(StockList_cost_COL).AutoFit
        .Columns(StockList_number_COL).AutoFit
        .Columns(StockList_sum_COL).AutoFit
        .Columns(StockList_item_price_COL).AutoFit
        .Visible = xlSheetVisible
        Call WindowFreezePanes(shtMy, .Range("a6"))
    End With
    Selection.End(xlDown).Select
ending:
    Call SheetProtect("select")
    Set shtStock = Nothing
    Set shtMy = Nothing
End Function
Sub makeMonthDistributerAccount()
'���x�̕s��r�����ׂ��쐬����
    Dim strState As String
    Dim varState As Variant
    Dim rngMy As Range
    Set rngMy = getFindRange(Sheets("���ύ�"), SettleArticles_id_COL)
    If rngMy Is Nothing Then
        MsgBox ("���ς��ꂽ�i�͂���܂���"): GoTo ending
    End If
    For Each varState In rngMy
        strState = strState & CStr(varState) & " "
    Next
    strState = Trim(strState)
    Call makeDistributerAccount(strState)
ending:
End Sub
Function makeBuyList()
'���Ƀ��X�g���쐬����
    Dim shtBuyItem As Worksheet
    Dim shtMy As Worksheet
    Dim rngMy As Range
    Dim rngPut As Range
    Dim varId As Variant
    Dim Buyitem As BuyArticles
    
    Set shtBuyItem = Sheets("����")
    Set shtMy = Sheets("���Ƀ��X�g")
    
    shtMy.Visible = xlSheetVisible
    Call SheetUnprotect(shtMy)
    
    Call clearList(shtMy)
    Call FilterOff(shtMy)
    
    Set rngMy = getFindRange(shtBuyItem, BuyArticles_id_COL)
    Set rngPut = shtMy.Cells(getEndRow("a", shtMy) + 1, 1)
    If rngMy Is Nothing Then MsgBox ("���ɂ͂���܂���"): GoTo ending
    For Each varId In rngMy
        Call getBuyItem(CStr(varId), Buyitem)
        Call putBuyList(rngPut, postBuyItemToBuyList(Buyitem))
        Set rngPut = rngPut.Offset(1, 0)
    Next
    Call WindowFreezePanes(shtMy, shtMy.Range("a6"))
    
ending:
    Call SheetProtect("select")
    Set rngPut = Nothing
    Set rngMy = Nothing
    Set shtBuyItem = Nothing
    Set shtMy = Nothing
End Function

Sub makeDistributerAccount(strSettleItemIds As String)
'�s��r�����ׂ��쐬����
    Dim strSID() As String
    Dim varId As Variant
    Dim shtMy As Worksheet
    Dim rngMy As Range
    Set shtMy = Sheets("�s��r������")
    
    Call SheetUnprotect(shtMy)
    Call clearList(shtMy)
    Call FilterOff(shtMy)
    Set rngMy = shtMy.Cells(getEndRow("a", shtMy) + 1, 1)
    If chkSplit(strSettleItemIds, strSID) = False Then GoTo ending
    For Each varId In strSID
        Call putDistributerAccount(rngMy, postDistributerAccount(CStr(varId)))
        Set rngMy = rngMy.Offset(1, 0)
    Next
    shtMy.Visible = xlSheetVisible
ending:
    Call SheetProtect("select")
    Set shtMy = Nothing
    Set rngMy = Nothing
End Sub
Function MakeLossList(strName As String) As String
'���X���X�g���쐬����
    Dim shtMy As Worksheet
    MakeLossList = "nil"
    If MakeSheet(strName) Like "OK" Then
        Call putProtectList(strName, "f")
        Call initLoss(strName)
    End If
    Set shtMy = Sheets(strName)
    Call FilterOff(shtMy)
    Call clearList(shtMy)
    
    MakeLossList = "OK"
End Function
Function MakeSheet(strName As String) As String
'activeworkbook�ɓ����̃V�[�g���Ȃ��ꍇ�A�V�[�g�̍Ō��
'strName�ŃV�[�g���쐬����
    Dim strShtName() As String
    Dim varNAME As Variant
    
    MakeSheet = "nil"
    strShtName = Split(getWorksheetNames(), ":")
    For Each varNAME In strShtName
        If UCase(strName) Like UCase(varNAME) Then
            GoTo ending
        End If
    Next
    With Worksheets
        .Add after:=.item(.Count)
    End With
    ActiveSheet.name = strName
    MakeSheet = "OK"
ending:
End Function
Function makeApprovalList() As String
'�ۍL���F�肢���쐬����
    Dim Ditem As DeliveryArticles
    Dim rngMy As Range
    Dim shtMy As Worksheet
    Dim varData As Variant
    Dim dteClosingDate As Date
    Dim strApprovalItemId As String
    Dim shtPaste As Worksheet
    Dim rngPaste As Range
    Dim lngCounter As Long
    Const BillType As String = "�[�i�`�["
    
    dteClosingDate = getClosingdate
    Set shtPaste = Sheets("�ۍL���F��")
    Set shtMy = Sheets("�o��")
    Call SheetUnprotect(shtPaste)
    Call clearList(shtPaste)
    '���t�f�[�^�擾
    Set rngMy = getFindRange(shtMy, DeliveryArticles_delivery_date_COL)
    If rngMy Is Nothing Then makeApprovalList = "�o�ɂ����ǋ��͂���܂���ł���": GoTo ending
    Set rngPaste = shtPaste.Cells(getEndRow("a:z", shtPaste) + 1, 1)
    For Each varData In rngMy
        If varData < getClosingdate And varData.Offset(0, -1).Value Like BillType Then
            lngCounter = lngCounter + 1
            strApprovalItemId = shtMy.Cells(varData.Row, 1).Value
            Call getDeliveryItem(strApprovalItemId, Ditem)
            Call putDeliveryAccount(rngPaste, postDeliveryItemToDeliveryAccount(Ditem))
            Set rngPaste = rngPaste.Offset(1, 0)
        End If
    Next
    '���F���̃`�F�b�N
    If lngCounter <= 0 Then makeApprovalList = "���F��K�v�Ƃ���o�ɂ�����܂���ł���": GoTo ending
    '�����f�[�^�L��
    Set rngPaste = shtPaste.Cells(getEndRow("a:z", shtPaste) + 1, DeliveryAccount_sum_COL)
    rngPaste.Value = getSumFunction(getFindRange(shtPaste, DeliveryAccount_sum_COL))
    Set rngPaste = rngPaste.Offset(0, -7).Resize(1, 7)
    rngPaste.Merge
    rngPaste = "��           �v"
    '�\��f�[�^�L��
    shtPaste.Range("a2") = dteClosingDate
    shtPaste.Range("b4") = Range("OFFICE_NAME")
    shtPaste.Range("d4") = BillType
    shtPaste.Range("c2").Value = "�ۍL�S�ݓX" & _
                                Range("STORE_NAME").Value & " �l �ǋ��ތ����@���F�肢"
    '��������ݒ�
    Set rngPaste = shtPaste.Range(shtPaste.Cells(DATA_START_ROW, 1), shtPaste.Cells(getEndRow("a:z", shtPaste), 8)).Offset(-1, 0)
    Set rngPaste = rngPaste.Resize(rngPaste.rows.Count + 1, rngPaste.Columns.Count)
    Call standerdPrintSetUp(shtPaste)
    Call PrintFormatBill(rngPaste)
    makeApprovalList = CStr(lngCounter)
    shtPaste.Visible = xlSheetVisible
    Call WindowFreezePanes(shtPaste, shtPaste.Range("a6"))
ending:
    Call SheetProtect("all")
    Set rngPaste = Nothing
    Set shtPaste = Nothing
    Set rngMy = Nothing
    Set shtMy = Nothing
End Function
Sub makeSalesList2(da() As DeliveryAccount, _
                   ta() As TenantAccaunts, _
                   ta_is_null As Boolean, _
                   settle_date As Date)
'���グ�ꗗ�\���쐬����
    Dim shtMy As Worksheet
    Dim rngMaruhiro As Range
    Dim rngTenant As Range
    Dim price As String
    Dim price_with_tax As String
    Dim cost As String
    Dim t_code As String
    Dim bill_type As String
    Dim accaunt() As TenantAccaunts
    Dim i As Long, j As Long, k As Long

    Set shtMy = Sheets("����ꗗ�\")
    Set rngMaruhiro = shtMy.Range("b6:h6")
    Set rngTenant = shtMy.Range("a9:h55")
    
'������
    Call SheetUnprotect(shtMy)
    rngMaruhiro.Value = ""
    rngTenant.Value = ""
    
'�]�L
    '�\��
    shtMy.Range("h3") = Range("OFFICE_NAME")
    shtMy.Range("a2") = getClosingdate(settle_date)
    '�S�ݓX��
    price = getMaruhiroTotalPrice(da)
    cost = getMaruhiroTotalCost(da)
    shtMy.Range("b6") = Range("STORE_NAME")
    shtMy.Range("d6") = price
    shtMy.Range("e6") = cost
    shtMy.Range("f6") = price - cost
    shtMy.Range("g6") = "�[�i�`�["
    '�e�i���g��
    ReDim accaunt(0)
    j = 0
    '�Ō�̃e�i���g�f�[�^��else�߂ŏ��������邽�߂̔z��g��
    ReDim Preserve ta(UBound(ta) + 1)
    t_code = ta(0).tenant_code & ta(0).bill_type
    For i = 0 To UBound(ta)
        If t_code Like ta(i).tenant_code & ta(i).bill_type Then
            Call addTenantAccauntData(ta(i), accaunt(), j)
        Else
            ReDim Preserve accaunt(j - 1)
    '�W�v���z�̎Z�o
            cost = CStr(SumOfCost(accaunt))
            Select Case accaunt(0).bill_type
            Case "������"
                price = ""
                price_with_tax = "����������"
            Case Else
                price = CStr(SumOfPrice(accaunt))
                price_with_tax = CStr(tax(price))
            End Select
            
            Call putSaleTA(accaunt(0), price, price_with_tax, cost, rngTenant)
            t_code = ta(i).tenant_code & ta(i).bill_type
            j = 0
            ReDim accaunt(j)
            Call addTenantAccauntData(ta(i), accaunt(), j)
        End If
    Next
    With shtMy
        .Visible = xlSheetVisible
        Call WindowFreezePanes(shtMy, .Range("a6"))
    End With
ending:
    Call SheetProtect
    Set shtMy = Nothing
    Set rngMaruhiro = Nothing
    Set rngTenant = Nothing
End Sub
Function SumOfPrice(accaunt() As TenantAccaunts) As Long
    Dim item As Variant
    Dim i As Long
    For i = 0 To UBound(accaunt)
        SumOfPrice = SumOfPrice + CLng(accaunt(i).price * accaunt(i).number)
    Next
End Function
Function SumOfCost(accaunt() As TenantAccaunts) As Long
    Dim i As Long
    For i = 0 To UBound(accaunt)
        SumOfCost = SumOfCost + CLng(accaunt(i).cost * accaunt(i).number)
    Next
End Function
Sub makeTenantAccaunts2(data() As TenantAccaunts, settle_date As Date)
'�e�i���g����������쐬����
    Dim shtMy As Worksheet
    Dim shtAccaunt As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim rngMy As Range
    Dim rngSum As Range
    Dim t_code As String
    Dim sum As SumList
    Dim accaunt() As TenantAccaunts
    Dim price As Long
    Dim rngSubtotal As Range
    Dim lngColumn As Long
    Const SubtotalRows As Long = 3
    
    Set shtMy = Sheets("���ύ�")
    Set shtAccaunt = Sheets("�e�i���g��������")
    
    ReDim accaunt(0)
    Call SheetUnprotect(shtAccaunt)
    Call clearList(shtAccaunt)
    Set rngMy = shtAccaunt.Cells(getEndRow("a:z", shtAccaunt) + 1, 1)
    Set rngSum = shtAccaunt.Cells(getEndRow("a:z", shtAccaunt) + 1, TenantAccaunts_sum_COL)
    t_code = data(0).tenant_code
    j = 0
    '�Ō�̃e�i���g�f�[�^��else�߂ŏ��������邽�߂̔z��g��
    ReDim Preserve data(UBound(data) + 1)
    For i = 0 To UBound(data)
        If t_code Like data(i).tenant_code Then
            Call addTenantAccauntData(data(i), accaunt(), j)
        Else
            '�e�i���g�f�[�^����A��������
            ReDim Preserve accaunt(j - 1)
            Call TenantAccauntSort(accaunt, LBound(accaunt), UBound(accaunt))
            Call putTenantAccaunts(accaunt, price, rngMy)
            Debug.Print accaunt(0).tenant_code & " : " & UBound(accaunt) + 1
            '���v��������
            Call putTenantAccauntsSubTotal(price, sum, rngMy)
            t_code = data(i).tenant_code
            j = 0
            ReDim accaunt(j)
            Call addTenantAccauntData(data(i), accaunt(), j)
        End If
    Next
    Call putTATotal(sum.price_without_tax, sum.tax, sum.price, rngMy)
    
    Set rngMy = shtAccaunt.Range(shtAccaunt.Cells(DATA_START_ROW, 1), _
                                 shtAccaunt.Cells(getEndRow("a:z", shtAccaunt), TenantAccaunts_sum_COL)).Offset(-1, 0)
    Set rngMy = rngMy.Resize(rngMy.rows.Count + 1, rngMy.Columns.Count)
    '�\����t��������
    With shtAccaunt
        .Range("a2").Value = getClosingdate(settle_date)
        .Range("b4").Value = Range("OFFICE_NAME")
    End With
    '����p�����ݒ�
    Call standerdPrintSetUp(shtAccaunt)
    Call PrintFormatBill(rngMy, SubtotalRows)
    '���v���C���쐬
    Set rngSubtotal = rngMy.Resize(1, TenantAccaunts_sum_COL)
    lngColumn = rngSubtotal.Count
    For i = 2 To rngMy.rows.Count
        If rngSubtotal.Cells(1).Value Like "" And Not rngSubtotal.Cells(lngColumn) = "" Then
            Call drowSubtotalLine(rngSubtotal.Resize(SubtotalRows, lngColumn))
            Set rngSubtotal = rngSubtotal.Offset(SubtotalRows, 0)
            i = i + SubtotalRows - 1
        Else
            Set rngSubtotal = rngSubtotal.Offset(1, 0)
        End If
    Next
    With shtAccaunt
        .Visible = xlSheetVisible
        Call WindowFreezePanes(shtAccaunt, .Range("a6"))
    End With
ending:
    Call SheetProtect
    Set rngSum = Nothing
    Set rngMy = Nothing
    Set shtAccaunt = Nothing
    Set shtMy = Nothing
End Sub
Sub MakeDeliveryAccount2(data() As DeliveryAccount, BillDate As Date)
'�ۍL����������쐬����
    Dim AccountIds As String
    Dim AccountId() As String
    Dim strBilltype() As String
    Dim varId As Variant
    Dim varType As Variant
    Dim curSum As Currency
    Dim TmpSettleItem As TmpSettleArticles
    Dim lngCount As Long
    
    Dim sht As Worksheet
    Dim rngMy As Range
    Dim i As Long
    
    Set sht = Worksheets("�ۍL��������")
    Call SheetUnprotect(sht)
    Call standerdPrintSetUp(sht)
    Call clearList(sht)
    Call FilterOff(sht)
    Set rngMy = sht.Cells(getEndRow("a:z", sht) + 1, DeliveryAccount_delivery_date_COL)
    For i = 0 To UBound(data)
        Call putDeliveryAccount(rngMy, data(i))
        Set rngMy = rngMy.Offset(1, 0)
    Next
    
    '�\��f�[�^
    sht.Range("a2").Value = BillDate
    sht.Range("b4").Value = Range("OFFICE_NAME").Value
    sht.Range("d4").Value = "�[�i�`�["
    sht.Range("c2").Value = "�ۍL�S�ݓX " & Range("STORE_NAME").Value & " �l"
    Set rngMy = sht.Cells(getEndRow("a:z", sht) + 1, DeliveryAccount_sum_COL)
    rngMy.Value = getSumFunction(getFindRange(sht, DeliveryAccount_sum_COL))
    Set rngMy = rngMy.Offset(0, -7).Resize(1, 7)
    rngMy.Merge
    rngMy.Value = "���v(�Ŕ���)"
'    MsgBox CStr(UBound(data) + 1) & " ���ł�"
    Set rngMy = sht.Range(sht.Cells(DATA_START_ROW, 1), sht.Cells(getEndRow("a:z", sht), 8)).Offset(-1, 0)
    Set rngMy = rngMy.Resize(rngMy.rows.Count + 1, rngMy.Columns.Count)
    '�����ݒ�
    Call PrintFormatBill(rngMy)
    '�\���ݒ�
    sht.Visible = xlSheetVisible
    Call WindowFreezePanes(sht, sht.Range("a6"))
ending:
    Call SheetProtect
    Set sht = Nothing
End Sub
Sub makeBillList2(da() As DeliveryAccount, _
                  ta() As TenantAccaunts, _
                  ta_is_null As Boolean, _
                  settle_date As Date)
'�������z�ꗗ���쐬����
    Dim shtMy As Worksheet
    Dim rngMaruhiro As Range
    Dim rngTenant As Range
    Dim i As Long, j As Long, k As Long
    Dim p_with_tax As Long
    Dim accaunt() As TenantAccaunts
    Dim t_code As String
    
    Set shtMy = Sheets("�������z�ꗗ�\")
    Set rngMaruhiro = shtMy.Range("b8:f8")
    Set rngTenant = shtMy.Range("a12:f57")
    
'������
    Call SheetUnprotect(shtMy)
    rngMaruhiro.Value = ""
    rngTenant.Value = ""
'�]�L
    '�\��
    shtMy.Range("e4") = "�s��r���T�[�r�X " & Range("OFFICE_NAME")
    shtMy.Range("a2") = getClosingdate(settle_date)
    '�S�ݓX��
    shtMy.Range("b8") = Range("STORE_NAME")
    shtMy.Range("e8") = getMaruhiroTotalPrice(da)
    '�e�i���g��
    ReDim accaunt(0)
    j = 0
    '�Ō�̃e�i���g�f�[�^��else�߂ŏ��������邽�߂̔z��g��
    ReDim Preserve ta(UBound(ta) + 1)
    t_code = ta(0).tenant_code & ta(0).bill_type
    For i = 0 To UBound(ta)
        If t_code Like ta(i).tenant_code & ta(i).bill_type Then
            Call addTenantAccauntData(ta(i), accaunt(), j)
        Else
            '�e�i���g�f�[�^����A�e�i���g�������z�v�Z
            ReDim Preserve accaunt(j - 1)
            Call TenantAccauntSort(accaunt, LBound(accaunt), UBound(accaunt))
            For k = 0 To UBound(accaunt)
                p_with_tax = p_with_tax + accaunt(k).sum
            Next
            p_with_tax = tax(p_with_tax)
            '���v��������
            Call putBillTA(accaunt(0), p_with_tax, rngTenant)
            '������
            p_with_tax = 0
            t_code = ta(i).tenant_code & ta(i).bill_type
            j = 0
            ReDim accaunt(j)
            Call addTenantAccauntData(ta(i), accaunt(), j)
        End If
    Next
    With shtMy
        .Visible = xlSheetVisible
        Call WindowFreezePanes(shtMy, .Range("a6"))
    End With
ending:
    Call SheetProtect
    Set shtMy = Nothing
    Set rngMaruhiro = Nothing
    Set rngTenant = Nothing
End Sub

