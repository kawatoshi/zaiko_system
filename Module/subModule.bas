Attribute VB_Name = "subModule"
Option Explicit
Function delRows(strRows As String, Optional shtDel As Worksheet) As String
'strRows�ɗ^����ꂽ�s���폜����
    Dim strRow() As String
    Dim i As Long, j As Long
    Dim varRow As Variant
    
    Call SheetUnprotect(shtDel)
    If shtDel Is Nothing Then Set shtDel = ActiveSheet
    strRow() = Split(strRows, " ")
    i = UBound(strRow())
    If i = -1 Then delRows = "nodata ERROR": GoTo ending
    i = 0
    For i = LBound(strRow) To UBound(strRow)
        j = 5 - Len(strRow(i))
        If j > 0 Then
            strRow(i) = Space(j) & strRow(i)
            strRow(i) = Replace(strRow(i), " ", "0")
        End If
    Next
    Call Selectionsort(strRow, LBound(strRow), UBound(strRow))
    For i = UBound(strRow) To LBound(strRow) Step -1
        If IsNumeric(CLng(strRow(i))) = True Then
            shtDel.rows(CLng(strRow(i))).Delete Shift:=xlShiftUp
            delRows = delRows & CStr(CLng(strRow(i))) & vbCrLf
        End If
    Next
    delRows = delRows & "Row Delete"
    Call SheetProtect("all")
    Call SheetProtect("select")
ending:
End Function
Sub Selectionsort(values() As String, _
                  ByVal min As Long, _
                  ByVal max As Long)
' Sort items in the values array with bounds min and max.
'���X�g��1000���x�̏ꍇ�Ɏg�p���邱��
  Dim i As Long
  Dim j As Long
  Dim smallest_value As String
  Dim smallest_j As Long

  For i = min To max - 1
    ' Find the smallest remaining value in entries
    ' i through num.
    smallest_value = values(i)
    smallest_j = i

    For j = i + 1 To max
      ' See if values(j) is smaller.
      If values(j) < smallest_value Then
        ' Save the new smallest value.
        smallest_value = values(j)
        smallest_j = j
      End If
    Next j

    If smallest_j <> i Then
      ' Swap items i and smallest_j.
      values(smallest_j) = values(i)
      values(i) = smallest_value
    End If

  Next i

End Sub
Function DuplicationMerge(strValue() As String) As String
'�z����\�[�g���ďd���𖳂������z��ɂ��ĕԂ�
    Dim i As Long, j As Long
    Dim lngStart As Long
    If UBound(strValue) = -1 Then GoTo ending
    Selectionsort strValue, LBound(strValue), UBound(strValue)
    j = LBound(strValue)
    lngStart = j + 1
    For i = lngStart To UBound(strValue)
        If strValue(i) <> strValue(j) Then
            j = j + 1
            strValue(j) = strValue(i)
        End If
    Next i
    ReDim Preserve strValue(j)
    DuplicationMerge = Join(strValue, " ")
ending:
End Function
Function DuplicationDraw(strValue() As String) As String
'�z����\�[�g���ďd�����Ă�����̂�z��ŕԂ�
    Dim i As Long, j As Long, k As Long
    Dim lngStart As Long
    Dim strtest As String
    Dim strI() As String
    
    If UBound(strValue) = -1 Then GoTo ending
    
    Selectionsort strValue, LBound(strValue), UBound(strValue)
    ReDim strI(UBound(strValue))
    j = LBound(strValue)
    lngStart = j + 1
    For i = lngStart To UBound(strValue)
        If strValue(i) <> strValue(j) Then
            j = j + 1
            strValue(j) = strValue(i)
        Else
            strI(k) = strValue(i)
            k = k + 1
        End If
    Next i
    If k = 0 Then ReDim strValue(0): GoTo ending
    ReDim Preserve strI(k - 1)
    Selectionsort strI, LBound(strI), UBound(strI)
    DuplicationMerge strI
    strValue = strI
    DuplicationDraw = Join(strValue, " ")
ending:
End Function

Function moveStockToDelivery(strItemId As String, dblNumber As Double, _
                             strCustomerId As String, strBilltype As String) As String
'���i���݌ɂ���o�ɂֈړ�������
    Dim strRow() As String
    Dim shtMy As Worksheet
    Dim shtStore As Worksheet
    Dim rngMy As Range
    Dim Sitem() As StockArticles
    Dim Ditem As DeliveryArticles
    Dim i As Long
    Dim dblSum As Double
    Dim dblNum As Double
    
    Set shtMy = ActiveWorkbook.Sheets("�o��")
    Set shtStore = ActiveWorkbook.Sheets("�݌�")
    Set rngMy = getFindRange(shtStore, StockArticles_item_id_COL)
    If chkSplit(getTableDatas(shtStore, strItemId, StockArticles_item_id_COL, StockArticles_id_COL), strRow()) = False Then _
        moveStockToDelivery = "�݌ɂ����݂��܂��� ERROR": GoTo ending
    ReDim Sitem(UBound(strRow))
    For i = 0 To UBound(strRow)
        Call getStockItem(CLng(strRow(i)), Sitem(i))
        dblSum = dblSum + CDbl(Sitem(i).number)
    Next
    If dblNumber > dblSum Then _
        moveStockToDelivery = "�݌ɂ�����܂��� ERROR": GoTo ending
    dblNum = dblNumber
    For i = 0 To UBound(strRow)
        Set rngMy = shtMy.Cells(getEndRow("a", shtMy) + 1, 1)
        dblNum = CDbl(Sitem(i).number) - dblNum
            If dblNum = 0 Then
                Debug.Print postStockToDeliveryItem(Sitem(i), Ditem, dblNum, strCustomerId, strBilltype)
                Debug.Print delStockItem(Sitem(i))
                Debug.Print putDeliveryItem(rngMy, Ditem)
                GoTo ending
            End If
            If dblNum > 0 Then
                Debug.Print postStockToDeliveryItem(Sitem(i), Ditem, dblNum, strCustomerId, strBilltype)
                
                Debug.Print putDeliveryItem(rngMy, Ditem)
                Set rngMy = shtStore.Cells(getKeyData(Sitem(i).id, _
                                            getFindRange(shtStore, StockArticles_id_COL), "row", xlWhole), _
                                            StockArticles_id_COL)
                Debug.Print putStockItem(rngMy, Sitem(i))
                GoTo ending
            End If
                Debug.Print postStockToDeliveryItem(Sitem(i), Ditem, dblNum, strCustomerId, strBilltype)
                Debug.Print delStockItem(Sitem(i))
                Debug.Print putDeliveryItem(rngMy, Ditem)
                dblNum = -1 * dblNum
    Next
ending:
    Set shtMy = Nothing
    Set rngMy = Nothing
End Function
Function moveStockToLost(strItemId As String, dblNumber As Double) As String
    Dim strRow() As String
    Dim shtMy As Worksheet
    Dim shtStore As Worksheet
    Dim rngMy As Range
    Dim Sitem() As StockArticles
    Dim LostItem As LostArticles
    Dim i As Long
    Dim dblSum As Double
    Dim dblNum As Double
    
    Set shtMy = ActiveWorkbook.Sheets("���X")
    Set shtStore = ActiveWorkbook.Sheets("�݌�")
    Set rngMy = getFindRange(shtStore, StockArticles_item_id_COL)
    If chkSplit(getTableDatas(shtStore, strItemId, StockArticles_item_id_COL, StockArticles_id_COL), strRow()) = False Then GoTo ending
    '�݌Ɏ擾
    ReDim Sitem(UBound(strRow))
    For i = 0 To UBound(strRow)
        Call getStockItem(CLng(strRow(i)), Sitem(i))
        dblSum = dblSum + CDbl(Sitem(i).number)
    Next
    '�݌ɐ��̊m�F
    If dblNumber > dblSum Then _
        MsgBox ("�݌ɂ�����܂���"): moveStockToLost = "ERROR": GoTo ending
    dblNum = dblNumber
    '�]�L����
    For i = 0 To UBound(strRow)
        Set rngMy = shtMy.Cells(getEndRow("a", shtMy) + 1, 1)
        dblNum = CDbl(Sitem(i).number) - dblNum
            If dblNum = 0 Then
                Debug.Print postStockToLost(Sitem(i), LostItem, dblNum)
                Debug.Print delStockItem(Sitem(i))
                Debug.Print putLostItem(rngMy, LostItem)
                GoTo ending
                moveStockToLost = getHinmoku(strItemId)

            End If
            If dblNum > 0 Then
                Debug.Print postStockToLost(Sitem(i), LostItem, dblNum)
                
                Debug.Print putLostItem(rngMy, LostItem)
                Set rngMy = shtStore.Cells(getKeyData(Sitem(i).id, _
                                            getFindRange(shtStore, StockArticles_id_COL), "row", xlWhole), _
                                            StockArticles_id_COL)
                Debug.Print putStockItem(rngMy, Sitem(i))
                moveStockToLost = getHinmoku(strItemId)
                GoTo ending
            End If
                Debug.Print postStockToLost(Sitem(i), LostItem, dblNum)
                Debug.Print delStockItem(Sitem(i))
                Debug.Print putLostItem(rngMy, LostItem)
                dblNum = -1 * dblNum
    Next
    moveStockToLost = getHinmoku(strItemId)
ending:
    Set shtMy = Nothing
    Set rngMy = Nothing
End Function

Function delStockItem(stockItem As StockArticles) As String
'�݌ɂ���������
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strRows As String
    Set shtMy = ActiveWorkbook.Sheets("�݌�")
    Set rngFind = getFindRange(shtMy, StockArticles_id_COL)
    Call SheetUnprotect(shtMy)
    
    strRows = getKeyData(stockItem.id, rngFind, "row", xlWhole)
    shtMy.rows(CLng(strRows)).Delete Shift:=xlUp
    delStockItem = stockItem.id & " is Dell!"
    
    Call SheetProtect
    Set shtMy = Nothing
    Set rngFind = Nothing
End Function
Sub Dup(strState As String, strLists As String)
'strLists�ɗ^����ꂽ���X�g��strState�ɗ^����ꂽ
'���X�g���������ďd���̂Ȃ����X�g��strLists�ɖ߂�
'strState�ɒl���Ȃ��Ƃ��ɂ͏������s��Ȃ�
    Dim strList() As String
    If strState = "" Then GoTo ending
    strLists = strLists & " " & strState
    strLists = Trim(strLists)
    strList() = Split(strLists, " ")
    Call DuplicationDraw(strList())
    strLists = Join(strList(), " ")
ending:
End Sub

Function returnDeleveryToStock(strDeliveryItemId As String, dblNumber As Double) As String
'�o�ɂ������i�̕ԕi�����B�o�ɂ���݌ɂւ̃f�[�^�ړ�
'�����ɕύX�������L�^����
    Dim shtReturn As Worksheet
    Dim shtStock As Worksheet
    Dim shtDelivery As Worksheet
    Dim rngMy As Range
    Dim Sitem As StockArticles
    Dim DeliveryItem As DeliveryArticles
    Dim Ritem As ReturnArticles
    Dim dblNum As Double
    Dim strState As String
    
    Set shtDelivery = ActiveWorkbook.Sheets("�o��")
    Set shtReturn = ActiveWorkbook.Sheets("�ԕi����")
    Set shtStock = ActiveWorkbook.Sheets("�݌�")
'    Set rngMy = getFindRange(shtDelivery, DeliveryArticles_id_COL)
    '�o�ɕi�擾
    strState = getDeliveryItem(strDeliveryItemId, DeliveryItem)
    If strState Like "*ERROR" Then _
        returnDeleveryToStock = "�o�ɃG���[" & vbCrLf & strState & " ERROR": GoTo ending
    '�ԕi���̊m�F
    If chkNumber(DeliveryItem.number) Like "*ERROR" _
        Then returnDeleveryToStock = "���l����͂��Ă������� ERROR": GoTo ending
    '�]�L����
    Set rngMy = shtStock.Cells(getEndRow("a", shtStock) + 1, 1)
    dblNum = CDbl(DeliveryItem.number) - dblNum
    strState = postDeliveryToStock(DeliveryItem, Sitem, dblNumber)
    If strState Like "*ERROR" Then _
        returnDeleveryToStock = strState: GoTo ending
    '�o�ɏ�������
    If CDbl(DeliveryItem.number) = 0 Then
        Debug.Print delDeliveryItem(DeliveryItem)
    Else
        strState = getKeyData(DeliveryItem.id, _
                              getFindRange(shtDelivery, DeliveryArticles_id_COL), _
                              "row", xlWhole)
        Set rngMy = shtDelivery.Cells(strState, DeliveryArticles_id_COL)
        Debug.Print putDeliveryItem(rngMy, DeliveryItem)
    End If
    '�݌ɏ�������
    strState = getKeyData(Sitem.id, _
                          getFindRange(shtStock, StockArticles_id_COL), _
                          "row", xlWhole)
    If strState = "" Then
        strState = CStr(getEndRow("a", shtStock) + 1)
    End If
    Set rngMy = shtStock.Cells(strState, StockArticles_id_COL)
    Debug.Print putStockItem(rngMy, Sitem)
    '�ԕi��������
    Ritem.id = DeliveryItem.id
    Ritem.buy_article_id = DeliveryItem.buy_article_id
    Ritem.stock_article_id = DeliveryItem.stock_article_id
    Ritem.item_id = DeliveryItem.item_id
    Ritem.customer_id = DeliveryItem.customer_id
    Ritem.number = CStr(dblNumber)
    Ritem.cost = DeliveryItem.cost
    Ritem.item_price = DeliveryItem.item_price
    Ritem.return_date = Now()
    strState = CStr(getEndRow("a", shtReturn) + 1)
    Set rngMy = shtReturn.Cells(strState, ReturnArticles_id_COL)
    Debug.Print putReturnItem(rngMy, Ritem)
    returnDeleveryToStock = "OK"
ending:
    Set shtReturn = Nothing
    Set shtStock = Nothing
    Set rngMy = Nothing
End Function
Function delDeliveryItem(Ditem As DeliveryArticles) As String
'�݌ɂ���������
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strRows As String
    Set shtMy = ActiveWorkbook.Sheets("�o��")
    Set rngFind = getFindRange(shtMy, DeliveryArticles_id_COL)
    Call SheetUnprotect(shtMy)
    
    strRows = getKeyData(Ditem.id, rngFind, "row", xlWhole)
    shtMy.rows(CLng(strRows)).Delete Shift:=xlUp
    delDeliveryItem = Ditem.id & " is Dell!"
    Call SheetProtect("select")
    
    Set shtMy = Nothing
    Set rngFind = Nothing
End Function
Sub clearList(shtClear As Worksheet, _
                      Optional lngSrow As Long = DATA_START_ROW)
'���X�g����������
    Dim lngEndRow As Long
    Dim rngMy As Range
    Call SheetUnprotect(shtClear)
    
    lngEndRow = getEndRow("a:z", shtClear)
    If lngEndRow < DATA_START_ROW Then lngEndRow = DATA_START_ROW
    Set rngMy = shtClear.rows(CStr(lngSrow) & ":" & CStr(lngEndRow))
    Call clearBorder(rngMy)
    rngMy.Delete Shift:=xlUp
    
    Set rngMy = Nothing
End Sub
Sub clearBorder(rngMy As Range)
'�{�[�_�[����������
On Error GoTo HandleErr
    Dim brdMy As Borders
    Dim i As Long
    Set brdMy = rngMy.Borders
    
    For i = xlDiagonalDown To xlInsideHorizontal
        brdMy(i).LineStyle = xlNone
    Next
    Set brdMy = Nothing
    Exit Sub
HandleErr:
    Debug.Print Err.number & ": " & Err.Description & " subModule.clearBorder"
    Resume Next
End Sub
Function moveDeliveryToSettleItem() As String
'�o�ɂ��猈�ςɃf�[�^���ړ�����
'���ϓ��ȑO�̏o�Ƀf�[�^�݂̂��ړ�����
    Dim rngMy As Range
    Dim shtDArticles As Worksheet
    Dim shtSettleArticles As Worksheet
    Dim Ditem() As DeliveryArticles
    Dim SettleItem() As SettleArticles
    Dim varId As Variant
    Dim lngCount As Long
    Dim i As Long
    
    Set shtSettleArticles = Sheets("���ύ�")
    Set shtDArticles = Sheets("�o��")
    Set rngMy = getFindRange(shtDArticles, DeliveryArticles_id_COL)
    '�o�Ƀf�[�^�擾
    Application.ScreenUpdating = False
    If rngMy Is Nothing Then moveDeliveryToSettleItem = "�f�[�^������܂���": GoTo ending
    lngCount = rngMy.Count
    ReDim Ditem(lngCount)
    For i = 1 To lngCount
        Call getDeliveryItem(rngMy.Cells(i), Ditem(i))
    Next
    '�f�[�^�ړ�
    Call postDeliveryToSettleItem(Ditem(), SettleItem(), getClosingdate)
    '�ړ��f�[�^�̏�������
    For i = 1 To UBound(SettleItem)
        Set rngMy = shtSettleArticles.Cells(getEndRow("a", shtSettleArticles) + 1, 1)
        Call putSettleItem(rngMy, SettleItem(i))
    Next
    moveDeliveryToSettleItem = i & " ���̃f�[�^������܂�"
    '�ړ��ς݃f�[�^�̏���
    Call delMovedDeliveryItem(Ditem())
ending:
    Application.ScreenUpdating = True
    Set shtSettleArticles = Nothing
    Set shtDArticles = Nothing
    Set rngMy = Nothing
End Function

Function delMovedDeliveryItem(Ditem() As DeliveryArticles) As String
'�Y������f�[�^����������
    Dim i As Long
    Dim strRow As String
    Dim strState As String
    Dim shtD As Worksheet
    
    delMovedDeliveryItem = "delMovedDeliveryItem NG"
    Set shtD = Sheets("�o��")
    For i = 1 To UBound(Ditem)
        strRow = getKeyData(Ditem(i).id, getFindRange(shtD, DeliveryArticles_id_COL), "row", xlWhole)
        If Not strRow = "" Then _
            Call delRows(strRow, shtD)
    Next
    delMovedDeliveryItem = "delMovedDeliveryItem OK"
ending:
    Set shtD = Nothing
End Function

Function CopySettleItemToTmpSettleItem(strBillIds As String) As String
'���ύς݃V�[�g���疢�����f�[�^��TMP���σV�[�g�֓]�L����
    Dim shtTmp As Worksheet
    Dim rngMy As Range
    Dim strId() As String
    Dim varValue As Variant
    Dim SettleItem As SettleArticles
    
    Set shtTmp = ActiveWorkbook.Sheets("Tmp����")
    Call clearList(shtTmp)
    If chkSplit(strBillIds, strId) = False Then _
        CopySettleItemToTmpSettleItem = "�������̃f�[�^�͂���܂���ł���": GoTo ending
    Set rngMy = shtTmp.Cells(getEndRow("a", shtTmp) + 1, TmpSettleArticles_id_COL)
    For Each varValue In strId
        Call getSettleItem(CStr(varValue), SettleItem)
        Call putTmpSettleItem(rngMy, postTmpSettleItem(SettleItem))
        Set rngMy = rngMy.Offset(1, 0)
    Next
    CopySettleItemToTmpSettleItem = CStr(UBound(strId) + 1) & " ���̃f�[�^������܂���"
ending:
    Set shtTmp = Nothing
    Set rngMy = Nothing
End Function

Sub FilterOff(shtOff As Worksheet)
'shtOff�̃I�[�g�t�B���^�[����������
    With shtOff
        If .FilterMode Then
            On Error Resume Next
            .ShowAllData
            On Error GoTo 0
        End If
    End With
End Sub

Sub PrintFormatBill(rngFormat As Range, Optional footerRows As Long = 1)
'������ׂ̏�����ݒ肷��
    Dim rngMy As Range
    Dim shtMy As Worksheet
    Dim lngRows As Long
    Dim lngColumns As Long
    
    
    Set shtMy = Sheets(rngFormat.Parent.name)
    rngFormat.RowHeight = 15
    Call InsertPageSpace(rngFormat, footerRows)
    '�S��
    With rngFormat.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngFormat.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngFormat.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngFormat.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngFormat.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With rngFormat.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    '�擪
    lngColumns = rngFormat.Columns.Count
    Set rngMy = rngFormat.Cells(1).Resize(1, lngColumns)
    With rngMy.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    '��
    Set rngMy = rngFormat.Cells(rngFormat.rows.Count, 1).Offset(1 - footerRows, 0).Resize(footerRows, lngColumns)
    With rngMy.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Set rngMy = Nothing
    Set shtMy = Nothing
End Sub
Sub standerdPrintSetUp(shtPrint As Worksheet)
'�v�����g�̕W���ݒ���s��
    With shtPrint.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "&P/&N"
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.53)
        .TopMargin = Application.InchesToPoints(0.8)
        .BottomMargin = Application.InchesToPoints(0.94)
        .HeaderMargin = Application.InchesToPoints(0.512)
        .FooterMargin = Application.InchesToPoints(0.46)
        .PrintHeadings = False
        .PrintGridlines = False
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
    End With

End Sub
Sub InsertPageSpace(rngFormat As Range, Optional footerRows As Long = 1)
'rngFormat�̍ŏI�s���y�[�W�̍Ō�ɂȂ�悤�ɋ󔒍s��}������
    Dim rngMy As Range
    Dim shtMy As Worksheet
    Dim lngPageRow As Long
    Dim lngErow As Long
    Dim lngPages As Long
    Dim i As Long
    Dim lngRowCount As Long
    
    Set shtMy = Sheets(rngFormat.Parent.name)
    With shtMy
        .Visible = xlSheetVisible
        .Activate
        .Cells(rngFormat.End(xlDown).Row, 1).Select
        lngPages = .HPageBreaks.Count
    End With
    If Not lngPages = 0 Then
        shtMy.Activate
'        If rngFormat.End(xlDown).Row = shtMy.HPageBreaks(lngPages).Location.Row - 1 Then GoTo ending
    End If
    lngRowCount = rngFormat.rows.Count
    Set rngMy = rngFormat.Cells(lngRowCount - footerRows + 1, 1).Resize(10, 1)
    lngRowCount = 0
    Do Until lngPages + 1 = i
        lngRowCount = lngRowCount + 1
        rngMy.EntireRow.Insert
        shtMy.rows(rngMy.Row).Select
        i = shtMy.HPageBreaks.Count
        If lngRowCount > 20 Then Exit Do
    Loop
'##�y�[�W���\������Ȃ��ƍŏI�y�[�W�̍s���擾����Ȃ����߂�2�s�ǉ�##
    shtMy.Activate
    shtMy.rows(rngMy.Row).Select
'###################################################################
    lngPages = shtMy.HPageBreaks.Count
    lngPageRow = shtMy.HPageBreaks(lngPages).Location.Row
    If getEndRow("a:z", shtMy) = lngPageRow - 1 Then GoTo ending
    lngPages = shtMy.HPageBreaks.Count
    lngErow = getEndRow("a:z", shtMy)
    lngPageRow = shtMy.HPageBreaks(lngPages).Location.Row
    shtMy.rows(lngErow - footerRows - (lngErow - lngPageRow) & ":" & lngErow - footerRows).Delete
    With shtMy
        Set rngFormat = .Range(rngFormat.Cells(1).address & ":" & .Cells(getEndRow("a:z", shtMy), rngFormat.Columns.Count).address)
    End With
ending:
    Set rngMy = Nothing
    Set shtMy = Nothing
End Sub
Function PrintPage(strState As String, shtMy As Worksheet) As String
'�V�[�g�ɃG���[���Ȃ���Έ������
    Application.Wait (Now + TimeValue("00:00:03"))
    If Not CLng(Val(strState)) = 0 Then
        shtMy.PrintOut
    Else
        MsgBox (strState): GoTo ending
    End If
ending:

End Function


Sub drowSubtotalLine(rngMy As Range)
'���v�̃��C��������

    With rngMy.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngMy.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

End Sub

Function SpinUpNum(strNum As String) As String
'�X�s���{�^���Ő��𑝂₷�B
    Dim dblNum As Double
    If strNum = "" Then strNum = "0"
    On Error Resume Next
        dblNum = CDbl(strNum)
        If Err.number <> 0 Then SpinUpNum = "": GoTo ending
    On Error GoTo 0
    dblNum = dblNum + 1
    If dblNum < 0 Then SpinUpNum = "": GoTo ending
    
    SpinUpNum = CStr(dblNum)
ending:

End Function
Function SpinDownNum(strNum As String) As String
'�X�s���{�^���Ő������炷
    Dim dblNum As Double
    On Error Resume Next
        dblNum = CDbl(strNum)
        If Err.number <> 0 Then SpinDownNum = "": GoTo ending
    On Error GoTo 0
    dblNum = dblNum - 1
    If dblNum < 0 Then SpinDownNum = "": GoTo ending
    
    SpinDownNum = CStr(dblNum)
ending:
End Function
Function getVer() As String
'    getVer = ThisWorkbook.name
    getVer = strVer
End Function
Sub SheetProtect(Optional mode As String = "all")
'�V�[�g���v���e�N�g����
    Dim i As Long
    Dim varsheetname As Variant
    Dim bolMode As Boolean
    Dim dblApVer As Double
    
    dblApVer = Application.Version
    Select Case mode
        Case "all"
            varsheetname = Range("ALL_PROTECT")
            bolMode = False
        Case "select"
            varsheetname = Range("SELECT_PROTECT")
            bolMode = True
    End Select
    If dblApVer <= 9 Then
        For i = 1 To UBound(varsheetname)
            Sheets(varsheetname(i, 1)).Protect Password:="kawakita", _
                                               UserInterfaceOnly:=True
        Next
    Else
        For i = 1 To UBound(varsheetname)
            Sheets(varsheetname(i, 1)).Protect Password:="kawakita", _
                                               UserInterfaceOnly:=True, _
                                               AllowFiltering:=bolMode
        Next
    End If
End Sub
Function SheetUnvisible() As String
'�V�[�g��S�ĕs���ɂ���
    Dim i As Long
    Dim varsheetname As Variant
    
    If ThisWorkbook.Path Like "*honbu*" Then
        Exit Function
    End If
    varsheetname = Range("ALL_PROTECT")
    For i = 1 To UBound(varsheetname)
        Sheets(varsheetname(i, 1)).Visible = False
    Next
    varsheetname = Range("SELECT_PROTECT")
        For i = 1 To UBound(varsheetname)
        Sheets(varsheetname(i, 1)).Visible = False
    Next

End Function
Sub SheetUnprotect(shtUnprotect As Worksheet)
'�V�[�g���A���v���e�N�g����
    shtUnprotect.Unprotect "kawakita"
End Sub

Sub initList(strList() As String, strListName As String)
'�e��z��̏����l��Ԃ�
    Dim rngUniq As Range
    Dim shtMy As Worksheet
    
    Select Case strListName
        Case "item_name"
            Set shtMy = ActiveWorkbook.Sheets("�i��")
            Set rngUniq = getFindRange(shtMy, Articles_name_COL)
        Case "item_product_number"
            Set shtMy = ActiveWorkbook.Sheets("�i��")
            Set rngUniq = getFindRange(shtMy, Articles_product_number_COL)
        Case "item_maker_call_name"
            Set shtMy = ActiveWorkbook.Sheets("���[�J�[")
            Set rngUniq = getFindRange(shtMy, Makers_call_name_COL)
        Case "customer_site"
            Set shtMy = ActiveWorkbook.Sheets("�����")
            Set rngUniq = getFindRange(shtMy, Customers_site_COL)
        Case "customer_floor"
            Set shtMy = ActiveWorkbook.Sheets("�����")
            Set rngUniq = getFindRange(shtMy, Customers_floor_COL)
        Case "customer_place"
            Set shtMy = ActiveWorkbook.Sheets("�����")
            Set rngUniq = getFindRange(shtMy, Customers_place_COL)
        Case Else
            ReDim strList(0)
            GoTo ending
    End Select
    
    getArray rngUniq, strList
    DuplicationMerge strList()
    Call Selectionsort(strList, LBound(strList), UBound(strList))

ending:
    Set rngUniq = Nothing
    Set shtMy = Nothing
End Sub
Sub ItemProductNumChenge(strName As String, _
                         strMaker As String, _
                         strList() As String)
'�i���ƃ��[�J�[�Ăі�����i�ԃ��X�g���X�V����
    Dim strProductNumOnName As String
    Dim strProductNumOnMaker As String
    Dim strStockItemIDs As String
    Dim strProductNumOnStok As String
    Dim shtMy As Worksheet
    
    Set shtMy = Sheets("�i��")
    strProductNumOnName = getTableDatas(shtMy, strName, _
                                        Articles_name_COL, Articles_product_number_COL)
    strProductNumOnMaker = getItemProductNumFromItemIds(getItemIdsFromMakerCallName(strMaker))
    If strProductNumOnName = "" And strProductNumOnMaker = "" Then
        Call initList(strList, "item_product_number")
        GoTo ending
    End If
    strList = Split(Trim(strProductNumOnName & " " & strProductNumOnMaker), " ")
    If strProductNumOnName = "" Or strProductNumOnMaker = "" Then
        Call DuplicationMerge(strList)
    Else
        Call DuplicationDraw(strList)
    End If
ending:
    Set shtMy = Nothing
End Sub
Sub ItemMakerChenge(strName As String, strProductNum As String, _
                    strList() As String)
'�i���ƌ^�Ԃ��烁�[�J�[���X�g���X�V����
    Dim strMakerNamesOnName As String
    Dim strMakerNamesOnProduct As String
    
    strMakerNamesOnName = getMakerCallNamesFromItemName(strName)
    strMakerNamesOnProduct = getMakerCallNameFromItemProductNum(strProductNum)
    If strMakerNamesOnName = "" And strMakerNamesOnProduct = "" Then
        Call initList(strList(), "item_maker_call_name")
        GoTo ending
    End If
    strList = Split(Trim(strMakerNamesOnName & " " & strMakerNamesOnProduct), " ")
    If strMakerNamesOnName = "" Or strMakerNamesOnProduct = "" Then
        Call DuplicationMerge(strList)
        GoTo ending
    Else
        Call DuplicationDraw(strList)
    End If
ending:
End Sub
Sub itemNameChenge(strMaker As String, strProductNum As String, strList() As String)
'���[�J�[���ƌ^�Ԃ���i�����X�g���X�V����
    Dim strNamesOnMaker As String
    Dim strNamesOnProductNum As String
    Dim shtItem As Worksheet
    Set shtItem = ActiveWorkbook.Sheets("�i��")
    
    '���[�J�[�Ăі�����i�����X�g���擾
    strNamesOnMaker = getTableDatas(shtItem, getMakerIdFromMakerCallName(strMaker), _
                                    Articles_maker_id_COL, Articles_name_COL)
    '�^�Ԃ���i�����X�g���擾
    strNamesOnProductNum = getTableDatas(shtItem, strProductNum, _
                                         Articles_product_number_COL, Articles_name_COL)
    strList = Split(Trim(strNamesOnMaker & " " & strNamesOnProductNum), " ")
    Call DuplicationMerge(strList)
ending:
    Set shtItem = Nothing
End Sub
Sub initItemProductList(strProduct() As String)
'�^�Ԕz��̏����l��Ԃ�
    Dim rngUniq As Range
    Dim shtMy As Worksheet
    
    Set shtMy = ActiveWorkbook.Sheets("�i��")
    Set rngUniq = getFindRange(shtMy, Articles_product_number_COL)
    getArray rngUniq, strProduct
    DuplicationMerge strProduct()
    Call Selectionsort(strProduct, LBound(strProduct), UBound(strProduct))
    Set rngUniq = Nothing
    Set shtMy = Nothing
End Sub
Sub initItemNameList(strName() As String)
'�i���z��̏����l��Ԃ�
    Dim rngUniq As Range
    Dim shtMy As Worksheet
    
    Set shtMy = ActiveWorkbook.Sheets("�i��")
    Set rngUniq = getFindRange(shtMy, Articles_name_COL)
    getArray rngUniq, strName
    DuplicationMerge strName()
    Call Selectionsort(strName, LBound(strName), UBound(strName))
    Set rngUniq = Nothing
    Set shtMy = Nothing

End Sub
Sub itemNameListChenge(itemName As String, ItemProductNum As String, _
                       makerCallName As String, strMakerCallNameList() As String, _
                       strItemProductNumberList() As String)
'�i���I�����X�V���ꂽ�Ƃ��̊e���ڃ��X�g���쐬����
    ReDim strMakerCallNameList(0)
    ReDim strItemProductNumberList(0)
    '�G���[�`�F�b�N
    If Not chkInputStr(makerCallName) Like "ok" Then GoTo ending
    If Not chkInputStr(ItemProductNum) Like "ok" Then GoTo ending
    If Not chkInputStr(itemName) Like "ok" Then GoTo ending
    '���[�J�[���X�g�X�V
    Call ItemMakerChenge(itemName, ItemProductNum, strMakerCallNameList)
    '�^�ԁE�i�ԍX�V
    Call ItemProductNumChenge(itemName, makerCallName, strItemProductNumberList)
ending:
End Sub
Function MakerListChenge(makerCallName As String, ItemProductNum As String, itemName As String, _
                    strItemNameList() As String, strItemProductNumList() As String) As String
'���[�J�[���I�����X�V���ꂽ�Ƃ��̊e���ڃ��X�g���쐬����
    Dim strId As String
    ReDim strItemNameList(0)
    ReDim strItemProductNumList(0)
    '�G���[�`�F�b�N
    If Not chkInputStr(makerCallName) Like "ok" Then GoTo ending
    If Not chkInputStr(ItemProductNum) Like "ok" Then GoTo ending
    If Not chkInputStr(itemName) Like "ok" Then GoTo ending
    '�i�����X�g�X�V
    Call itemNameChenge(makerCallName, ItemProductNum, strItemNameList)
    '�^�ԁE�i�ԍX�V
    Call ItemProductNumChenge(itemName, makerCallName, strItemProductNumList)
    
    strId = getItemIdFromMakerCallNameAndProductNumber(makerCallName, ItemProductNum)
    If Not strId = "" Then _
        MakerListChenge = strId
ending:
End Function
Function ProductListChenge(makerCallName As String, ProductNumber As String, itemName As String, _
                      strMakerCallNameList() As String, strItemNameList() As String) As String
'�^�ԁE�i�Ԃ��X�V���ꂽ�Ƃ��̊e���ڃ��X�g���쐬����
    Dim strId As String
    ReDim strMakerCallNameList(0)
    ReDim strItemNameList(0)
    '�G���[�`�F�b�N
    If Not chkInputStr(makerCallName) Like "ok" Then GoTo ending
    If Not chkInputStr(ProductNumber) Like "ok" Then GoTo ending
    If Not chkInputStr(itemName) Like "ok" Then GoTo ending
    '�i�����X�g�X�V
    Call itemNameChenge(makerCallName, ProductNumber, strItemNameList)
    '���[�J�[���X�g�X�V
    Call ItemMakerChenge(itemName, ProductNumber, strMakerCallNameList)
    
    strId = getItemIdFromMakerCallNameAndProductNumber(makerCallName, ProductNumber)
    If Not strId = "" Then _
        ProductListChenge = strId
ending:
End Function
Function ProductListChengeNew(makerCallName As String, ProductNumber As String, itemName As String, _
                                MakerCallNameList() As String, ItemNameList() As String, _
                                Optional makerCallNameChenge As Boolean = True) As String
'�^�ԁE�i�Ԃ��X�V���ꂽ�Ƃ��̊e���ڃ��X�g���쐬����
'���̃��X�g�X�V��product number���X�V���ꂽ�Ƃ��݂̂Ɏg�p�����function�Ȃ̂�

    Dim shtMy As Worksheet
    Dim shtMaker As Worksheet
    Dim strId(2) As String
    Dim strMakerId As String
    Dim splitId() As String
    Dim mergedItemIDs As String
    
    Set shtMy = Worksheets("�i��")
    Set shtMaker = Worksheets("���[�J�[")
    '���[�J�Ăі������item_id�擾�y�у��X�g�擾
    Select Case makerCallName
    Case ""
        Debug.Print "Maker count: No check"
    Case Else
        strMakerId = getMakerIdFromMakerCallName(makerCallName)
        strId(0) = getTableDatas(shtMy, strMakerId, Articles_maker_id_COL, Articles_id_COL)
        Debug.Print "Maker count: " & CStr(UBound(Split(strId(0))) + 1)
        MakerCallNameList = Split(getTableDatas(shtMaker, strMakerId, _
                                                Makers_id_COL, Makers_call_name_COL), " ")
    End Select
    '�^�Ԃ����item_id�擾
    Select Case ProductNumber
    Case ""
        'productNum�����݂��Ȃ��ꍇ
        makerCallNameChenge = False
        Debug.Print "ProductNum count: No check"
    Case Else
        'ProductNum�����݂���ꍇ�^�Ԃ���Y��Item id�̎擾
        strId(1) = getTableDatas(shtMy, ProductNumber, Articles_product_number_COL, Articles_id_COL)
'        splitId = Split(strId(0) & " " & strId(1), " ")
        Debug.Print "ProductNum count: " & CStr(UBound(Split(strId(1))) + 1)
        If UBound(Split(strId(1))) = 0 Then
            '�^�Ԃ���1��Item id���擾�ł����ꍇ
            ProductListChengeNew = Trim(strId(1))
            splitId = Split(strId(1))
            Exit Function
        Else
            'Item id�������擾���ꂽ���A�擾����Ȃ������ꍇ
            If strId(0) = "" Then
                'MakerCallName�����݂��Ȃ��ꍇ
                makerCallNameChenge = False
                splitId = Split(strId(1))
                Debug.Print "MakerCallName not found"
                If UBound(splitId) > 0 Then
                    Debug.Print "same product count: " & CStr(UBound(splitId) + 1)
                    Call getMakerCallNameFromMakerIds(getTableDatas(shtMy, _
                                                                    Join(splitId), _
                                                                    Articles_id_COL, _
                                                                    Articles_maker_id_COL), _
                                                      MakerCallNameList)
                makerCallNameChenge = True
                Else
                    Debug.Print "no same product count: " & CStr(UBound(splitId))
                End If
            Else
                'MakerCallName�����݂����ꍇ
                splitId = Split(strId(0) & " " & strId(1), " ")
                mergedItemIDs = DuplicationDraw(splitId)
                Debug.Print "merge Item count: " & CStr(UBound(splitId) + 1)
                makerCallNameChenge = True
                Call getMakerCallNameFromMakerIds(getTableDatas(shtMy, _
                                                                mergedItemIDs, _
                                                                Articles_id_COL, _
                                                                Articles_maker_id_COL), _
                                                  MakerCallNameList)
            End If
        End If
    End Select
    splitId = Split(strId(0) & " " & strId(1), " ")
    mergedItemIDs = DuplicationDraw(splitId)
    If UBound(splitId) = 0 Then
        ProductListChengeNew = mergedItemIDs
    End If
    strMakerId = DuplicationDraw(splitId)
End Function
Function CopyFile(bokMy As Workbook, strPasteFolder As String) As String
'�t�@�C���������t�H���_�փR�s�[����
    Dim strPathMy As String
    Dim strBokName As String
    
    If Not chkFolder(strPasteFolder) Like "folder" Then
        CopyFile = "Paste Folder ERROR": GoTo ending
    End If
    With bokMy
        strPathMy = .Path
        strBokName = .name
    End With
    On Error Resume Next
    bokMy.SaveCopyAs Filename:=strPasteFolder & "\" & strBokName
    If Err.number <> 0 Then
        CopyFile = "FileCopy ERROR" & vbCrLf & _
                   "error.num=" & Err.number & _
                   bokMy.name & " : " & strPasteFolder
        GoTo ending
    End If
    On Error GoTo 0
    CopyFile = "ok"
ending:
End Function
Sub WindowFreezePanes(shtMy As Worksheet, rngFreeze As Range)
'�g���Œ肷��
    Application.ScreenUpdating = False
    On Error Resume Next
    With shtMy
        .Activate
        With ActiveWindow
            If .FreezePanes = True Then .FreezePanes = False
        End With
    Application.ScreenUpdating = True
        .Range("a1").Activate
    End With
    rngFreeze.Select
    If Err.number <> 0 Then GoTo ending
    On Error GoTo 0
    ActiveWindow.FreezePanes = True
ending:
    Application.ScreenUpdating = True
End Sub

Sub HonbuView()
'�{���ł̈ꗗ�\���p
    Dim varShtName As Variant
    Dim varNAME As Variant
    
    varShtName = Array("�ۍL��������", "�e�i���g��������", "�������z�ꗗ�\", "����ꗗ�\")
    For Each varNAME In varShtName
        Sheets(varNAME).Visible = xlSheetVisible
    Next
End Sub
Sub sendJanRegistMail(Mailto As String, JANCODE As String, itemID As String)
'JAN�o�^�\����sendAddress���ɑ��t����
    Dim item As Articles
    Dim maker As makers
    Dim Trader As Traders
    Dim basp As Object
    Dim NS As String
    Dim Mailfrom As String
    Dim Subject As String
    Dim Body As String
    Dim msg As String
    
    Call getItem(itemID, item)
    Call getMaker(item.maker_id, maker)
    Call getTrader(item.trader_id, Trader)
    Set basp = CreateObject("basp21")
    NS = "mail.fbs"
    Mailfrom = "syoumouhin"
    Subject = "JAN�o�^�\��"
    Body = "�\���c�Ə�:" & Chr(13) & _
           Chr(9) & Range("OFFICE_NAME").Value & Chr(13) & _
           "JANCODE:" & Chr(13) & _
           Chr(9) & JANCODE & Chr(13) & _
           "id:" & Chr(13) & _
           Chr(9) & item.id & Chr(13) & _
           "���[�J�[:" & Chr(13) & _
           Chr(9) & maker.call_name & Chr(13) & _
           "�i��:" & Chr(13) & _
           Chr(9) & item.product_number & Chr(13) & _
           "�����:" & Chr(13) & _
           Chr(9) & Trader.company_name & Chr(13) & Chr(13) & _
           strVer
    msg = basp.SendMail(NS, Mailto, Mailfrom, Subject, Body, "")
    Set basp = Nothing
    If msg <> "" Then
        MsgBox ("�\�����o���܂���ł���" & Chr(13) & _
                "�Ǘ��҂֘A�����Ă�������" & Chr(13) & _
                msg)
    Else
        MsgBox ("�\�����������܂���")
    End If
End Sub
Function initLoss(strName As String)
'���X���׃V�[�g�̏����ݒ�
    Dim shtMy As Worksheet
    Dim rngMy As Range
    Dim shpMy As Shape
    Dim strColName() As String
    Dim varColWidth As Variant
    Dim i As Long
    Set shtMy = ActiveWorkbook.Sheets(strName)
    Set rngMy = shtMy.Range("a1:i2")
    rngMy.Merge
    With rngMy
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .Font.Size = 24
    End With
    rngMy.Value = strName
    shtMy.Buttons.Add(28.5, 20.75, 72, 24).Select
    Selection.OnAction = "zaiko.xls!���C��"
    Selection.Characters.text = "���C��"
    rows("7:7").Select
    ActiveWindow.FreezePanes = True

    Set rngMy = shtMy.Range("a5")
    strColName = Split("���[�J�[,�J�e�S���[,�i��,�i�ԁE�^��,����,����,���v,���X�g��", ",")
    varColWidth = Array(7, 8, 8, 19, 7, 5.38, 7.38, 10.88)
    With shtMy.Cells
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = True
        .ReadingOrder = xlContext
    End With
    For i = 0 To UBound(strColName)
        rngMy.Value = strColName(i)
        rngMy.ColumnWidth = varColWidth(i)
        Set rngMy = rngMy.Offset(0, 1)
    Next
    shtMy.Columns("e:g").NumberFormatLocal = "#,##0;[��]-#,##0"
    shtMy.Columns("h").NumberFormatLocal = "ge.mm.dd hh:mm"
    Range("B5").Select
    Selection.AutoFilter
    Set shtMy = Nothing
    Set rngMy = Nothing
End Function
Sub TenantAccauntSort(data() As TenantAccaunts, min As Long, max As Long)
'�N�C�b�N�\�[�g
    Dim i As Long, j As Long
    Dim base As Long
    Dim tmp As TenantAccaunts
    
    base = CLng(data((min + max) / 2).delivery_date)
    i = min
    j = max
    Do
        Do While CLng(data(i).delivery_date) < base
            i = i + 1
        Loop
        Do While CLng(data(j).delivery_date) > base
            j = j - 1
        Loop
        If i >= j Then
            Exit Do
        Else
            tmp = data(i)
            data(i) = data(j)
            data(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If min < i - 1 Then
        Call TenantAccauntSort(data, min, i - 1)
    End If
    If max > j + 1 Then
        Call TenantAccauntSort(data, j + 1, max)
    End If

End Sub
Sub itemSort(data() As Articles, min As Long, max As Long)
'�N�C�b�N�\�[�g
    Dim i As Long, j As Long
    Dim base As String
    Dim tmp As Articles
    
    base = data((min + max) / 2).id
    i = min
    j = max
    Do
        Do While data(i).id < base
            i = i + 1
        Loop
        Do While data(j).id > base
            j = j - 1
        Loop
        If i >= j Then
            Exit Do
        Else
            tmp = data(i)
            data(i) = data(j)
            data(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If min < i - 1 Then
        Call itemSort(data, min, i - 1)
    End If
    If max > j + 1 Then
        Call itemSort(data, j + 1, max)
    End If
End Sub
Sub SettleItemSortByTenantCode(data() As SettleArticles, min As Long, max As Long)
'�N�C�b�N�\�[�g
    Dim i As Long, j As Long
    Dim base As String
    Dim tmp As SettleArticles
    
    base = data((min + max) / 2).tenant_code
    i = min
    j = max
    Do
        Do While data(i).tenant_code < base
            i = i + 1
        Loop
        Do While data(j).tenant_code > base
            j = j - 1
        Loop
        If i >= j Then
            Exit Do
        Else
            tmp = data(i)
            data(i) = data(j)
            data(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If min < i - 1 Then
        Call SettleItemSortByTenantCode(data, min, i - 1)
    End If
    If max > j + 1 Then
        Call SettleItemSortByTenantCode(data, j + 1, max)
    End If

End Sub
Sub addTenantAccauntData(data As TenantAccaunts, accaunt() As TenantAccaunts, j As Long)
    accaunt(j) = data
    j = j + 1
    ReDim Preserve accaunt(j)
End Sub
Function tax(price) As Long
    tax = Int(CCur(price) * CCur(1.05))
End Function
Sub MonthDegreeProcessP2(MaruhiroP As Boolean, _
                            TenantP As Boolean, _
                            BillP As Boolean, _
                            UriageP As Boolean, _
                            bill_date As String)
    Dim mySht As Worksheet
    Dim settle_item() As SettleArticles
    Dim customer() As Customers
    Dim da() As DeliveryAccount
    Dim ta() As TenantAccaunts
    Dim ta_is_null As Boolean
    Dim bill_type As String
    Dim strState As String
                            
    strState = "1"
    '�o�Ƀf�[�^�����ύς݂ֈړ�
    Call moveDeliveryToSettleItem
    
    '�ۍL����������쐬����
    If getSettleItemList(bill_date, settle_item) = False Then
        MsgBox ("�����̐����͂���܂���")
        Exit Sub
    End If
    Call postSettleItemsToDeliveryAccount(da, settle_item)
    Call MakeDeliveryAccount2(da, settle_item(0).settle_date)
    Set mySht = Sheets("�ۍL��������")
    If MaruhiroP = True Then
        Call PrintPage(strState, mySht)
    End If

    '�e�i���g����������쐬����
    If postSettleItemsToTenantAccaunt(ta, settle_item, "�e�i���g�T��") = True Then
        Call makeTenantAccaunts2(ta, settle_item(0).settle_date)
        Set mySht = Sheets("�e�i���g��������")
        If TenantP = True Then
            Call PrintPage(strState, mySht)
        End If
    Else
        MsgBox "�e�i���g��������͂���܂���"
        ReDim ta(0)
    End If
    
    '�������z�ꗗ���쐬����
    ta_is_null = True
    If postSettleItemsToTenantAccaunt(ta, settle_item, "�e�i���g�T��") = True Then
        ta_is_null = False
    End If
    Call makeBillList2(da, ta, ta_is_null, settle_item(0).settle_date)
    Set mySht = Sheets("�������z�ꗗ�\")
    If BillP = True Then
        Call PrintPage(strState, mySht)
    End If
    
    '����ꗗ���쐬����
    ta_is_null = True
    bill_type = getBillTypes
    If postSettleItemsToTenantAccaunt(ta, settle_item, bill_type) = True Then
        ta_is_null = False
    End If
    Call makeSalesList2(da, ta, ta_is_null, settle_item(0).settle_date)
    Set mySht = Sheets("����ꗗ�\")
    If MaruhiroP = True Then
        Call PrintPage(strState, mySht)
    End If
    
    Range("a1").Select
    MsgBox ("���x�������������܂���")
ending:
    Set mySht = Nothing
End Sub

