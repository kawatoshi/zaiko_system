Attribute VB_Name = "getModule"
Option Explicit
Function getItem(strId As String, itemData As Articles) As String
'�A�C�e��ID����item���擾����
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("�i��")
    Set rngFind = getFindRange(shtMy, Articles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getItem = "findERROR": GoTo ending
    On Error GoTo 0
    itemData.id = rngFind.Cells(1, Articles_id_COL).Value
    itemData.category = rngFind.Cells(1, Articles_category_COL).Value
    itemData.name = rngFind.Cells(1, Articles_name_COL).Value
    itemData.product_number = rngFind.Cells(1, Articles_product_number_COL).Value
    itemData.maker_id = rngFind.Cells(1, Articles_maker_id_COL).Value
    itemData.fujibil_code = rngFind.Cells(1, Articles_fujibil_code_COL).Value
    itemData.JAN_CODE = rngFind.Cells(1, Articles_JAN_code_COL).Value
    itemData.cost = rngFind.Cells(1, Articles_cost_COL).Value
    itemData.price_without_tax = rngFind.Cells(1, Articles_price_without_tax_COL).Value
    itemData.tax = rngFind.Cells(1, Articles_tax_COL).Value
    itemData.price = rngFind.Cells(1, Articles_price_COL).Value
    itemData.trader_id = rngFind.Cells(1, Articles_trader_id_COL).Value
    itemData.entry_date = rngFind.Cells(1, Articles_entry_date_COL).Value
    getItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function getTrader(strId As String, TraderData As Traders) As String
'�����Ǝ�ID��������Ǝ҃f�[�^���擾����
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getTrader = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("����Ǝ�")
    Set rngFind = getFindRange(shtMy, Traders_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getTrader = "findERROR": GoTo ending
    On Error GoTo 0
    TraderData.id = rngFind.Cells(1, Traders_id_COL).Value
    TraderData.company_name = rngFind.Cells(1, Traders_company_name_COL).Value
    TraderData.tel = rngFind.Cells(1, Traders_tel_COL).Value
    TraderData.address = rngFind.Cells(1, Traders_address_COL).Value
    TraderData.person_name = rngFind.Cells(1, Traders_person_name_COL).Value
    getTrader = "OK"
ending:
    Set shtMy = Nothing
End Function
Function getMaker(strId As String, makerData As makers) As String
'���[�JID���烁�[�J�[�f�[�^���擾����
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getMaker = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("���[�J�[")
    Set rngFind = getFindRange(shtMy, Makers_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getMaker = "findERROR": GoTo ending
    On Error GoTo 0
    makerData.id = rngFind.Cells(1, Makers_id_COL).Value
    makerData.name = rngFind.Cells(1, Makers_name_COL).Value
    makerData.call_name = rngFind.Cells(1, Makers_call_name_COL).Value
    getMaker = "OK"
ending:
    Set shtMy = Nothing
End Function
Function getCustomer(strId As String, CustomerData As Customers) As String
'�����Ǝ�ID��������Ǝ҃f�[�^���擾����
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getCustomer = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("�����")
    Set rngFind = getFindRange(shtMy, Customers_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getCustomer = "findERROR": GoTo ending
    On Error GoTo 0
    CustomerData.id = rngFind.Cells(1, Customers_id_COL).Value
    CustomerData.site = rngFind.Cells(1, Customers_site_COL).Value
    CustomerData.floor = rngFind.Cells(1, Customers_floor_COL).Value
    CustomerData.place = rngFind.Cells(1, Customers_place_COL).Value
    CustomerData.claim_name = rngFind.Cells(1, Customers_claim_name_COL).Value
    CustomerData.tenant_code = rngFind.Cells(1, Customers_tenant_code_COL).Value
    CustomerData.A_table = rngFind.Cells(1, Customers_A_table_COL).Value
    CustomerData.bill_type = rngFind.Cells(1, Customers_bill_type_COL).Value
    getCustomer = "OK"
ending:
    Set shtMy = Nothing
End Function
Function getStockItem(strId As String, StockItemData As StockArticles) As String
'�݌�ID����݌Ƀf�[�^���擾����
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getStockItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("�݌�")
    Set rngFind = getFindRange(shtMy, StockArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getStockItem = "findERROR": GoTo ending
    On Error GoTo 0
    StockItemData.id = rngFind.Cells(1, StockArticles_id_COL).Value
    StockItemData.buy_article_id = rngFind.Cells(1, StockArticles_buy_article_id_COL).Value
    StockItemData.item_id = rngFind.Cells(1, StockArticles_item_id_COL).Value
    StockItemData.cost = rngFind.Cells(1, StockArticles_cost_COL).Value
    StockItemData.number = rngFind.Cells(1, StockArticles_number_COL).Value
    StockItemData.final_delivery_date = rngFind.Cells(1, StockArticles_final_delivery_date_COL).Value
    StockItemData.receipt_article_id = rngFind.Cells(1, StockArticles_receipt_article_id_COL).Value
    getStockItem = "OK"
ending:
    Set shtMy = Nothing
End Function
Function getDeliveryItem(strId As String, DeliveryItemData As DeliveryArticles) As String
'�o�ɍς�ID����o�ɍς݃f�[�^���擾����
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getDeliveryItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("�o��")
    Set rngFind = getFindRange(shtMy, DeliveryArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getDeliveryItem = "findERROR": GoTo ending
    On Error GoTo 0
    DeliveryItemData.id = rngFind.Cells(1, DeliveryArticles_id_COL).Value
    DeliveryItemData.buy_article_id = rngFind.Cells(1, DeliveryArticles_buy_article_id_COL).Value
    DeliveryItemData.stock_article_id = rngFind.Cells(1, DeliveryArticles_stock_article_id_COL).Value
    DeliveryItemData.item_id = rngFind.Cells(1, DeliveryArticles_item_id_COL).Value
    DeliveryItemData.customer_id = rngFind.Cells(1, DeliveryArticles_customer_id_COL).Value
    DeliveryItemData.cost = rngFind.Cells(1, DeliveryArticles_cost_COL).Value
    DeliveryItemData.item_price_without_tax = rngFind.Cells(1, DeliveryArticles_item_price_without_tax_COL).Value
    DeliveryItemData.item_price = rngFind.Cells(1, DeliveryArticles_item_price_COL).Value
    DeliveryItemData.number = rngFind.Cells(1, DeliveryArticles_number_COL).Value
    DeliveryItemData.sum = rngFind.Cells(1, DeliveryArticles_sum_COL).Value
    DeliveryItemData.bill_type = rngFind.Cells(1, DeliveryArticles_bill_type_COL).Value
    DeliveryItemData.delivery_date = rngFind.Cells(1, DeliveryArticles_delivery_date_COL).Value
    getDeliveryItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function getSettleItem(strId As String, SettleItemData As SettleArticles) As String
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getSettleItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("���ύ�")
    Set rngFind = getFindRange(shtMy, SettleArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getSettleItem = "findERROR": GoTo ending
    On Error GoTo 0
    SettleItemData.id = rngFind.Cells(1, SettleArticles_id_COL).Value
    SettleItemData.buy_article_id = rngFind.Cells(1, SettleArticles_buy_article_id_COL).Value
    SettleItemData.stock_article_id = rngFind.Cells(1, SettleArticles_stock_article_id_COL).Value
    SettleItemData.item_id = rngFind.Cells(1, SettleArticles_item_id_COL).Value
    SettleItemData.customer_id = rngFind.Cells(1, SettleArticles_customer_id_COL).Value
    SettleItemData.cost = rngFind.Cells(1, SettleArticles_cost_COL).Value
    SettleItemData.item_price_without_tax = rngFind.Cells(1, SettleArticles_item_price_without_tax_COL).Value
    SettleItemData.item_price = rngFind.Cells(1, SettleArticles_item_price_COL).Value
    SettleItemData.number = rngFind.Cells(1, SettleArticles_number_COL).Value
    SettleItemData.sum = rngFind.Cells(1, SettleArticles_sum_COL).Value
    SettleItemData.bill_type = rngFind.Cells(1, SettleArticles_bill_type_COL).Value
    SettleItemData.tenant_code = rngFind.Cells(1, SettleArticles_tenant_code_COL).Value
    SettleItemData.delivery_date = rngFind.Cells(1, SettleArticles_delivery_date_COL).Value
    SettleItemData.settle_date = rngFind.Cells(1, SettleArticles_settle_date_COL).Value
    SettleItemData.bill_date = rngFind.Cells(1, SettleArticles_bill_date_COL).Value
    getSettleItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function getLostItem(strId As String, LostItemData As LostArticles) As String
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getLostItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("���X")
    Set rngFind = getFindRange(shtMy, LostArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getLostItem = "findERROR": GoTo ending
    On Error GoTo 0
    LostItemData.id = rngFind.Cells(1, LostArticles_id_COL).Value
    LostItemData.buy_article_id = rngFind.Cells(1, LostArticles_buy_article_id_COL).Value
    LostItemData.stock_article_id = rngFind.Cells(1, LostArticles_stock_article_id_COL).Value
    LostItemData.item_id = rngFind.Cells(1, LostArticles_item_id_COL).Value
    LostItemData.cost = rngFind.Cells(1, LostArticles_cost_COL).Value
    LostItemData.number = rngFind.Cells(1, LostArticles_number_COL).Value
    LostItemData.lost_date = rngFind.Cells(1, LostArticles_lost_date_COL).Value
    LostItemData.note = rngFind.Cells(1, LostArticles_note_COL).Value
    getLostItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function getReturnItem(strId As String, ReturnItemData As ReturnArticles) As String
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getReturnItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("�ԕi����")
    Set rngFind = getFindRange(shtMy, ReturnArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getReturnItem = "findERROR": GoTo ending
    On Error GoTo 0
    ReturnItemData.id = rngFind.Cells(1, ReturnArticles_id_COL).Value
    ReturnItemData.buy_article_id = rngFind.Cells(1, ReturnArticles_buy_article_id_COL).Value
    ReturnItemData.stock_article_id = rngFind.Cells(1, ReturnArticles_stock_article_id_COL).Value
    ReturnItemData.item_id = rngFind.Cells(1, ReturnArticles_item_id_COL).Value
    ReturnItemData.customer_id = rngFind.Cells(1, ReturnArticles_customer_id_COL).Value
    ReturnItemData.cost = rngFind.Cells(1, ReturnArticles_cost_COL).Value
    ReturnItemData.item_price = rngFind.Cells(1, ReturnArticles_item_price_COL).Value
    ReturnItemData.number = rngFind.Cells(1, ReturnArticles_number_COL).Value
    ReturnItemData.return_date = rngFind.Cells(1, ReturnArticles_return_date_COL).Value
    getReturnItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function getBuyItem(strId As String, BuyItemData As BuyArticles) As String
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getBuyItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("����")
    Set rngFind = getFindRange(shtMy, BuyArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getBuyItem = "findERROR": GoTo ending
    On Error GoTo 0
    BuyItemData.id = rngFind.Cells(1, BuyArticles_id_COL).Value
    BuyItemData.payment_id = rngFind.Cells(1, BuyArticles_payment_id_COL).Value
    BuyItemData.item_id = rngFind.Cells(1, BuyArticles_item_id_COL).Value
    BuyItemData.trader_id = rngFind.Cells(1, BuyArticles_trader_id_COL).Value
    BuyItemData.cost = rngFind.Cells(1, BuyArticles_cost_COL).Value
    BuyItemData.number = rngFind.Cells(1, BuyArticles_number_COL).Value
    BuyItemData.in_stock_date = rngFind.Cells(1, BuyArticles_in_stock_date_COL).Value
    getBuyItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Public Function getEndRow(strColumns As String, Optional shtGets As Worksheet) As Long
'���߂�������ŕ\������Ă���ŏI�s���擾����
'���͂��Ȃ��ꍇ��0��Ԃ�
    Dim rngColumns As Range
    Dim rngSort As Range

    If shtGets Is Nothing Then Set shtGets = ActiveSheet
    Set rngColumns = shtGets.Columns(strColumns)

    With rngColumns
        Set rngSort = .Find(what:="*", after:=.Cells(1), LookIn:=xlValues, searchorder:=xlByRows, searchdirection:=xlPrevious)
    End With
    If rngSort Is Nothing Then
        getEndRow = 0
    Else
        getEndRow = rngSort.Row
    End If
End Function
Function getMaxNo(rngMy As Range) As Long
'�^����ꂽRange�̍ő吔��long�l�ŕԂ�
    If rngMy Is Nothing Then
        getMaxNo = 0
    Else
        getMaxNo = WorksheetFunction.max(rngMy)
    End If
End Function
Function getDeliveryItemMaxNo() As Long
'�o�ɕi�̍ő�No��Ԃ�
    Dim shtDelivery As Worksheet
    Dim shtSettle As Worksheet
    Dim rngMy As Range
    Dim DMax As Long, SMax As Long
    
    Set shtDelivery = Sheets("�o��")
    Set shtSettle = Sheets("���ύ�")
    Set rngMy = getFindRange(shtDelivery, DeliveryArticles_id_COL)
    DMax = getMaxNo(rngMy)
    Set rngMy = getFindRange(shtSettle, SettleArticles_id_COL)
    SMax = getMaxNo(rngMy)
    If DMax > SMax Then getDeliveryItemMaxNo = DMax + 1: GoTo ending
    If SMax > DMax Then getDeliveryItemMaxNo = SMax + 1: GoTo ending
    If DMax = 0 And SMax = 0 Then getDeliveryItemMaxNo = 1: GoTo ending
    getDeliveryItemMaxNo = SMax
ending:
    Set rngMy = Nothing
    Set shtDelivery = Nothing
    Set shtSettle = Nothing
End Function

Public Function getKeyData(strKey As String, rngFind As Range, _
                           Optional strType As String = "address", Optional lngLook As Long = xlPart) As String
'rngKey�ɗ^����ꂽ�͈͂ɑ��݂���strKey�ƃf�[�^���󔒕t�̕�����ŕԂ�
' strType:="
'address=>��΃A�h���X
'row=>�s
Dim lngStartRow As Long, lngRow As Long
Dim strStartAddress As String
Dim rngSort As Range

If strKey = "" Then GoTo ending
If strKey = "*" Then GoTo ending
Set rngFind = rngFind.Offset(-1, 0).Resize(rngFind.rows.Count + 1)
With rngFind
'    Debug.Print rngFind.Worksheet.Name
'    .Activate
    Set rngSort = .Find(what:=strKey, LookIn:=xlValues, _
                    lookat:=lngLook, searchorder:=xlByColumns, searchdirection:=xlNext, _
                    MatchCase:=False)
    If Not rngSort Is Nothing Then
        strStartAddress = rngSort.address
        Do
            Select Case strType
                Case "address"
                    getKeyData = getKeyData & rngSort.address & " "
                Case "row"
                    getKeyData = getKeyData & rngSort.Row & " "
            End Select
            Set rngSort = .FindNext(rngSort)
        Loop While Not rngSort Is Nothing And rngSort.address <> strStartAddress
        getKeyData = RTrim(getKeyData)
    End If
End With
ending:
Set rngSort = Nothing
End Function
Function getFindRange(shtFind As Worksheet, lngDataCol As Long) As Range
'�f�[�^���X�g�͈̔͂��擾����
    Dim lngMaxRow As Long
    Dim strAddress As String
    strAddress = Columns(lngDataCol).address
    lngMaxRow = getEndRow(strAddress, shtFind)
'    lngMaxRow = shtFind.Cells.SpecialCells(xlCellTypeLastCell).Row
    If DATA_START_ROW > lngMaxRow Then
        GoTo ending
    End If
    With shtFind
    Set getFindRange = .Range(.Cells(DATA_START_ROW, lngDataCol), _
                                   .Cells(lngMaxRow, lngDataCol))
    End With
ending:
End Function
Function getDataRange(shtFind As Worksheet, lngEndCol As Long) As Range
'�f�[�^�͈͂��擾����
    Dim rng As Range
    Set rng = getFindRange(shtFind, 1)
    Set getDataRange = rng.Resize(rng.rows.Count, lngEndCol)
End Function
Function getTableDatas(shtGet As Worksheet, strKey As String, _
                        lngFindCol As Long, lngGetCol As Long) As String
'shtGet�̃V�[�g����strKey�ɊY������lngGetCol��f�[�^���擾���ĕԂ�
    Dim rngFind As Range
    Dim lngEndCol As Long
    Dim varData As Variant
    Dim varKeyData As Variant
    Dim varKey As Variant
    Dim i As Long
    If strKey = "" Then GoTo ending
    
    '�����͈͂̔z��擾
    If lngFindCol < lngGetCol Then
        lngEndCol = lngGetCol
    Else
        lngEndCol = lngFindCol
    End If
    Set rngFind = getFindRange(shtGet, lngFindCol)
    If rngFind Is Nothing Then GoTo ending
    Set rngFind = rngFind.Offset(0, -lngFindCol + 1).Resize(rngFind.rows.Count, lngEndCol)
    varData = rngFind.Value
    varKeyData = Split(strKey)
    '�擾�f�[�^�̌����i�f�[�^���P�����Ȃ��ꍇ�j
    If rngFind.Count = 1 Then
        For Each varKey In varKeyData
            If CStr(varData) Like CStr(varKey) Then
                getTableDatas = CStr(varData)
                GoTo ending
            End If
        Next
    End If
    '�擾�f�[�^�̌����i�f�[�^���P�ȏ�j
    For i = 1 To UBound(varData)
        For Each varKey In varKeyData
            If varData(i, lngFindCol) Like CStr(varKey) Then
                getTableDatas = getTableDatas & " " & CStr(varData(i, lngGetCol))
            End If
        Next
    Next
    getTableDatas = Trim(getTableDatas)
ending:
    Set rngFind = Nothing
End Function

Sub getArray(rngArray As Range, strData() As String)
'rngArray��strList()�̔z��ɓ����
    Dim i As Long
    If rngArray Is Nothing Then GoTo ending
    ReDim strData(rngArray.rows.Count - 1)
    For i = 0 To UBound(strData)
        strData(i) = rngArray.Cells(i + 1)
    Next
ending:
End Sub
Function getItemIDsFromItemName(itemName As String) As String
'itemName����i�����������Aid��Ԃ�
    Dim shtMy As Worksheet
    Set shtMy = Sheets("�i��")
    If itemName = "" Then itemName = "*"
    getItemIDsFromItemName = getTableDatas(shtMy, itemName, _
                                        Articles_name_COL, Articles_id_COL)
    Set shtMy = Nothing
End Function

Function getItemIdFromItemJanCode(strJan As String) As String
'strJan����i���e�[�u��JAN_code��̊Y���s���擾���Aid��Ԃ�
    Dim shtMy As Worksheet
    Dim id() As String
    Set shtMy = Sheets("�i��")
    
    id = Split(getTableDatas(shtMy, strJan, Articles_JAN_code_COL, Articles_id_COL), " ")
    Select Case UBound(id)
    Case -1
        getItemIdFromItemJanCode = ""
    Case Else
        getItemIdFromItemJanCode = id(0)
    End Select
    Set shtMy = Nothing
End Function

Function getItemIdFromMakerCallNameAndProductNumber(strMakerCallName As String, strProductNumber As String) As String
'���[�J�[�Ăі��ƕi�Ԃ���ItemId��Ԃ�
    Dim ItemIdOnMakerCallName As String
    Dim ItemIdOnProductNum As String
    Dim strId() As String
    Dim shtMy As Worksheet
    Set shtMy = Sheets("�i��")
    
    ItemIdOnMakerCallName = getItemIdsFromMakerCallName(strMakerCallName)
    strId = Split(ItemIdOnMakerCallName)
    If UBound(strId) = 0 Then
        getItemIdFromMakerCallNameAndProductNumber = strId(0)
        GoTo ending
    End If
    ItemIdOnProductNum = getTableDatas(shtMy, strProductNumber, _
                                       Articles_product_number_COL, Articles_id_COL)
    strId = Split(ItemIdOnProductNum)
    If UBound(strId) = 0 Then
        getItemIdFromMakerCallNameAndProductNumber = strId(0)
        GoTo ending
    End If
    strId = Split(Trim(ItemIdOnMakerCallName & " " & ItemIdOnProductNum), " ")
    If UBound(strId) = -1 Then GoTo ending
    Call DuplicationDraw(strId)
    If UBound(strId) = 0 Then
        getItemIdFromMakerCallNameAndProductNumber = strId(0)
    End If
ending:
End Function

Function getProductIDsFromKeys(strList() As String, strKeys As String, lngQueryCol As Long) As String
'�i�ڃV�[�g��lngQueryCol��strKeys�̂����ꂩ�̏����ɓ��Ă͂܂�lngAnsCol��f�[�^��
'strList�֎��[����B
    Dim shtMy As Worksheet
    Dim strKey() As String
    Dim strId() As String
    Dim i As Long
    
    Set shtMy = Sheets("�i��")
    strKey = Split(strKeys, " ")
    ReDim strId(UBound(strKey))
    For i = 0 To UBound(strKey)
        strId(i) = getTableDatas(shtMy, strKey(i), lngQueryCol, Articles_id_COL)
    Next
    getProductIDsFromKeys = DuplicationMerge(strId)
    strList = Split(getItemProductNumFromItemIds(getProductIDsFromKeys), " ")
End Function

Function getItemNameList(strList() As String, strKeys As String, queryCol As Long) As String
'queryCol��strkeys�̂����ꂩ�̏����ɓ��Ă͂܂�ansCol��f�[�^��
'strList�֎��[���Aitem_id�̕������Ԃ�
    Dim shtMy As Worksheet
    Dim strKey() As String
    Dim strId() As String
    Dim i As Long
    
    Set shtMy = Sheets("�i��")
    strKey = Split(strKeys, " ")
    ReDim strId(UBound(strKey))
    For i = 0 To UBound(strKey)
        strId(i) = getTableDatas(shtMy, strKey(i), queryCol, Articles_id_COL)
    Next
    getItemNameList = DuplicationMerge(strId)
    strList = Split(getTableDatas(shtMy, getItemNameList, Articles_id_COL, Articles_name_COL), " ")
    Call DuplicationMerge(strList)
    Set shtMy = Nothing
    End Function

Function getClosingdate(Optional dteDate As Date) As Date
'�������s�����ߓ���Ԃ�
    Dim lngDay As Long
    Dim lngYear As Long
    Dim lngMonth As Long
    Dim ClosingDay As Long
    If dteDate = 0 Then dteDate = Now
    ClosingDay = Range("closing_day")
    lngDay = Day(dteDate)
'    lngDay = 26
    If lngDay <= ClosingDay Then
        getClosingdate = DateAdd("m", -1, DateSerial(Year(dteDate), Month(dteDate), ClosingDay + 1))
    Else
        getClosingdate = DateSerial(Year(dteDate), Month(dteDate), ClosingDay + 1)
    End If
End Function

Function getBilldateOnStr(Optional dteDate As Date) As String
'��������Ԃ�
    Dim dteMy As Date
    If dteDate = 0 Then
        dteMy = getClosingdate
    Else
        dteMy = dteDate
    End If
    getBilldateOnStr = Year(dteMy) & "/" & Month(dteMy) & "����"
End Function

Function getHinmoku(strItemId As String) As String
'�i��ID�����ӂɔ��ʏo����^�Ԃ�Ԃ�
    Dim item As Articles
    Dim maker As makers

    Call getItem(strItemId, item)
    Call getMaker(item.maker_id, maker)
    
    getHinmoku = maker.call_name & " : " & item.product_number
End Function
Function getBillIds(Optional strKey As String = "������") As String
'��������Ă��Ȃ��o��id��Ԃ�
    Dim shtMy As Worksheet
    Dim rngMy As Range
    Set shtMy = Sheets("���ύ�")
    Set rngMy = getFindRange(shtMy, SettleArticles_bill_date_COL)
    
    getBillIds = getTableDatas(shtMy, strKey, _
                               SettleArticles_bill_date_COL, SettleArticles_id_COL)
    
    Set shtMy = Nothing
    Set rngMy = Nothing
End Function

Function getTmpSettleItem(strId As String, TmpSettleItemData As TmpSettleArticles) As String
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim strAddress As String

    If strId = "" Then getTmpSettleItem = "idERROR": GoTo ending
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("Tmp����")
    Set rngFind = getFindRange(shtMy, TmpSettleArticles_id_COL)
    strAddress = getKeyData(strId, rngFind, , xlWhole)
    Set rngFind = shtMy.Range(strAddress)
    If Err.number <> 0 Then getTmpSettleItem = "findERROR": GoTo ending
    On Error GoTo 0
    TmpSettleItemData.id = rngFind.Cells(1, TmpSettleArticles_id_COL).Value
    TmpSettleItemData.delivery_date = rngFind.Cells(1, TmpSettleArticles_delivery_date_COL).Value
    TmpSettleItemData.customer = rngFind.Cells(1, TmpSettleArticles_customer_COL).Value
    TmpSettleItemData.claim_name = rngFind.Cells(1, TmpSettleArticles_claim_name_COL).Value
    TmpSettleItemData.maker = rngFind.Cells(1, TmpSettleArticles_maker_COL).Value
    TmpSettleItemData.item_name = rngFind.Cells(1, TmpSettleArticles_item_name_COL).Value
    TmpSettleItemData.item = rngFind.Cells(1, TmpSettleArticles_item_COL).Value
    TmpSettleItemData.item_price = rngFind.Cells(1, TmpSettleArticles_item_price_COL).Value
    TmpSettleItemData.number = rngFind.Cells(1, TmpSettleArticles_number_COL).Value
    TmpSettleItemData.sum = rngFind.Cells(1, TmpSettleArticles_sum_COL).Value
    TmpSettleItemData.bill_type = rngFind.Cells(1, TmpSettleArticles_bill_type_COL).Value
    getTmpSettleItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function getMakerIdFromItemIds(strKeyList As String) As String
'ItemId����d�����Ȃ�MakerId��Ԃ�
     Dim shtMy As Worksheet
     Dim i As Long
     Dim strList() As String
     Set shtMy = Sheets("�i��")
     If chkSplit(strKeyList, strList) = False Then GoTo ending
     For i = 0 To UBound(strList)
        getMakerIdFromItemIds = getMakerIdFromItemIds & _
                                 getTableDatas(shtMy, strList(i), Articles_id_COL, Articles_maker_id_COL) & _
                                 " "
     Next
     getMakerIdFromItemIds = Trim(getMakerIdFromItemIds)
     If chkSplit(getMakerIdFromItemIds, strList) = True Then
        Call DuplicationMerge(strList)
        getMakerIdFromItemIds = Join(strList, " ")
     End If
ending:
     Set shtMy = Nothing
End Function
Function getMakerCallNamesFromItemName(strKey As String, Optional stockOnly As Boolean = False) As String
'�i������d�����Ȃ����[�J�[�Ăі���Ԃ�
    Dim shtMy As Worksheet
    Dim shtMaker As Worksheet
    Dim strMakerIDs As String
    Dim strItemIDs As String
    Dim strStockItemIDs As String
    Dim varData As Variant
    Dim strList() As String
    Set shtMy = Sheets("�i��")
    Set shtMaker = Sheets("���[�J�[")
    
    strMakerIDs = getTableDatas(shtMy, strKey, Articles_name_COL, Articles_maker_id_COL)
    strList = Split(strMakerIDs, " ")
    strMakerIDs = DuplicationMerge(strList)
    If stockOnly = True Then
        strItemIDs = getTableDatas(shtMy, strMakerIDs, Articles_maker_id_COL, Articles_id_COL)
        strStockItemIDs = getTableDatas(Worksheets("�݌�"), _
                                        strItemIDs, _
                                        StockArticles_item_id_COL, _
                                        StockArticles_item_id_COL)
        strList = Split(strStockItemIDs)
        strStockItemIDs = DuplicationMerge(strList)
        strMakerIDs = getTableDatas(shtMy, strStockItemIDs, Articles_id_COL, Articles_maker_id_COL)
    End If
    If strMakerIDs = "" Then GoTo ending
    strList = Split(strMakerIDs, " ")
    strMakerIDs = DuplicationMerge(strList)
    For Each varData In strList
        getMakerCallNamesFromItemName = getMakerCallNamesFromItemName & " " & _
                                        getTableDatas(shtMaker, CStr(varData), Makers_id_COL, Makers_call_name_COL)
    Next
ending:
    Set shtMy = Nothing
    Set shtMaker = Nothing
End Function

Function getMakerCallNameFromItemProductNum(strProductNum As String, Optional stockOnly As Boolean = False) As String
'�i�Ԃ���d�����Ȃ����[�J�[�Ăі���Ԃ�
    Dim strMakerIDs As String
    Dim strList() As String
    
    Select Case stockOnly
    Case False
        strMakerIDs = getMakerIdsFromItemProductNum(strProductNum)
    Case Else
        strMakerIDs = getMakerIdsFromStockProductNum(strProductNum)
    End Select
    getMakerCallNameFromItemProductNum = getMakerCallNameFromMakerIds(strMakerIDs, strList)
End Function

Function getMakerCallNameFromMakerIds(strItemIDs As String, strList() As String) As String
'���[�J�[id����d���̂Ȃ����[�J�[�ď̂�Ԃ�
    Dim shtMy As Worksheet
    Dim strState As String
    Dim strId() As String
    Dim varData As Variant
    If strItemIDs = "" Then GoTo ending
    Set shtMy = ActiveWorkbook.Sheets("���[�J�[")
    
    strId = Split(strItemIDs, " ")
    For Each varData In strId
        strState = strState & " " & _
                    getTableDatas(shtMy, CStr(varData), Makers_id_COL, Makers_call_name_COL)
    Next
    strState = Trim(strState)
    If strState = "" Then GoTo ending
    strList = Split(strState)
    Call DuplicationMerge(strList)
    getMakerCallNameFromMakerIds = Join(strList, " ")
ending:
    Set shtMy = Nothing
End Function

Function getMakerCallNameFromJanCode(strJan As String, strList() As String) As String
'JANCODE����d���̂Ȃ����[�J�[�Ăі���Ԃ�
'������strList�Ɍ��ʂ̈ꗗ�A�Ԓl�Ƀ��[�J�[id��Ԃ�
    Dim strMakerIDs As String

    strMakerIDs = getMakerIdsFromItems(Left(strJan, 6) & "*", Articles_JAN_code_COL, Articles_maker_id_COL)
    Call getMakerCallNameFromMakerIds(strMakerIDs, strList)
    getMakerCallNameFromJanCode = strMakerIDs
End Function

Function getItemIdsFromMakerCallName(key As String) As String
'���[�J�[�ď̂ɊY������i��ID��Ԃ�
    Dim shtItem As Worksheet
    Dim strMakerId As String
    Set shtItem = Sheets("�i��")
    strMakerId = getMakerIdFromMakerCallName(key)
    If strMakerId = "" Then _
        getItemIdsFromMakerCallName = "": GoTo ending
    getItemIdsFromMakerCallName = _
        getTableDatas(shtItem, strMakerId, Articles_maker_id_COL, Articles_id_COL)
ending:
    Set shtItem = Nothing
End Function
Function getItemProductNumFromItemIds(strItemIDs As String) As String
'�i��ID�Q����d���̂Ȃ��i�Ԃ�Ԃ�
    Dim strId() As String
    Dim strList() As String
    Dim shtMy As Worksheet
    Dim strProductNums As String
    Dim varData As Variant
    Set shtMy = Sheets("�i��")
    If strItemIDs = "" Then GoTo ending
    
    strProductNums = getTableDatas(shtMy, strItemIDs, Articles_id_COL, Articles_product_number_COL)

    strList = Split(Trim(strProductNums), " ")
    If UBound(strList) < 0 Then GoTo ending
    Call DuplicationMerge(strList)
    getItemProductNumFromItemIds = Join(strList, " ")
ending:
    Set shtMy = Nothing
End Function
Function getMakerIdFromMakerCallName(key As String) As String
'���[�J�[�Ăі����烁�[�J�[ID���擾����
    Dim shtMaker As Worksheet
    Set shtMaker = Sheets("���[�J�[")
    getMakerIdFromMakerCallName = getTableDatas(shtMaker, key, _
                                                    Makers_call_name_COL, Makers_id_COL)
    Set shtMaker = Nothing
End Function

Function getDeliveryAccount2(strBilltype As String) As String
'�������@����Y������id�𒊏o���ĕԂ�
    Dim shtMy As Worksheet
    Set shtMy = Worksheets("Tmp����")
    
    getDeliveryAccount2 = getTableDatas(shtMy, strBilltype, TmpSettleArticles_bill_type_COL, TmpSettleArticles_id_COL)
ending:
    Set shtMy = Nothing
End Function
Function getTmpSettleItemClaimName(strBilltype As String) As String
'Tmp���ς��strBillType�ɊY���������於�̈ꗗ���擾����
    Dim shtMy As Worksheet
    Dim ClaimName() As String
    Set shtMy = Worksheets("Tmp����")
    
    getTmpSettleItemClaimName = getTableDatas(shtMy, strBilltype, _
                                              TmpSettleArticles_bill_type_COL, TmpSettleArticles_claim_name_COL)
    ClaimName = Split(getTmpSettleItemClaimName, " ")
    Call DuplicationMerge(ClaimName)
    getTmpSettleItemClaimName = Join(ClaimName, " ")
ending:
    Set shtMy = Nothing
End Function

Function getNumOfStock(strItemId As String) As String
'�݌ɂ̃X�g�b�N����Ԃ�
    Dim strState As String
    Dim num() As String
    Dim varData As Variant
    Dim dblNum As Double
    
    strState = getTableDatas(Sheets("�݌�"), strItemId, StockArticles_item_id_COL, StockArticles_number_COL)
    If chkSplit(strState, num()) = False Then getNumOfStock = "0": GoTo ending
    For Each varData In num()
        dblNum = dblNum + CDbl(varData)
    Next
    getNumOfStock = CStr(dblNum)
ending:
End Function
Function getSumOfStock(strItemId As String) As String
'�݌ɂ̋��z��Ԃ�
    Dim strState As String
    Dim cost() As String
    Dim num() As String
    Dim i As Long
    Dim dblSum As Double
    
    strState = getTableDatas(Sheets("�݌�"), strItemId, StockArticles_item_id_COL, StockArticles_cost_COL)
    If chkSplit(strState, cost()) = False Then getSumOfStock = "0": GoTo ending
    strState = getTableDatas(Sheets("�݌�"), strItemId, StockArticles_item_id_COL, StockArticles_number_COL)
    If chkSplit(strState, num()) = False Then getSumOfStock = "0": GoTo ending
    For i = 0 To UBound(cost)
        dblSum = dblSum + (CDbl(num(i)) * CDbl(cost(i)))
    Next
    getSumOfStock = CStr(dblSum)
ending:
End Function
Function getNewerBuyDate(strItemId As String) As Date
'�݌ɂ̍ŐV���ɓ���Ԃ�
    Dim strState As String
    Dim strId() As String
    Dim shtMy As Worksheet
    Set shtMy = Sheets("����")
    strState = getTableDatas(shtMy, strItemId, BuyArticles_item_id_COL, BuyArticles_id_COL)
    If chkSplit(strState, strId) = False Then GoTo ending
    getNewerBuyDate = getTableDatas(shtMy, strId(UBound(strId)), BuyArticles_id_COL, BuyArticles_in_stock_date_COL)
ending:
    Set shtMy = Nothing
End Function
Function getNewerDeliveryDate(strItemId As String) As Date
'�݌ɂ̍ŏI�o�ɓ���Ԃ�
    Dim strState As String
    Dim strId() As String
    Dim shtMy As Worksheet
    Set shtMy = Sheets("�݌�")
    strState = getTableDatas(shtMy, strItemId, StockArticles_item_id_COL, StockArticles_id_COL)
    If chkSplit(strState, strId) = False Then GoTo ending
    strState = getTableDatas(shtMy, strId(LBound(strId)), _
                             StockArticles_id_COL, StockArticles_final_delivery_date_COL)
    If strState = "" Then getNewerDeliveryDate = 0: GoTo ending
    getNewerDeliveryDate = CDate(strState)
ending:
    Set shtMy = Nothing
End Function

Function getSumFunction(rngSum As Range) As String
'rngSum�̍��v���v�Z���鎮�𕶎���ŕԂ�
    Dim strStartAddress As String
    Dim strEndAddress As String
    Dim lngE As Long
    lngE = rngSum.Count
    strStartAddress = rngSum.Cells(1).address
    strEndAddress = rngSum.Cells(lngE).address
    getSumFunction = "=subtotal(9," & strStartAddress & ":" & strEndAddress & ")"
End Function

Function getMakerFromItemId(strItemId As String) As makers
'�i��ID���烁�[�J�[�f�[�^���擾����
    Dim strState As String
    Dim shtMy As Worksheet
    Set shtMy = Sheets("�i��")
    strState = getTableDatas(shtMy, strItemId, Articles_id_COL, Articles_maker_id_COL)
    If strState = "" Then GoTo ending
    Call getMaker(strState, getMakerFromItemId)
ending:
    Set shtMy = Nothing
End Function
Function getMakerIdsFromItemProductNum(strProductNum As String) As String
'�^�Ԃ��烁�[�J�[Id���擾����
    Dim strState As String
    Dim shtMy As Worksheet
    Dim strMakerId() As String
    Dim i As Long
    Dim varData As Variant
    
    Set shtMy = Sheets("�i��")
    strState = getTableDatas(shtMy, strProductNum, Articles_product_number_COL, Articles_maker_id_COL)
    If strState = "" Then GoTo ending
    strMakerId = Split(strState, " ")
    Call DuplicationMerge(strMakerId)
    getMakerIdsFromItemProductNum = Join(strMakerId, " ")
ending:
    Set shtMy = Nothing
End Function
Function getMakerIdsFromStockProductNum(productNum As String) As String
'�^�Ԃ���݌ɂ̃��[�J�[Id���擾����
    Dim StockItemIDs As String
    Dim ItemIDs As String
    Dim MakerIDs As String
    Dim makerid() As String
    Dim shtMy As Worksheet
    
    Set shtMy = Sheets("�i��")
    ItemIDs = getTableDatas(shtMy, productNum, Articles_product_number_COL, Articles_id_COL)
    StockItemIDs = getTableDatas(Worksheets("�݌�"), ItemIDs, StockArticles_item_id_COL, StockArticles_item_id_COL)
    MakerIDs = getTableDatas(shtMy, StockItemIDs, Articles_id_COL, Articles_maker_id_COL)
    makerid = Split(MakerIDs)
    getMakerIdsFromStockProductNum = DuplicationMerge(makerid)
    Set shtMy = Nothing
End Function
Function getMakerIdsFromItems(strData As String, lngQueryCol As Long, lngAnsCol As Long) As String
'�i�ڂ̖₢���킹�f�[�^��������̓�����Ԃ�
    Dim strState As String
    Dim shtMy As Worksheet
    Dim strMakerId() As String
    
    Set shtMy = Sheets("�i��")
    strState = getTableDatas(shtMy, strData, lngQueryCol, lngAnsCol)
    If strState = "" Then GoTo ending
    strMakerId = Split(strState, " ")
    Call DuplicationMerge(strMakerId)
    getMakerIdsFromItems = Join(strMakerId, " ")
ending:
    Set shtMy = Nothing
End Function
Function getSumOfBill(strSettleIds As String, Optional strGetMode As String = "price") As Double
'�w�肳�ꂽ�����̃f�[�^��Ԃ�
    Dim strId() As String
    Dim Sitem As SettleArticles
    Dim i As Long
    Dim dblSum As Double
    
    If strSettleIds = "" Then GoTo ending
    strId = Split(strSettleIds, " ")
    For i = 0 To UBound(strId)
        Call getSettleItem(strId(i), Sitem)
        Select Case strGetMode
            Case "price"
                dblSum = dblSum + CDbl(Sitem.item_price) * CDbl(Sitem.number)
            Case "price_without_tax"
                dblSum = dblSum + CDbl(Sitem.item_price_without_tax) * CDbl(Sitem.number)
            Case "cost"
                dblSum = dblSum + CDbl(Sitem.cost) * CDbl(Sitem.number)
        End Select
    Next
    getSumOfBill = dblSum
ending:
    
End Function
Function getPlaceOfBill(strId() As String) As String
'�ꗗ�\�p����ꖼ��Ԃ�
    Dim customer As Customers
    Dim i As Long
    Dim strName As String
    
    Call getCustomer(strId(0), customer)
    getPlaceOfBill = customer.place
    If UBound(strId) = 0 Then GoTo ending
    strName = getPlaceOfBill
    For i = 1 To UBound(strId)
        Call getCustomer(strId(i), customer)
        If strName <> customer.place Then strName = strName & "�A��": Exit For
    Next
    getPlaceOfBill = strName
ending:
End Function

Function getTenantBillSum(strBillDateOnStrs As String, _
                          strBillTypes As String, _
                          strTenantCode As String) As SumList
'�e�i���g�̔���グ�ꗗ��Ԃ�
    Dim strSettleIds As String
    Dim strSettleIdsOnTcode As String
    Dim strId() As String
    Dim strCustomerId() As String
    Dim i As Long
    Dim customer As Customers
    Dim strPlace As String
    Dim shtMy As Worksheet
    Set shtMy = Sheets("���ύ�")
    
'�����Ɉ�v���錈��ID����������
    strSettleIds = getIDsBillTypeAndBillDateFromSettleItem(strBillTypes, strBillDateOnStrs)
    strSettleIdsOnTcode = getTableDatas(shtMy, strTenantCode, SettleArticles_tenant_code_COL, SettleArticles_id_COL)
    strSettleIds = strSettleIds & " " & strSettleIdsOnTcode
    strSettleIds = Trim(strSettleIds)
    If strSettleIds = "" Then GoTo ending
    strId = Split(strSettleIds, " ")
    strSettleIds = DuplicationDraw(strId)
'�f�[�^�Q�̎擾
    ReDim strCustomerId(UBound(strId))
    For i = 0 To UBound(strId)
        strCustomerId(i) = getTableDatas(shtMy, strId(i), SettleArticles_id_COL, SettleArticles_customer_id_COL)
    Next
    Call DuplicationMerge(strCustomerId)
    Call getCustomer(strCustomerId(0), customer)
    
    getTenantBillSum.claim_name = customer.claim_name
    getTenantBillSum.floor = customer.floor
    getTenantBillSum.place = getPlaceOfBill(strCustomerId)
    getTenantBillSum.tenant_code = customer.tenant_code
    getTenantBillSum.price_without_tax = getSumOfBill(strSettleIds, "price_without_tax")
    getTenantBillSum.price = CStr(Int(CDbl(getTenantBillSum.price_without_tax) * 1.05))
    getTenantBillSum.tax = CStr(CDbl(getTenantBillSum.price) - CDbl(getTenantBillSum.price_without_tax))
    getTenantBillSum.cost = getSumOfBill(strSettleIds, "cost")
    getTenantBillSum.profit = getTenantBillSum.price_without_tax - getTenantBillSum.cost
    getTenantBillSum.BillType = strBillTypes
ending:
    Set shtMy = Nothing
 End Function

Function getMaruhiroBillSum(strBillDateOnStrs As String) As SumList
'�ۍL�̔���グ�ꗗ��Ԃ�
    Dim strSettleIds As String
    Const BillType As String = "�[�i�`�["
    
    strSettleIds = getIDsBillTypeAndBillDateFromSettleItem(BillType, strBillDateOnStrs)
    
    getMaruhiroBillSum.place = Range("STORE_NAME")
    getMaruhiroBillSum.price = getSumOfBill(strSettleIds)
    getMaruhiroBillSum.price_without_tax = getSumOfBill(strSettleIds, "price_without_tax")
    getMaruhiroBillSum.cost = getSumOfBill(strSettleIds, "cost")
    getMaruhiroBillSum.profit = getMaruhiroBillSum.price_without_tax - getMaruhiroBillSum.cost
End Function
Function getIDsBillTypeAndBillDateFromSettleItem(strBillTypes As String, _
                                                 strBillDates As String) As String
'���ύς݃V�[�g���琿�����@����ѐ������̗����ɊY������settliId�S���쐬���ĕԂ�
    Dim shtMy As Worksheet
    Dim shtAccaunt As Worksheet
    Dim i As Long
    Dim strId() As String
    Dim strIdsOnBillType As String
    Dim strIdsOnBillDate As String
    
    Set shtMy = Sheets("���ύ�")
    Set shtAccaunt = Sheets("�e�i���g��������")
    '�w�萿�����@�ɊY������ID���擾����
    strId = Split(strBillTypes, " ")
    For i = 0 To UBound(strId)
        strIdsOnBillType = strIdsOnBillType & getTableDatas(shtMy, _
                                                            strId(i), _
                                                            SettleArticles_bill_type_COL, _
                                                            SettleArticles_id_COL) & " "
    Next
    strIdsOnBillType = Trim(strIdsOnBillType)
    If strIdsOnBillType = "" Then GoTo ending
    '�w�茎�ɊY������ID���擾����
    strId = Split(strBillDates, " ")
    For i = 0 To UBound(strId)
        strIdsOnBillDate = strIdsOnBillDate & getTableDatas(shtMy, _
                                                            strId(i), _
                                                            SettleArticles_bill_date_COL, _
                                                            SettleArticles_id_COL) & " "
    Next
    strIdsOnBillDate = Trim(strIdsOnBillDate)
    If strIdsOnBillDate = "" Then GoTo ending
    '�w�茎�w�萿�����@��id�𒊏o����
    strId() = Split(strIdsOnBillType & " " & strIdsOnBillDate, " ")
    getIDsBillTypeAndBillDateFromSettleItem = DuplicationDraw(strId)
ending:
    Set shtMy = Nothing
    Set shtAccaunt = Nothing
End Function
Function getWorksheetNames(Optional wbkMy As Workbook) As String
'�^����ꂽBook�̃V�[�g����Ԃ�
    Dim i As Long
    Dim lngSheetCount As Long
    Dim strName() As String
    
    If wbkMy Is Nothing Then
        Set wbkMy = ActiveWorkbook
    End If
    lngSheetCount = wbkMy.Sheets.Count
    ReDim strName(lngSheetCount - 1)
    For i = 1 To lngSheetCount
        strName(i - 1) = wbkMy.Sheets(i).name
    Next
    getWorksheetNames = Join(strName, ":")
    Set wbkMy = Nothing
End Function
Function getZaikoNum(strItemId As String) As String
'�݌ɐ���Ԃ�(����܂���\�L��p�j
    Dim dblNum As Double
    dblNum = CDbl(getNumOfStock(strItemId))
    If dblNum <= 0 Then
        getZaikoNum = "����܂���"
    Else
        getZaikoNum = CStr(dblNum)
    End If
End Function
Function getNoRegistItemIDs() As String
'JAN CODE�������Ȃ�ItemId��Ԃ�
    Dim shtMy As Worksheet
    Dim strIDs(1) As String
    Dim strId() As String
    
    Set shtMy = Worksheets("�i��")
    strIDs(0) = getRegistItemIDs
    strIDs(1) = getTableDatas(shtMy, "*", Articles_id_COL, Articles_id_COL)
    strId = Split(Join(strIDs))
    getNoRegistItemIDs = DuplicationMerge(strId)
    Set shtMy = Nothing
End Function
Function getRegistItemIDs() As String
'JAN CODE������ItemId��Ԃ�
    Dim shtMy As Worksheet
    Dim strJanKey As String
    
    Set shtMy = Worksheets("�i��")
    strJanKey = "1* 2* 3* 4* 5* 6* 7* 8* 9*"
    getRegistItemIDs = getTableDatas(shtMy, strJanKey, Articles_JAN_code_COL, Articles_id_COL)
    Set shtMy = Nothing
End Function
Function getNoRegistItemIDsFromItemName(strName As String, _
                                        strCallName() As String, _
                                        callName As String) As String
'�i���A���[�J�[�Ăі�����JAN�R�[�h�o�^����Ă��Ȃ�item��ID��Ԃ�
'�i���͈�A�Ăі��͕����̖��O���󂯎��
'��ɕi���ύX�̍ۂ�produt num�̃��X�g���쐬���邽�߂Ɏg�p����
'�z��̌��������ꍇ�ł���̔z���n�����Ɓ@��) strCallName = split("")
    Dim ItemIDs(2) As String
    Dim varCallName As Variant
    Dim strState As String
    Dim i As Long
    Dim strId() As String
    
    ItemIDs(0) = getNoRegistItemIDs
    ItemIDs(1) = getItemIDsFromItemName(strName)
    If callName = "" Then
        For Each varCallName In strCallName
            strState = getItemIdsFromMakerCallName(CStr(varCallName))
            If Not strState = "" Then ItemIDs(2) = ItemIDs(2) & " " & strState
        Next
        ItemIDs(2) = Trim(ItemIDs(2))
    Else
        ItemIDs(2) = getItemIdsFromMakerCallName(callName)
    End If
    strState = ""
    If Not ItemIDs(0) = "" Then strState = Trim(ItemIDs(0))
    If Not ItemIDs(1) = "" Then
        strState = strState & " " & Trim(ItemIDs(1))
    End If
    If Not ItemIDs(2) = "" Then
        strState = strState & " " & Trim(ItemIDs(2))
    End If
    strId = Split(strState)
    getNoRegistItemIDsFromItemName = DuplicationDraw(strId)
End Function
Function getBillTypes() As String
    '�[�i�`�[�ȊO�̐����^�C�v��Ԃ�
    Dim rng As Range
    Dim bill_type As Variant
    
    Set rng = Range("bill_type_list")
    For Each bill_type In rng
        Select Case bill_type
        Case "�[�i�`�["
        Case Else
            getBillTypes = getBillTypes & " " & bill_type
        End Select
    Next
    getBillTypes = Trim(getBillTypes)
End Function
Function getSettleItemList(strBillDate As String, SettleItem() As SettleArticles) As Boolean
'���ύσ��X�g����strBillDate�ɓ��Ă͂܂�f�[�^���X�g�𒊏o����
    Dim sht As Worksheet
    Dim data As Range
    Dim rows As Long
    Dim i As Long, j As Long
    
    getSettleItemList = False
    Set sht = ActiveWorkbook.Worksheets("���ύ�")
    Set data = getDataRange(sht, 15)
    rows = data.rows.Count
    ReDim SettleItem(rows - 1)
    j = 0
    For i = 1 To rows
        If data.rows(i).Cells(1, 15) Like strBillDate Then
            SettleItem(j).id = data.rows(i).Cells(1, 1)
            SettleItem(j).buy_article_id = data.rows(i).Cells(1, 2)
            SettleItem(j).stock_article_id = data.rows(i).Cells(1, 3)
            SettleItem(j).item_id = data.rows(i).Cells(1, 4)
            SettleItem(j).customer_id = data.rows(i).Cells(1, 5)
            SettleItem(j).cost = data.rows(i).Cells(1, 6)
            SettleItem(j).item_price_without_tax = data.rows(i).Cells(1, 7)
            SettleItem(j).item_price = data.rows(i).Cells(1, 8)
            SettleItem(j).number = data.rows(i).Cells(1, 9)
            SettleItem(j).sum = data.rows(i).Cells(1, 10)
            SettleItem(j).bill_type = data.rows(i).Cells(1, 11)
            SettleItem(j).tenant_code = data.rows(i).Cells(1, 12)
            SettleItem(j).delivery_date = data.rows(i).Cells(1, 13)
            SettleItem(j).settle_date = data.rows(i).Cells(1, 14)
            SettleItem(j).bill_date = data.rows(i).Cells(1, 15)
            j = j + 1
        End If
    Next
    If j > 0 Then
        j = j - 1
        getSettleItemList = True
    End If
    ReDim Preserve SettleItem(j)
End Function
Sub getCustomerList(customer() As Customers)
'����悩��f�[�^���X�g���擾����
    Dim sht As Worksheet
    Dim rng As Range
    Dim rows As Long
    Dim i As Long
    
    Set sht = Worksheets("�����")
    Set rng = getDataRange(sht, 8)
    rows = rng.rows.Count
    ReDim customer(rows - 1)
    For i = 1 To rows
        customer(i - 1).id = rng.rows(i).Cells(1, 1)
        customer(i - 1).site = rng.rows(i).Cells(1, 2)
        customer(i - 1).floor = rng.rows(i).Cells(1, 3)
        customer(i - 1).place = rng.rows(i).Cells(1, 4)
        customer(i - 1).claim_name = rng.rows(i).Cells(1, 5)
        customer(i - 1).tenant_code = rng.rows(i).Cells(1, 6)
        customer(i - 1).A_table = rng.rows(i).Cells(1, 7)
        customer(i - 1).bill_type = rng.rows(i).Cells(1, 8)
    Next
End Sub
Function getMaruhiroTotalPrice(da() As DeliveryAccount) As Long
    Dim i As Long
    
    For i = 0 To UBound(da)
        getMaruhiroTotalPrice = getMaruhiroTotalPrice + CLng(da(i).price * da(i).number)
    Next
End Function
Function getMaruhiroTotalCost(da() As DeliveryAccount) As Long
    Dim i As Long
    
    For i = 0 To UBound(da)
        getMaruhiroTotalCost = getMaruhiroTotalCost + CLng(da(i).cost * da(i).number)
    Next
End Function
Sub getItemList(item() As Articles)
'�i�ڈꗗ��z��Ŏ擾����
    Dim shtItem As Worksheet
    Dim rngItem As Range
    Dim rows As Long
    Dim i As Long, j As Long
    
    Set shtItem = Worksheets("�i��")
    Set rngItem = getDataRange(shtItem, Articles_entry_date_COL)
    rows = rngItem.rows.Count
    ReDim item(rows - 1)
    For i = 1 To rows
        item(i - 1).id = rngItem.rows(i).Cells(1, 1)
        item(i - 1).category = rngItem.rows(i).Cells(1, 2)
        item(i - 1).name = rngItem.rows(i).Cells(1, 3)
        item(i - 1).product_number = rngItem.rows(i).Cells(1, 4)
        item(i - 1).maker_id = rngItem.rows(i).Cells(1, 5)
        item(i - 1).fujibil_code = rngItem.rows(i).Cells(1, 6)
        item(i - 1).JAN_CODE = rngItem.rows(i).Cells(1, 7)
        item(i - 1).cost = rngItem.rows(i).Cells(1, 8)
        item(i - 1).price_without_tax = rngItem.rows(i).Cells(1, 9)
        item(i - 1).tax = rngItem.rows(i).Cells(1, 10)
        item(i - 1).price = rngItem.rows(i).Cells(1, 11)
        item(i - 1).trader_id = rngItem.rows(i).Cells(1, 12)
        item(i - 1).entry_date = rngItem.rows(i).Cells(1, 13)
    Next
End Sub
Function getMakerName(maker() As makers, item As Articles) As String
    'item���烁�[�J�[����Ԃ�
    Dim i As Long
    getMakerName = "no maker"
    For i = 0 To UBound(maker)
        If maker(i).id = item.maker_id Then
            getMakerName = maker(i).call_name
            Exit Function
        End If
    Next
End Function
Function getTenantTotalPrice(ta() As TenantAccaunts) As Long
    Dim i As Long
    
    For i = 0 To UBound(ta)
        getTenantTotalPrice = getTenantTotalPrice + CLng(ta(i).price * ta(i).number)
    Next
End Function
Function getTenantTotalCost(ta() As TenantAccaunts) As Long
    Dim i As Long
    
    For i = 0 To UBound(ta)
        getTenantTotalCost = getTenantTotalCost + CLng(ta(i).cost * ta(i).number)
    Next
End Function
Sub getMakerList(maker() As makers)
'���[�J�[�̈ꗗ��z��Ŏ擾����
    Dim shtMaker As Worksheet
    Dim rngMaker As Range
    Dim endrow As Long
    Dim data As Variant
    Dim i As Long, j As Long
    
    Set shtMaker = ActiveWorkbook.Worksheets("���[�J�[")
    endrow = getEndRow("a", shtMaker)
    Set rngMaker = getDataRange(shtMaker, 3)
    ReDim maker(rngMaker.rows.Count - 1)
    i = 0
    j = 0
    For Each data In rngMaker
        Select Case i
        Case 0
            maker(j).id = data
            i = i + 1
        Case 1
            maker(j).name = data
            i = i + 1
        Case 2
            maker(j).call_name = data
            i = 0
            j = j + 1
        End Select
    Next
End Sub

