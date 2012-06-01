Attribute VB_Name = "putModule"
Option Explicit
Function putPayment(rngFind As Range, PaymentData As Payments) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("支払い")
    If Err.number <> 0 Then putPayment = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, Payments_id_COL).Value = PaymentData.id
    rngFind.Cells(1, Payments_trader_id_COL).Value = PaymentData.trader_id
    rngFind.Cells(1, Payments_date_COL).Value = PaymentData.date
    rngFind.Cells(1, Payments_sum_COL).Value = PaymentData.sum
    rngFind.Cells(1, Payments_tax_COL).Value = PaymentData.tax
    putPayment = "OK"
ending:
    Set shtMy = Nothing
End Function
Sub putTenantAccauntsSubTotal(price As Long, sum As SumList, putRange As Range)
    Dim price_with_tax As Long
    Dim p_tax As Long
    
    price_with_tax = tax(price)
    p_tax = price_with_tax - price
    sum.price = sum.price + price_with_tax
    sum.price_without_tax = sum.price_without_tax + price
    sum.tax = sum.tax + p_tax
    Call putTATotal(price, p_tax, price_with_tax, putRange)
End Sub
Sub putTenantAccaunts(accaunt() As TenantAccaunts, _
                      price As Long, _
                      putRange As Range)
    Dim k As Long
    
    price = 0
    For k = 0 To UBound(accaunt)
        price = price + accaunt(k).sum
        Call putTenantAccaunt(putRange, accaunt(k))
        Set putRange = putRange.Offset(1, 0)
    Next
End Sub
Sub putTATotal(price As Long, tax As Long, p_with_tax As Long, putRange As Range)
    Dim data(2, 2) As String
    Dim i As Long, j As Long
    
    data(0, 0) = "小   計"
    data(0, 1) = "税抜き"
    data(0, 2) = CStr(price)
    data(1, 0) = "消費税"
    data(1, 1) = ""
    data(1, 2) = CStr(tax)
    data(2, 0) = "小   計"
    data(2, 1) = "税込み"
    data(2, 2) = CStr(p_with_tax)
    For i = 0 To 2
        For j = 0 To 2
            putRange.Offset(i, j + 7).Value = data(i, j)
        Next
    Next
    Set putRange = putRange.Offset(i, 0)
End Sub
Sub putBillTA(accaunt As TenantAccaunts, price As Long, putRange As Range)
    Set putRange = putRange.Resize(1, 1)
    putRange.Offset(0, 0).Value = accaunt.claim_name
    putRange.Offset(0, 1).Value = accaunt.floor
    putRange.Offset(0, 2).Value = accaunt.place
    putRange.Offset(0, 3).Value = accaunt.tenant_code
    putRange.Offset(0, 4).Value = price
    Set putRange = putRange.Offset(1, 0)
End Sub
Sub putSaleTA(accaunt As TenantAccaunts, _
              price As String, _
              p_with_tax As String, _
              cost As String, _
              putRange As Range)
    Set putRange = putRange.Resize(1, 1)
    putRange.Offset(0, 0).Value = accaunt.claim_name
    putRange.Offset(0, 1).Value = accaunt.floor
    putRange.Offset(0, 2).Value = accaunt.place
    putRange.Offset(0, 3).Value = price
    putRange.Offset(0, 4).Value = cost
    If IsNumeric(price) Then
        putRange.Offset(0, 5).Value = CLng(price) - CLng(cost)
    End If
    putRange.Offset(0, 6).Value = accaunt.bill_type
    putRange.Offset(0, 7).Value = p_with_tax
    Set putRange = putRange.Offset(1, 0)
End Sub
Function putProtectList(strName As String, strMode As String) As String
    Dim strCol As String
    Dim lngCol As Long
    Dim shtMy As Worksheet
    Set shtMy = ActiveWorkbook.Sheets("tmp")
    putProtectList = "nil"
    Select Case strMode
        Case "p"
            strCol = "e:e"
            lngCol = 5
        Case "f"
            strCol = "f:f"
            lngCol = 6
    End Select
    shtMy.Cells(getEndRow(strCol, shtMy) + 1, lngCol).Value = strName
    putProtectList = "OK"
    Set shtMy = Nothing
End Function
Function putMaker(rngFind As Range, makerData As makers) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("メーカー")
    If Err.number <> 0 Then putMaker = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, Makers_id_COL).Value = makerData.id
    rngFind.Cells(1, Makers_name_COL).Value = makerData.name
    rngFind.Cells(1, Makers_call_name_COL).Value = makerData.call_name
    putMaker = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putTrader(rngFind As Range, TraderData As Traders) As String
    Dim shtMy As Worksheet
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("取引業者")
    On Error GoTo 0
    If Err.number <> 0 Then putTrader = "sheetERROR": GoTo ending
    rngFind.Cells(1, Traders_id_COL).Value = TraderData.id
    rngFind.Cells(1, Traders_company_name_COL).Value = TraderData.company_name
    rngFind.Cells(1, Traders_tel_COL).Value = TraderData.tel
    rngFind.Cells(1, Traders_address_COL).Value = TraderData.address
    rngFind.Cells(1, Traders_person_name_COL).Value = TraderData.person_name
    putTrader = "OK"
ending:
    Set shtMy = Nothing
End Function
Function putCustomer(rngFind As Range, CustomerData As Customers) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("取引先")
    If Err.number <> 0 Then putCustomer = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, Customers_id_COL).Value = CustomerData.id
    rngFind.Cells(1, Customers_site_COL).Value = CustomerData.site
    rngFind.Cells(1, Customers_floor_COL).Value = CustomerData.floor
    rngFind.Cells(1, Customers_place_COL).Value = CustomerData.place
    rngFind.Cells(1, Customers_claim_name_COL).Value = CustomerData.claim_name
    rngFind.Cells(1, Customers_tenant_code_COL).Value = CustomerData.tenant_code
    rngFind.Cells(1, Customers_A_table_COL).Value = CustomerData.A_table
    rngFind.Cells(1, Customers_bill_type_COL).Value = CustomerData.bill_type
    putCustomer = "OK"
ending:
    Set shtMy = Nothing
End Function
Function putBuyItem(rngFind As Range, BuyItemData As BuyArticles) As String
    Dim shtMy As Worksheet
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("入庫")
    On Error GoTo 0
    If Err.number <> 0 Then putBuyItem = "sheetERROR": GoTo ending
    rngFind.Cells(1, BuyArticles_id_COL).Value = BuyItemData.id
    rngFind.Cells(1, BuyArticles_item_id_COL).Value = BuyItemData.item_id
    rngFind.Cells(1, BuyArticles_trader_id_COL).Value = BuyItemData.trader_id
    rngFind.Cells(1, BuyArticles_cost_COL).Value = BuyItemData.cost
    rngFind.Cells(1, BuyArticles_number_COL).Value = BuyItemData.number
    rngFind.Cells(1, BuyArticles_in_stock_date_COL).Value = BuyItemData.in_stock_date
    putBuyItem = "OK"
ending:
    Set shtMy = Nothing
End Function
Function putStockItem(rngFind As Range, StockItemData As StockArticles) As String
    Dim shtMy As Worksheet
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("在庫")
    On Error GoTo 0
    If Err.number <> 0 Then putStockItem = "sheetERROR": GoTo ending
    rngFind.Cells(1, StockArticles_id_COL).Value = StockItemData.id
    rngFind.Cells(1, StockArticles_buy_article_id_COL).Value = StockItemData.buy_article_id
    rngFind.Cells(1, StockArticles_item_id_COL).Value = StockItemData.item_id
    rngFind.Cells(1, StockArticles_cost_COL).Value = StockItemData.cost
    rngFind.Cells(1, StockArticles_number_COL).Value = StockItemData.number
    If StockItemData.final_delivery_date <> 0 Then _
        rngFind.Cells(1, StockArticles_final_delivery_date_COL).Value = StockItemData.final_delivery_date
    rngFind.Cells(1, StockArticles_receipt_article_id_COL).Value = StockItemData.receipt_article_id
    putStockItem = "OK"
ending:
    Set shtMy = Nothing
End Function
    Function putDeliveryItem(rngFind As Range, DeliveryItemData As DeliveryArticles) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("出庫")
    If Err.number <> 0 Then putDeliveryItem = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, DeliveryArticles_id_COL).Value = DeliveryItemData.id
    rngFind.Cells(1, DeliveryArticles_buy_article_id_COL).Value = DeliveryItemData.buy_article_id
    rngFind.Cells(1, DeliveryArticles_stock_article_id_COL).Value = DeliveryItemData.stock_article_id
    rngFind.Cells(1, DeliveryArticles_item_id_COL).Value = DeliveryItemData.item_id
    rngFind.Cells(1, DeliveryArticles_customer_id_COL).Value = DeliveryItemData.customer_id
    rngFind.Cells(1, DeliveryArticles_cost_COL).Value = DeliveryItemData.cost
    rngFind.Cells(1, DeliveryArticles_item_price_without_tax_COL).Value = DeliveryItemData.item_price_without_tax
    rngFind.Cells(1, DeliveryArticles_item_price_COL).Value = DeliveryItemData.item_price
    rngFind.Cells(1, DeliveryArticles_number_COL).Value = DeliveryItemData.number
    rngFind.Cells(1, DeliveryArticles_sum_COL).Value = DeliveryItemData.sum
    rngFind.Cells(1, DeliveryArticles_bill_type_COL).Value = DeliveryItemData.bill_type
    rngFind.Cells(1, DeliveryArticles_delivery_date_COL).Value = DeliveryItemData.delivery_date
    putDeliveryItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putLostItem(rngFind As Range, LostItemData As LostArticles) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("ロス")
    If Err.number <> 0 Then putLostItem = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, LostArticles_id_COL).Value = LostItemData.id
    rngFind.Cells(1, LostArticles_buy_article_id_COL).Value = LostItemData.buy_article_id
    rngFind.Cells(1, LostArticles_stock_article_id_COL).Value = LostItemData.stock_article_id
    rngFind.Cells(1, LostArticles_item_id_COL).Value = LostItemData.item_id
    rngFind.Cells(1, LostArticles_cost_COL).Value = LostItemData.cost
    rngFind.Cells(1, LostArticles_number_COL).Value = LostItemData.number
    rngFind.Cells(1, LostArticles_lost_date_COL).Value = LostItemData.lost_date
    rngFind.Cells(1, LostArticles_note_COL).Value = LostItemData.note
    putLostItem = "OK"
ending:
    Set shtMy = Nothing
End Function
Function putReturnItem(rngFind As Range, ReturnItemData As ReturnArticles) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("返品履歴")
    If Err.number <> 0 Then putReturnItem = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, ReturnArticles_id_COL).Value = ReturnItemData.id
    rngFind.Cells(1, ReturnArticles_buy_article_id_COL).Value = ReturnItemData.buy_article_id
    rngFind.Cells(1, ReturnArticles_stock_article_id_COL).Value = ReturnItemData.stock_article_id
    rngFind.Cells(1, ReturnArticles_item_id_COL).Value = ReturnItemData.item_id
    rngFind.Cells(1, ReturnArticles_customer_id_COL).Value = ReturnItemData.customer_id
    rngFind.Cells(1, ReturnArticles_cost_COL).Value = ReturnItemData.cost
    rngFind.Cells(1, ReturnArticles_item_price_COL).Value = ReturnItemData.item_price
    rngFind.Cells(1, ReturnArticles_number_COL).Value = ReturnItemData.number
    rngFind.Cells(1, ReturnArticles_return_date_COL).Value = ReturnItemData.return_date
    putReturnItem = "OK"
ending:
    Set shtMy = Nothing
End Function
Function putDeliveryList(rngFind As Range, DeliveryListData As DeliveryList) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("出庫リスト")
    If Err.number <> 0 Then putDeliveryList = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, DeliveryList_id_COL).Value = DeliveryListData.id
    rngFind.Cells(1, DeliveryList_type_name_COL).Value = DeliveryListData.type_name
    rngFind.Cells(1, DeliveryList_item_name_COL).Value = DeliveryListData.item_name
    rngFind.Cells(1, DeliveryList_item_price_COL).Value = DeliveryListData.item_price
    rngFind.Cells(1, DeliveryList_number_COL).Value = DeliveryListData.number
    rngFind.Cells(1, DeliveryList_sum_COL).Value = DeliveryListData.sum
    rngFind.Cells(1, DeliveryList_customer_name_COL).Value = DeliveryListData.customer_name
    rngFind.Cells(1, DeliveryList_delivery_date_COL).Value = DeliveryListData.delivery_date
    putDeliveryList = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putSettleItem(rngFind As Range, SettleItemData As SettleArticles) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("決済済")
    If Err.number <> 0 Then putSettleItem = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, SettleArticles_id_COL).Value = SettleItemData.id
    rngFind.Cells(1, SettleArticles_buy_article_id_COL).Value = SettleItemData.buy_article_id
    rngFind.Cells(1, SettleArticles_stock_article_id_COL).Value = SettleItemData.stock_article_id
    rngFind.Cells(1, SettleArticles_item_id_COL).Value = SettleItemData.item_id
    rngFind.Cells(1, SettleArticles_customer_id_COL).Value = SettleItemData.customer_id
    rngFind.Cells(1, SettleArticles_cost_COL).Value = SettleItemData.cost
    rngFind.Cells(1, SettleArticles_item_price_without_tax_COL).Value = SettleItemData.item_price_without_tax
    rngFind.Cells(1, SettleArticles_item_price_COL).Value = SettleItemData.item_price
    rngFind.Cells(1, SettleArticles_number_COL).Value = SettleItemData.number
    rngFind.Cells(1, SettleArticles_sum_COL).Value = SettleItemData.sum
    rngFind.Cells(1, SettleArticles_bill_type_COL).Value = SettleItemData.bill_type
    rngFind.Cells(1, SettleArticles_tenant_code_COL).Value = SettleItemData.tenant_code
    rngFind.Cells(1, SettleArticles_delivery_date_COL).Value = SettleItemData.delivery_date
    rngFind.Cells(1, SettleArticles_settle_date_COL).Value = SettleItemData.settle_date
    rngFind.Cells(1, SettleArticles_bill_date_COL).Value = SettleItemData.bill_date
    putSettleItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putTmpSettleItem(rngFind As Range, TmpSettleItemData As TmpSettleArticles) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("Tmp決済")
    If Err.number <> 0 Then putTmpSettleItem = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, TmpSettleArticles_id_COL).Value = TmpSettleItemData.id
    rngFind.Cells(1, TmpSettleArticles_delivery_date_COL).Value = TmpSettleItemData.delivery_date
    rngFind.Cells(1, TmpSettleArticles_customer_COL).Value = TmpSettleItemData.customer
    rngFind.Cells(1, TmpSettleArticles_claim_name_COL).Value = TmpSettleItemData.claim_name
    rngFind.Cells(1, TmpSettleArticles_maker_COL).Value = TmpSettleItemData.maker
    rngFind.Cells(1, TmpSettleArticles_item_name_COL).Value = TmpSettleItemData.item_name
    rngFind.Cells(1, TmpSettleArticles_item_COL).Value = TmpSettleItemData.item
    rngFind.Cells(1, TmpSettleArticles_item_price_COL).Value = TmpSettleItemData.item_price
    rngFind.Cells(1, TmpSettleArticles_number_COL).Value = TmpSettleItemData.number
    rngFind.Cells(1, TmpSettleArticles_sum_COL).Value = TmpSettleItemData.sum
    rngFind.Cells(1, TmpSettleArticles_bill_type_COL).Value = TmpSettleItemData.bill_type
    putTmpSettleItem = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putDeliveryAccount(rngFind As Range, DeliveryAccountData As DeliveryAccount) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("丸広請求内訳")
    Set shtMy = ActiveWorkbook.Sheets("丸広承認願")
    If Err.number <> 0 Then putDeliveryAccount = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, DeliveryAccount_delivery_date_COL).Value = DeliveryAccountData.delivery_date
    rngFind.Cells(1, DeliveryAccount_customer_COL).Value = DeliveryAccountData.customer
    rngFind.Cells(1, DeliveryAccount_maker_COL).Value = DeliveryAccountData.maker
    rngFind.Cells(1, DeliveryAccount_item_name_COL).Value = DeliveryAccountData.item_name
    rngFind.Cells(1, DeliveryAccount_produnt_number_COL).Value = DeliveryAccountData.produnt_number
    rngFind.Cells(1, DeliveryAccount_price_COL).Value = DeliveryAccountData.price
    rngFind.Cells(1, DeliveryAccount_number_COL).Value = DeliveryAccountData.number
    rngFind.Cells(1, DeliveryAccount_sum_COL).Value = DeliveryAccountData.sum
    putDeliveryAccount = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putStockList(rngFind As Range, StockListData As StockList) As String
    Dim shtMy As Worksheet
    Dim dteMy As Date
    
    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("在庫リスト")
    If Err.number <> 0 Then putStockList = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, StockList_id_COL).Value = StockListData.id
    rngFind.Cells(1, StockList_type_name_COL).Value = StockListData.type_name
    rngFind.Cells(1, StockList_item_name_COL).Value = StockListData.item_name
    rngFind.Cells(1, StockList_cost_COL).Value = StockListData.cost
    rngFind.Cells(1, StockList_number_COL).Value = StockListData.number
    rngFind.Cells(1, StockList_sum_COL).Value = StockListData.sum
    rngFind.Cells(1, StockList_item_price_COL).Value = StockListData.item_price
    rngFind.Cells(1, StockList_stock_date_COL).Value = StockListData.stock_date
    dteMy = StockListData.delivery_date
    If Not dteMy = 0 Then
        rngFind.Cells(1, StockList_delivery_date_COL).Value = dteMy
    End If
    putStockList = "OK"
ending:
    Set shtMy = Nothing
End Function
Function putBuyList(rngFind As Range, BuyListData As BuyList) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("入庫リスト")
    If Err.number <> 0 Then putBuyList = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, BuyList_id_COL).Value = BuyListData.id
    rngFind.Cells(1, BuyList_item_id_COL).Value = BuyListData.item_id
    rngFind.Cells(1, BuyList_name_COL).Value = BuyListData.name
    rngFind.Cells(1, BuyList_product_number_COL).Value = BuyListData.product_number
    rngFind.Cells(1, BuyList_cost_COL).Value = BuyListData.cost
    rngFind.Cells(1, BuyList_number_COL).Value = BuyListData.number
    rngFind.Cells(1, BuyList_stock_date_COL).Value = BuyListData.stock_date
    putBuyList = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putDistributerAccount(rngFind As Range, DistributerAccountData As DistributerAccount) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("不二ビル明細")
    If Err.number <> 0 Then putDistributerAccount = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, DistributerAccount_id_COL).Value = DistributerAccountData.id
    rngFind.Cells(1, DistributerAccount_delivery_date_COL).Value = DistributerAccountData.delivery_date
    rngFind.Cells(1, DistributerAccount_bill_date_COL).Value = DistributerAccountData.bill_date
    rngFind.Cells(1, DistributerAccount_item_name_COL).Value = DistributerAccountData.item_name
    rngFind.Cells(1, DistributerAccount_maker_name_COL).Value = DistributerAccountData.maker_name
    rngFind.Cells(1, DistributerAccount_product_number_COL).Value = DistributerAccountData.product_number
    rngFind.Cells(1, DistributerAccount_customer_name_COL).Value = DistributerAccountData.customer_name
    rngFind.Cells(1, DistributerAccount_floor_COL).Value = DistributerAccountData.floor
    rngFind.Cells(1, DistributerAccount_claim_name_COL).Value = DistributerAccountData.claim_name
    rngFind.Cells(1, DistributerAccount_bill_type_COL).Value = DistributerAccountData.bill_type
    rngFind.Cells(1, DistributerAccount_item_price_COL).Value = DistributerAccountData.item_price
    rngFind.Cells(1, DistributerAccount_number_COL).Value = DistributerAccountData.number
    rngFind.Cells(1, DistributerAccount_sum_of_price_COL).Value = DistributerAccountData.sum_of_price
    putDistributerAccount = "OK"
ending:
    Set shtMy = Nothing
End Function

Function putTenantAccaunt(rngFind As Range, TenantAccauntData As TenantAccaunts) As String
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("テナント請求内訳")
    If Err.number <> 0 Then putTenantAccaunt = "sheetERROR": GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, TenantAccaunts_delivery_date_COL).Value = TenantAccauntData.delivery_date
    rngFind.Cells(1, TenantAccaunts_tenant_code_COL).Value = TenantAccauntData.tenant_code
    rngFind.Cells(1, TenantAccaunts_floor_COL).Value = TenantAccauntData.floor
    rngFind.Cells(1, TenantAccaunts_place_COL).Value = TenantAccauntData.place
    rngFind.Cells(1, TenantAccaunts_maker_COL).Value = TenantAccauntData.maker
    rngFind.Cells(1, TenantAccaunts_item_name_COL).Value = TenantAccauntData.item_name
    rngFind.Cells(1, TenantAccaunts_product_name_COL).Value = TenantAccauntData.product_name
    rngFind.Cells(1, TenantAccaunts_price_COL).Value = TenantAccauntData.price
    rngFind.Cells(1, TenantAccaunts_number_COL).Value = TenantAccauntData.number
    rngFind.Cells(1, TenantAccaunts_sum_COL).Value = TenantAccauntData.sum
    putTenantAccaunt = "OK"
ending:
    Set shtMy = Nothing
End Function
Sub putSalesList(rngFind As Range, tenant As SumList)
'売上一覧を書き込む
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("売上一覧表")
    If Err.number <> 0 Then GoTo ending
    On Error GoTo 0
    
    rngFind.Cells(1, 1).Value = tenant.claim_name
    rngFind.Cells(1, 2).Value = tenant.floor
    rngFind.Cells(1, 3).Value = tenant.place
    rngFind.Cells(1, 4).Value = tenant.price_without_tax
    rngFind.Cells(1, 5).Value = tenant.cost
    rngFind.Cells(1, 6).Value = tenant.profit
    rngFind.Cells(1, 7).Value = tenant.BillType
    rngFind.Cells(1, 8).Value = tenant.price
ending:
    Set shtMy = Nothing
End Sub
Sub putBillList(rngFind As Range, tenant As SumList)
    Dim shtMy As Worksheet

    On Error Resume Next
    Set shtMy = ActiveWorkbook.Sheets("請求金額一覧表")
    If Err.number <> 0 Then GoTo ending
    On Error GoTo 0
    rngFind.Cells(1, 1).Value = tenant.claim_name
    rngFind.Cells(1, 2).Value = tenant.floor
    rngFind.Cells(1, 3).Value = tenant.place
    rngFind.Cells(1, 4).Value = tenant.tenant_code
    rngFind.Cells(1, 5).Value = tenant.price
ending:
    Set shtMy = Nothing
End Sub


