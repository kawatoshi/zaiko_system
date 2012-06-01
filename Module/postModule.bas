Attribute VB_Name = "postModule"
Option Explicit

Function postBuyToStock(BItem As BuyArticles, Sitem As StockArticles) As String
'データを入庫から在庫へ移す
    Dim shtMy As Worksheet
    Set shtMy = ActiveWorkbook.Sheets("在庫")
    Sitem.id = CStr(getMaxNo(shtMy.Columns("a"))) + 1
    Sitem.buy_article_id = BItem.id
    Sitem.item_id = BItem.item_id
    Sitem.cost = BItem.cost
    Sitem.number = BItem.number
    postBuyToStock = "OK"
ending:
    Set shtMy = Nothing
End Function
Function postStockItem(stockItem As StockArticles) As String
'在庫を書き換える
    Dim shtMy As Worksheet
    Dim rngFind As Range
    Dim rngMy As Range
    Dim strRows As String
    
    Set shtMy = ActiveWorkbook.Sheets("在庫")
    Set rngFind = getFindRange(shtMy, StockArticles_id_COL)
    strRows = getKeyData(stockItem.id, rngFind, "row", xlWhole)
    Set rngMy = shtMy.Cells(CLng(strRows), 1)
    
    rngMy.Cells(1, StockArticles_number_COL).Value = stockItem.number
    rngMy.Cells(1, StockArticles_final_delivery_date_COL).Value = Now()
    postStockItem = stockItem.id & "の数量が" & stockItem.number
    Set rngMy = Nothing
    Set shtMy = Nothing
    Set rngFind = Nothing
End Function
Function postStockToDeliveryItem(stockItem As StockArticles, DeliveryItem As DeliveryArticles, _
                                 dblNum As Double, strCustomerId As String, strBilltype As String) As String
'在庫を出庫へ移動する
'dblNumは移動後の在庫数が与えられる
    Dim item As Articles
    Dim dblNumber As Double
    Call getItem(stockItem.item_id, item)
    If dblNum < 0 Then
        dblNumber = 0
    Else
        dblNumber = dblNum
    End If
    dblNumber = CDbl(stockItem.number) - dblNumber
'出庫
    DeliveryItem.id = getDeliveryItemMaxNo
    DeliveryItem.buy_article_id = stockItem.buy_article_id
    DeliveryItem.stock_article_id = stockItem.id
    DeliveryItem.item_id = stockItem.item_id
    DeliveryItem.customer_id = strCustomerId
    DeliveryItem.cost = stockItem.cost
    DeliveryItem.item_price_without_tax = item.price_without_tax
    DeliveryItem.item_price = item.price
    DeliveryItem.number = CStr(dblNumber)
    DeliveryItem.sum = CStr(CDbl(DeliveryItem.item_price) * CDbl(DeliveryItem.number))
    DeliveryItem.bill_type = strBilltype
    DeliveryItem.delivery_date = Now()
'在庫
    stockItem.number = dblNum
    stockItem.final_delivery_date = DeliveryItem.delivery_date
    postStockToDeliveryItem = "在庫数 " & stockItem.number & " ; 出庫数 " & DeliveryItem.number
ending:
End Function
Function postStockToLost(stockItem As StockArticles, LostItem As LostArticles, dblNum As Double) As String
'在庫をロスへ移動する
'dblNumは移動後の在庫数が与えられる
    Dim dblNumber As Double
    If dblNum < 0 Then
        dblNumber = 0
    Else
        dblNumber = dblNum
    End If
    dblNumber = CDbl(stockItem.number) - dblNumber
'ロス
    LostItem.id = getMaxNo(getFindRange(Sheets("ロス"), DeliveryArticles_id_COL)) + 1
    LostItem.buy_article_id = stockItem.buy_article_id
    LostItem.stock_article_id = stockItem.id
    LostItem.item_id = stockItem.item_id
    LostItem.cost = stockItem.cost
    LostItem.number = CStr(dblNumber)
    LostItem.lost_date = Now()
'在庫
    stockItem.number = dblNum
    postStockToLost = "在庫数 " & stockItem.number & " ; ロス数 " & LostItem.number
ending:
End Function

Function postDeliveryToStock(DeliveryItem As DeliveryArticles, Sitem As StockArticles, dblNumber As Double) As String
'出庫したものの返品処理
'idに同じものがない場合には新規で入力。
'同じものがある場合には数量を足す
    Dim strState As String
    
    If dblNumber <= 0 Then postDeliveryToStock = "0以下は入力出来ません ERROR": GoTo ending
    If CDbl(DeliveryItem.number) < dblNumber Then postDeliveryToStock = "数が多すぎます ERROR": GoTo ending
    strState = getStockItem(DeliveryItem.stock_article_id, Sitem)
    If Not strState Like "OK" Then
        Sitem.id = DeliveryItem.stock_article_id
        Sitem.buy_article_id = DeliveryItem.buy_article_id
        Sitem.item_id = DeliveryItem.item_id
        Sitem.cost = DeliveryItem.cost
        Sitem.number = CStr(dblNumber)
    Else
        Sitem.number = CStr(CDbl(Sitem.number) + dblNumber)
    End If
        DeliveryItem.number = CStr(CDbl(DeliveryItem.number) - dblNumber)
        DeliveryItem.sum = CStr(CDbl(DeliveryItem.item_price) * CDbl(DeliveryItem.number))
        postDeliveryToStock = "出庫" & DeliveryItem.number & " : 返品数" & Sitem.number
ending:
End Function

Function postDeliveryList(Ditem As DeliveryArticles, dlist As DeliveryList) As String
'出庫から出庫リストデータを作成する
    Dim shtItem As Worksheet
    Dim shtMaker As Worksheet
    Dim shtCustomer As Worksheet
    Dim strContents As String
    Dim strId As String
    Set shtItem = Worksheets("品目")
    Set shtMaker = Worksheets("メーカー")
    Set shtCustomer = Worksheets("取引先")
    
    postDeliveryList = "postDeliveryList NG"
    'id
    dlist.id = Ditem.id
    '種類
    dlist.type_name = getTableDatas(shtItem, Ditem.item_id, Articles_id_COL, Articles_name_COL)
    '品名
    dlist.item_name = getHinmoku(Ditem.item_id)
    dlist.customer_name = strContents
    '販売価格
    dlist.item_price = Ditem.item_price
    '数量
    dlist.number = Ditem.number
    '金額
    dlist.sum = Ditem.sum
    '取引先
    strContents = getTableDatas(shtCustomer, Ditem.customer_id, Customers_id_COL, Customers_place_COL)
    dlist.customer_name = strContents
    '出庫日
    dlist.delivery_date = Ditem.delivery_date
    postDeliveryList = "postDeliveryList OK"
ending:
    Set shtItem = Nothing
    Set shtMaker = Nothing
    Set shtCustomer = Nothing
End Function

Function postDeliveryToSettleItem(Ditem() As DeliveryArticles, SettleItem() As SettleArticles, _
                                  dteClosingDate As Date)
'決済を行うデータをSettleItemに書き込む
'除外された元データのidにはstayを入れ込む
    Dim dteMy As Date
    Dim i As Long, j As Long
    Dim shtCustomer As Worksheet
    Dim strTcode As String
    Set shtCustomer = Sheets("取引先")
    postDeliveryToSettleItem = "postDeliveryToSettleItem NG"
    j = 0
    ReDim SettleItem(UBound(Ditem))
    For i = 1 To UBound(Ditem)
        If dteClosingDate > Ditem(i).delivery_date Then
            strTcode = getTableDatas(shtCustomer, Ditem(i).customer_id, Customers_id_COL, Customers_tenant_code_COL)
            j = j + 1
            SettleItem(j).id = Ditem(i).id
            SettleItem(j).buy_article_id = Ditem(i).buy_article_id
            SettleItem(j).stock_article_id = Ditem(i).stock_article_id
            SettleItem(j).item_id = Ditem(i).item_id
            SettleItem(j).customer_id = Ditem(i).customer_id
            SettleItem(j).cost = Ditem(i).cost
            SettleItem(j).item_price_without_tax = Ditem(i).item_price_without_tax
            SettleItem(j).item_price = Ditem(i).item_price
            SettleItem(j).number = Ditem(i).number
            SettleItem(j).sum = Ditem(i).sum
            SettleItem(j).bill_type = Ditem(i).bill_type
            SettleItem(j).tenant_code = strTcode
            SettleItem(j).delivery_date = Ditem(i).delivery_date
            SettleItem(j).settle_date = Now
            SettleItem(j).bill_date = getBilldateOnStr(dteClosingDate)
        Else
            Ditem(i).id = "stay"
        End If
    Next
    ReDim Preserve SettleItem(j)
    postDeliveryToSettleItem = "postDeliveryToSettleItem OK"
    Set shtCustomer = Nothing
End Function

Function postTmpSettleItem(Sitem As SettleArticles) As TmpSettleArticles
'引数からTmp決済データを取得する
    Dim shtItem As Worksheet
    Dim shtCustomer As Worksheet
    Dim rngMy As Range
    Dim customer As Customers
    Dim item As Articles
    Dim maker As makers
    Set shtItem = Sheets("品目")
    Set shtCustomer = Sheets("取引先")
    Call getCustomer(Sitem.customer_id, customer)
    Call getItem(Sitem.item_id, item)
    Call getMaker(item.maker_id, maker)
    
    postTmpSettleItem.id = Sitem.id
    postTmpSettleItem.delivery_date = Sitem.delivery_date
    postTmpSettleItem.customer = customer.place
    postTmpSettleItem.claim_name = customer.claim_name
    postTmpSettleItem.maker = maker.call_name
    postTmpSettleItem.item_name = item.name
    postTmpSettleItem.item = item.product_number
    postTmpSettleItem.item_price = Sitem.item_price_without_tax
    postTmpSettleItem.number = Sitem.number
    postTmpSettleItem.sum = postTmpSettleItem.item_price * postTmpSettleItem.number
    postTmpSettleItem.bill_type = Sitem.bill_type
ending:
    Set shtItem = Nothing
    Set shtCustomer = Nothing
End Function

Function postStocklist(item As Articles) As StockList
'ストックリストデータを返す
    postStocklist.id = item.id
    postStocklist.type_name = item.name
    postStocklist.item_name = getHinmoku(item.id)
    postStocklist.cost = item.cost
    postStocklist.number = getNumOfStock(item.id)
    postStocklist.sum = getSumOfStock(item.id)
'    postStocklist.sum = CStr(CDbl(postStocklist.cost) * CDbl(postStocklist.number))
    postStocklist.item_price = item.price
    postStocklist.stock_date = getNewerBuyDate(item.id)
    postStocklist.delivery_date = getNewerDeliveryDate(item.id)
End Function
Function postBuyItemToBuyList(Buyitem As BuyArticles) As BuyList
'入庫データを返す
    Dim item As Articles
    Call getItem(Buyitem.item_id, item)
    
    postBuyItemToBuyList.id = Buyitem.id
    postBuyItemToBuyList.item_id = item.id
    postBuyItemToBuyList.name = item.name
    postBuyItemToBuyList.product_number = getHinmoku(item.id)
    postBuyItemToBuyList.cost = Buyitem.cost
    postBuyItemToBuyList.number = Buyitem.number
    postBuyItemToBuyList.stock_date = Buyitem.in_stock_date
End Function
Function postTmpSettleItemToDeliveryAccount(TmpSItem As TmpSettleArticles) As DeliveryAccount
'TmpSettleArticlesデータよりDeliveryAccountを作成する
    Dim Sitem As SettleArticles
    Dim customer As Customers
    Call getSettleItem(TmpSItem.id, Sitem)
    Call getCustomer(Sitem.customer_id, customer)
    
    postTmpSettleItemToDeliveryAccount.delivery_date = TmpSItem.delivery_date
    postTmpSettleItemToDeliveryAccount.customer = customer.floor & "  " & TmpSItem.customer
    postTmpSettleItemToDeliveryAccount.maker = TmpSItem.maker
    postTmpSettleItemToDeliveryAccount.item_name = TmpSItem.item_name
    postTmpSettleItemToDeliveryAccount.produnt_number = TmpSItem.item
    postTmpSettleItemToDeliveryAccount.price = TmpSItem.item_price
    postTmpSettleItemToDeliveryAccount.number = TmpSItem.number
    postTmpSettleItemToDeliveryAccount.sum = TmpSItem.sum
End Function
Function postDeliveryItemToDeliveryAccount(Ditem As DeliveryArticles) As DeliveryAccount
'DeliveryArticlesよりDeliveryAccountを作成する
    Dim customer As Customers
    Dim item As Articles
    Dim maker As makers
    Call getCustomer(Ditem.customer_id, customer)
    Call getItem(Ditem.item_id, item)
    Call getMaker(item.maker_id, maker)
    
    postDeliveryItemToDeliveryAccount.delivery_date = Ditem.delivery_date
    postDeliveryItemToDeliveryAccount.customer = customer.floor & "  " & customer.place
    postDeliveryItemToDeliveryAccount.maker = maker.call_name
    postDeliveryItemToDeliveryAccount.item_name = item.name
    postDeliveryItemToDeliveryAccount.produnt_number = item.product_number
    postDeliveryItemToDeliveryAccount.price = Ditem.item_price_without_tax
    postDeliveryItemToDeliveryAccount.number = Ditem.number
    postDeliveryItemToDeliveryAccount.sum = postDeliveryItemToDeliveryAccount.price * postDeliveryItemToDeliveryAccount.number
End Function
Function postDistributerAccount(strSID As String) As DistributerAccount
'不二ビル明細データを書き換える
    Dim item As Articles
    Dim customer As Customers
    Dim settle As SettleArticles
    Dim maker As makers
    
    Call getSettleItem(strSID, settle)
    Call getCustomer(settle.customer_id, customer)
    Call getItem(settle.item_id, item)
    Call getMaker(item.maker_id, maker)
    
    postDistributerAccount.id = strSID
    postDistributerAccount.delivery_date = settle.delivery_date
    postDistributerAccount.bill_date = settle.bill_date
    postDistributerAccount.item_name = item.name
    postDistributerAccount.maker_name = maker.call_name
    postDistributerAccount.product_number = item.product_number
    postDistributerAccount.customer_name = customer.claim_name
    postDistributerAccount.floor = customer.floor
    postDistributerAccount.claim_name = customer.place
    postDistributerAccount.bill_type = settle.bill_type
    postDistributerAccount.item_price = settle.item_price
    postDistributerAccount.number = settle.number
    postDistributerAccount.sum_of_price = settle.sum
    
End Function

Function postTenantAccaunt(strAccauntId As String) As TenantAccaunts
    Dim customer As Customers
    Dim maker As makers
    Dim item As Articles
    Dim Sitem As SettleArticles
    Call getSettleItem(strAccauntId, Sitem)
    Call getCustomer(Sitem.customer_id, customer)
    Call getItem(Sitem.item_id, item)
    Call getMaker(item.maker_id, maker)
    postTenantAccaunt.delivery_date = Sitem.delivery_date
    postTenantAccaunt.tenant_code = Sitem.tenant_code
    postTenantAccaunt.floor = customer.floor
    postTenantAccaunt.place = customer.place
    postTenantAccaunt.maker = maker.call_name
    postTenantAccaunt.item_name = item.name
    postTenantAccaunt.product_name = item.product_number
    postTenantAccaunt.price = Sitem.item_price_without_tax
    postTenantAccaunt.number = Sitem.number
    postTenantAccaunt.sum = postTenantAccaunt.price * postTenantAccaunt.number
End Function
Function postSettleItemsToDeliveryAccount(data() As DeliveryAccount, _
                                          settle_item_list() As SettleArticles) As String
    '決済済データから丸広請求内訳データを作成する
    Dim maker() As makers
    Dim item() As Articles
    Dim settle_item() As SettleArticles
    Dim customer() As Customers
    Dim i As Long
    Dim itm As Articles
    Dim mkr As makers
    
    Call getMakerList(maker)
    Call getItemList(item)
    Call itemSort(item, LBound(item), UBound(item))
    Call getCustomerList(customer)
    If findSettleItemsByBillType(settle_item, settle_item_list, "納品伝票") = True Then
        ReDim data(UBound(settle_item))
        For i = 0 To UBound(settle_item)
            If findItem(itm, item(), settle_item(i).item_id) = True Then
                data(i).produnt_number = itm.product_number
                data(i).item_name = itm.name
                data(i).price = itm.price_without_tax
                data(i).maker = getMakerName(maker(), itm)
            End If
            data(i).delivery_date = settle_item(i).delivery_date
            data(i).customer = findCustomerForDeliveryAccount(customer, settle_item(i).customer_id)
            data(i).cost = settle_item(i).cost
            data(i).number = settle_item(i).number
            data(i).sum = data(i).price * data(i).number
        Next
    End If
End Function
Function postSettleItemsToTenantAccaunt(data() As TenantAccaunts, _
                                        settle_item_list() As SettleArticles, _
                                        strBilltype As String) As Boolean
    '決済済データからテナント請求内訳データを作成する
    Dim maker() As makers
    Dim item() As Articles
    Dim settle_item() As SettleArticles
    Dim customer() As Customers
    Dim i As Long
    Dim itm As Articles
    Dim mkr As makers
    Dim cus As Customers
    
    postSettleItemsToTenantAccaunt = False
    Call getMakerList(maker)
    Call getItemList(item)
    Call itemSort(item, LBound(item), UBound(item))
    Call getCustomerList(customer)
    If findSettleItemsByBillType(settle_item, settle_item_list, strBilltype) = True Then
        Call SettleItemSortByTenantCode(settle_item, LBound(settle_item), UBound(settle_item))
        ReDim data(UBound(settle_item))
        For i = 0 To UBound(settle_item)
            If findItem(itm, item(), settle_item(i).item_id) = True Then
                data(i).product_name = itm.product_number
                data(i).item_name = itm.name
                data(i).maker = getMakerName(maker(), itm)
            End If
            If findCustomer(cus, customer(), settle_item(i).customer_id) = True Then
                data(i).claim_name = cus.claim_name
                data(i).place = cus.place
                data(i).floor = cus.floor
            End If
            data(i).delivery_date = settle_item(i).delivery_date
            data(i).tenant_code = settle_item(i).tenant_code
            data(i).price = settle_item(i).item_price_without_tax
            data(i).cost = settle_item(i).cost
            data(i).number = settle_item(i).number
            data(i).sum = CLng(data(i).price) * CLng(data(i).number)
            data(i).bill_type = settle_item(i).bill_type
        Next
        postSettleItemsToTenantAccaunt = True
    Else
        Exit Function
    End If
End Function

