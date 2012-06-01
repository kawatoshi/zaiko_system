VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuyItem 
   Caption         =   "入庫"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   OleObjectBlob   =   "BuyItem.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "BuyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public text As String

Private Sub btn10_Click()
    txtNumber.text = "10"
    btnOK.SetFocus
End Sub

Private Sub btn20_Click()
    txtNumber.text = "20"
    btnOK.SetFocus
End Sub

Private Sub btn25_Click()
    txtNumber.text = "25"
    btnOK.SetFocus
End Sub

Private Sub btn50_Click()
    txtNumber.text = "50"
    btnOK.SetFocus
End Sub

Private Sub btnClearNum_Click()
    txtNumber.text = ""
End Sub

Private Sub btnClose_Click()
    unload Buyitem
End Sub

Private Sub btnItemClear_Click()
    Call formClear
End Sub
Private Sub formClear()
    txtId.text = ""
    txtJan.text = ""
    txtItemName = ""
    txtMakerCallName = ""
    txtProductNumber = ""
    txtCost.text = ""
    txtTraderId.text = ""
    cmbTraderName.text = ""
    labZaikoNum.Caption = ""
    lstItemName.Clear
    lstMakerCallName.Clear
    lstProductNumber.Clear
    Call initItemList
    txtJan.SetFocus
End Sub

Private Sub btnOk_Click()
    Dim shtMy As Worksheet
    Dim rngMy As Range
    Dim shtStock As Worksheet
    Dim Buyitem As BuyArticles
    Dim item As Articles
    Dim Trader As Traders
    Dim stockItem As StockArticles
    Dim lngPutRow As Long
    Dim strState As String
    Dim strPutERROR As String
    strPutERROR = "書き込みにエラーが発生しました。" & vbCrLf & _
                  "入庫情報は保存されません"
                  
    Set shtMy = ActiveWorkbook.Sheets("入庫")
    Set shtStock = ActiveWorkbook.Sheets("在庫")
    Set rngMy = shtMy.Columns("a")
'入力値の取得
    Buyitem.item_id = txtId.text
    Buyitem.trader_id = txtTraderId.text
    Buyitem.cost = chkItemCost(txtCost.text)
    Buyitem.number = chkNumber(txtNumber.text)
'入力値の検証
    strState = chkItemOnForm(txtId.text, txtItemName.text, _
                              txtMakerCallName.text, txtProductNumber.text)
    If Not strState Like "OK" Then _
        Call msgERROR(strState, ""): GoTo ending
    If Buyitem.cost Like "*ERROR" Then _
        Call msgERROR("エラーです", Buyitem.cost): GoTo ending
    If Buyitem.number Like "*ERROR" Then _
        Call msgERROR("エラーです", Buyitem.number): GoTo ending
'付加情報の取得
    Buyitem.id = getMaxNo(rngMy) + 1
    Buyitem.in_stock_date = Now()
'確認ダイアログ
    strState = MsgBox(txtProductNumber.text & " を " & _
                Buyitem.number & " 個入庫します。" & vbCrLf & "よろしいですか？", vbYesNo, "入庫確認")
    If Not strState Like CStr(vbYes) Then _
        Call MsgBox("出庫を中止しました"): GoTo ending
    
'入庫書き込み
    lngPutRow = getEndRow("a", shtMy) + 1
    Set rngMy = shtMy.Cells(lngPutRow, 1)
    strState = putBuyItem(rngMy, Buyitem)
    If strState Like "*ERROR" Then
        Call msgERROR(strPutERROR, "putBuyItem ERROR"): GoTo ending
    End If
    '在庫へのデータ転記
    If postBuyToStock(Buyitem, stockItem) Like "*ERROR" Then
        Call msgERROR(strPutERROR, "postBuyToStock ERROR")
        Call delRows(CStr(lngPutRow), shtMy)
        GoTo ending
    End If
'在庫書き込み
    strState = putStockItem(shtStock.Cells(getEndRow("A", shtStock) + 1, 1), stockItem)
    If strState Like "*ERROR" Then
        Call msgERROR(strPutERROR, "putStockItem ERROR")
        Call delRows(CStr(lngPutRow), shtMy)
    End If
    txtCost.text = ""
    txtNumber.text = ""
ending:
    Set shtStock = Nothing
    Set rngMy = Nothing
    Set rngMy = Nothing
End Sub
Private Sub btnItemNameSearch_Click()
    '品名検索
    Call initItemNameListChenge
End Sub

Private Sub lstItemName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '品名検索
    Call initItemNameListChenge
End Sub
Sub initItemNameListChenge()
'品名検索
    Dim strMakerList() As String
    Dim strProductNumList() As String
    Call itemNameListChenge(lstItemName.text, lstProductNumber.text, _
                            lstMakerCallName.text, strMakerList, strProductNumList)
    lstMakerCallName.list = strMakerList
    lstProductNumber.list = strProductNumList
End Sub
Private Sub lstMakerCallName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'メーカー名検索
    Call initMakerListChenge
End Sub
Private Sub btnItemMakerCallNameSerch_Click()
    'メーカー名検索
    Call initMakerListChenge
End Sub
Sub initMakerListChenge()
'メーカー名検索
    Dim strNameList() As String
    Dim strProductNumList() As String
    Dim strId As String
    
    strId = MakerListChenge(lstMakerCallName.text, lstProductNumber.text, lstItemName.text, _
                         strNameList, strProductNumList)
    lstItemName.list = strNameList
    lstProductNumber.list = strProductNumList
    If Not strId = "" Then _
        txtId.text = strId
End Sub
Private Sub lstProductNumber_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '型番検索
    Call initProductListChenge
End Sub
Private Sub btnItemProductNumberSearch_Click()
    '型番検索
    Call initProductListChenge
    txtNumber.SetFocus
End Sub
Sub initProductListChenge()
'型番検索
    Dim strMakerCallNameList() As String
    Dim strItemNameList() As String
    Dim strId As String
    Dim itemData As Articles
    Dim Trader As Traders
    Dim maker As makers
    
    strId = ProductListChenge(lstMakerCallName.text, lstProductNumber.text, lstItemName.text, _
                           strMakerCallNameList, strItemNameList)
    lstItemName.list = strItemNameList
    lstMakerCallName.list = strMakerCallNameList
    'id転記
    Call putFormData(strId)
ending:
End Sub
Private Sub putFormData(strId As String)
    Dim itemData As Articles
    Dim Trader As Traders
    Dim maker As makers
    
    If Not strId = "" Then
        txtId.text = strId
        Call getItem(strId, itemData)
        Call getTrader(itemData.trader_id, Trader)
        Call getMaker(itemData.maker_id, maker)
        txtJan.text = itemData.JAN_CODE
        txtItemName.text = itemData.name
        txtMakerCallName.text = maker.call_name
        txtProductNumber = itemData.product_number
        txtCost = itemData.cost
        txtTraderId.text = Trader.id
        cmbTraderName.text = Trader.company_name
        labZaikoNum.Caption = getZaikoNum(txtId.text)
    End If
End Sub
Private Sub sbtnNumber_SpinDown()
    txtNumber.text = SpinDownNum(txtNumber.text)
End Sub

Private Sub sbtnNumber_SpinUp()
    txtNumber.text = SpinUpNum(txtNumber.text)
End Sub

Private Sub txtJan_AfterUpdate()
'Janが変更されたとき、該当するアイテムがあれば表示する
    Debug.Print "update!"
    putFormData (getItemIdFromItemJanCode(txtJan.text))
End Sub

Private Sub txtNumber_Change()
    If IsNumeric(txtNumber.text) = False Then GoTo ending
    If IsNumeric(txtCost.text) = False Then GoTo ending
    labSum.Caption = CCur(txtNumber.text) * CCur(txtCost.text)
ending:
End Sub

Private Sub txtTraderId_Change()
    Dim shtMy As Worksheet
    Dim Trader As Traders
    Dim maker As makers
    Dim strState As String
    
    If txtTraderId.text = "" Then GoTo ending
    Set shtMy = ActiveWorkbook.Sheets("取引業者")
    strState = getKeyData(txtTraderId.text, getFindRange(shtMy, Customers_id_COL), , xlWhole)
    If strState = "" Then
        cmbTraderName.text = ""
    Else
        strState = getTrader(txtTraderId.text, Trader)
        txtTraderId.text = Trader.id
        cmbTraderName.text = Trader.company_name
    End If
    Set shtMy = Nothing
ending:
End Sub

Private Sub UserForm_Initialize()
    Call initItemList
End Sub
Private Sub initItemList()
'品名リストを初期化する
    Dim strList() As String
    
    Call initList(strList, "item_name")
    lstItemName.list = strList
    Call initList(strList, "item_maker_call_name")
    lstMakerCallName.list = strList
    Call initList(strList, "item_product_number")
    lstProductNumber.list = strList
End Sub
