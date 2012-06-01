VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockToDelivery 
   Caption         =   "出庫"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   OleObjectBlob   =   "StockToDelivery.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "StockToDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn1_Click()
    txtDeliveryNum.text = "1"
End Sub

Private Sub btn2_Click()
    txtDeliveryNum.text = "2"
End Sub

Private Sub btn3_Click()
    txtDeliveryNum.text = "3"
End Sub

Private Sub btn4_Click()
    txtDeliveryNum.text = "4"
End Sub

Private Sub btn5_Click()
    txtDeliveryNum.text = "5"
End Sub

Private Sub btnClose_Click()
    unload StockToDelivery
End Sub

Private Sub btnCustomerClear_Click()
    txtCustomerId.text = ""
    txtCustomerSite.text = ""
    lstCustomerSite.Clear
    txtCustomerFloor.text = ""
    lstCustomerFloor.Clear
    txtCustomerPlace.text = ""
    lstCustomerPlace.Clear
    Call initCustomer
End Sub

Private Sub btnItemClear_Click()
    Call itemClear
End Sub
Private Sub itemClear()
    txtItemId.text = ""
    txtJan.text = ""
    txtName.text = ""
    txtMakerCallName = ""
    txtItemProductNum = ""
    lstName.Clear
    lstMakerCallName.Clear
    lstItemProductNum.Clear
    Call initItem
    txtJan.SetFocus
End Sub

Private Sub btnJanReport_Click()
'入力されているJANコードから登録申請を行う
    txtItemId.text = JanRegister(txtJan.text, txtItemId.text)
    Call setStockData(txtItemId.text)
End Sub

Private Sub btnOut_Click()
'出庫を行う
    Dim dblNum As Double
    Dim strState As String
    
    '整合性チェック
    strState = chkItemOnForm(txtItemId.text, txtName.text, _
                                txtMakerCallName.text, txtItemProductNum.text)
    If Not strState Like "OK" Then _
        Call msgERROR(strState, ""): GoTo ending
        
    If txtDeliveryNum.text = "" Then _
        MsgBox ("数量の入力がありません"): GoTo ending
    If txtItemPrice.text = "" Or CDbl(txtItemPrice.text) <= 0 Then _
        MsgBox ("販売価格の設定されていない商品は出庫出来ません"): GoTo ending
    On Error Resume Next
    dblNum = CDbl(txtDeliveryNum.text)
    If Err.number <> 0 Then
        MsgBox ("入力値が数値ではありません")
        On Error GoTo 0
        GoTo ending
    End If
    If chkBillType(cmbBilltype.text) = False Then _
        MsgBox ("請求方法の指定が正しくありません"): GoTo ending
    If txtCustomerId.text = "" Then MsgBox ("出庫先の指定がありません"): GoTo ending
    If getTableDatas(Sheets("取引先"), txtCustomerId.text, Customers_id_COL, Customers_id_COL) = "" Then _
        MsgBox ("取引先idが無効です"): GoTo ending
    '確認ダイアログ
    strState = MsgBox(txtItemProductNum.text & " を [ " & txtCustomerPlace.text & " ] に  " & _
                dblNum & " 個出庫します。" & vbCrLf & "よろしいですか？", vbYesNo, "出庫確認")
    If Not strState Like CStr(vbYes) Then _
        Call MsgBox("出庫を中止しました"): GoTo ending
    strState = moveStockToDelivery(txtItemId.text, dblNum, txtCustomerId.text, cmbBilltype.text)
    If strState Like "*ERROR" Then
        Call msgERROR(strState, "")
        GoTo ending
    End If
    Debug.Print "moveStockToDelivery " & strState
    txtDeliveryNum.text = ""
    putZaikoNum (txtItemId.text)
    txtJan.SetFocus
ending:
End Sub
Private Sub btnFloor_Click()
    If lstCustomerFloor.ListIndex = -1 Then GoTo ending
    
    Call CustomerIdChenge(lstCustomerSite.text, lstCustomerFloor.text, lstCustomerPlace.text)
    Call PlaseChenge(lstCustomerSite.text, lstCustomerFloor.text)
ending:
End Sub

Private Sub lstCustomerFloor_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CustomerIdChenge(lstCustomerSite.text, lstCustomerFloor.text, lstCustomerPlace.text)
    Call PlaseChenge(lstCustomerSite.text, lstCustomerFloor.text)
End Sub
Private Sub btnPlace_Click()
    If lstCustomerPlace.ListIndex = -1 Then GoTo ending
    
    Call FloorChenge(lstCustomerSite.text, lstCustomerPlace.text)
    Call CustomerIdChenge(lstCustomerSite.text, lstCustomerFloor.text, lstCustomerPlace.text)
    txtDeliveryNum.SetFocus
ending:
End Sub
Private Sub lstCustomerPlace_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call FloorChenge(lstCustomerSite.text, lstCustomerPlace.text)
    Call CustomerIdChenge(lstCustomerSite.text, lstCustomerFloor.text, lstCustomerPlace.text)
End Sub
Private Sub btnSite_Click()
    If lstCustomerSite.ListIndex = -1 Then GoTo ending
    
    Call FloorChenge(lstCustomerSite.text, lstCustomerPlace.text)
    Call PlaseChenge(lstCustomerSite.text, lstCustomerFloor.text)
ending:
End Sub

Private Sub lstCustomerSite_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call FloorChenge(lstCustomerSite.text, lstCustomerPlace.text)
    Call PlaseChenge(lstCustomerSite.text, lstCustomerFloor.text)
End Sub
Private Sub FloorChenge(strSite As String, strPlace As String)
    Dim strFloorOnSite As String
    Dim strFloorOnPlace As String
    Dim shtMy As Worksheet
    Dim strList() As String
    Set shtMy = ActiveWorkbook.Sheets("取引先")
    'フロアリストの更新
    strFloorOnSite = getTableDatas(shtMy, strSite, Customers_site_COL, Customers_floor_COL)
    strFloorOnPlace = getTableDatas(shtMy, strPlace, Customers_place_COL, Customers_floor_COL)
    
    strList = Split(Trim(strFloorOnSite & " " & strFloorOnPlace), " ")
    Call DuplicationMerge(strList)
    lstCustomerFloor.list = strList
ending:
    Set shtMy = Nothing
End Sub
Private Sub PlaseChenge(strSite As String, strFloor As String)
'サイトとフロアの引数から場所名リストを更新する
    Dim strPlaceOnSite As String
    Dim strPlaceOnFloor As String
    Dim strIDs As String
    Dim strId() As String
    Dim i As Long
    Dim shtMy As Worksheet
    Dim strList() As String
    Dim strLists As String
    Dim varId As Variant
    
    Set shtMy = ActiveWorkbook.Sheets("取引先")
    
    strPlaceOnSite = getTableDatas(shtMy, strSite, Customers_site_COL, Customers_id_COL)
    strPlaceOnFloor = getTableDatas(shtMy, strFloor, Customers_floor_COL, Customers_id_COL)
    strIDs = Trim(strPlaceOnSite & " " & strPlaceOnFloor)
    If Not strPlaceOnSite = "" And Not strPlaceOnFloor = "" Then
        strId = Split(strIDs, " ")
        strIDs = DuplicationDraw(strId)
    End If
    strLists = getTableDatas(shtMy, strIDs, Customers_id_COL, Customers_place_COL)
    If strLists = "" Then GoTo ending
    strList = Split(strLists)
    Call DuplicationMerge(strList)
    
    lstCustomerPlace.list = strList
ending:
End Sub
Private Sub CustomerIdChenge(strSite As String, strFloor As String, strPlace As String)
'サイト、フロア、場所名からCustomer_idを更新する
    Dim strIdOnSite As String
    Dim strIdOnFloor As String
    Dim strMixedId() As String
    Dim strMixedIds As String
    Dim strIdOnPlace As String
    Dim shtMy As Worksheet
    Dim strId() As String
    Set shtMy = ActiveWorkbook.Sheets("取引先")
    
    strIdOnSite = getTableDatas(shtMy, strSite, Customers_site_COL, Customers_id_COL)
    strIdOnFloor = getTableDatas(shtMy, strFloor, Customers_floor_COL, Customers_id_COL)
    strIdOnPlace = getTableDatas(shtMy, strPlace, Customers_place_COL, Customers_id_COL)
    If UBound(Split(strIdOnPlace, " ")) = 0 Then
        txtCustomerId.text = strIdOnPlace
        GoTo ending
    End If
    strMixedId = Split(Trim(strIdOnSite & " " & strIdOnFloor), " ")
    If UBound(strMixedId) = 0 Then
        strMixedIds = strMixedId(0)
    End If
    If UBound(strMixedId) < 0 Then
        GoTo ending
    Else
        If strIdOnSite = "" Or strIdOnFloor = "" Then
            strMixedIds = Join(strMixedId, " ")
        Else
            strMixedIds = Trim(DuplicationDraw(strMixedId))
        End If
        strId = Split(Trim(strMixedIds & " " & strIdOnPlace), " ")
    End If
    If UBound(strId) > 0 Then
        Call DuplicationDraw(strId)
    End If
    If UBound(strId) < 0 Then GoTo ending
    
    If UBound(strId) = 0 And Not strId(0) = "" Then
        txtCustomerId.text = strId(0)
    End If
ending:
    Set shtMy = Nothing
End Sub
Private Sub btnProductNum_Click()
    '型番検索
    Call initProductListChenge
End Sub
Private Sub lstItemProductNum_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '型番検索
    Call initProductListChenge
End Sub
Sub initProductListChenge()
'型番検索
    Dim strId As String
    Dim strMakerCallNameList() As String
    Dim strItemNameList() As String
    Dim callName As String
    Dim productNum As String
    Dim itemName As String
    Dim callNameChenge As Boolean
    
    txtItemId.text = ""
    txtItemProductNum.text = ""
    labStockNumOfSum.Caption = ""
    If Not txtMakerCallName.text = "" Then
        callName = txtMakerCallName.text
    Else
        callName = lstMakerCallName.text
    End If
    productNum = lstItemProductNum.text
    If Not txtName.text = "" Then
        itemName = txtName.text
    Else
        itemName = lstName.text
    End If
    
    Debug.Print "call name: " & callName
    Debug.Print "product number: " & productNum
    Debug.Print "item name: " & itemName
    
    strId = ProductListChengeNew(callName, productNum, itemName, strMakerCallNameList, strItemNameList, callNameChenge)
    If callNameChenge = True Then
        lstMakerCallName.list = strMakerCallNameList
    End If
    
    If Not strId = "" Then
        txtItemId.text = strId
        Call setStockData(strId)
    End If
End Sub
Private Sub btnMaker_Click()
    'メーカー名検索
    Call initMakerListChenge
End Sub
Private Sub lstMakerCallName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'メーカー名検索
    Call initMakerListChenge
End Sub
Sub initMakerListChenge()
'メーカー名検索
    Dim strNameList() As String
    Dim strProductNumList() As String
    Dim strId As String
    
    txtName.text = ""
    txtItemProductNum.text = ""
    txtMakerCallName.text = lstMakerCallName.text
    Call MakerListChenge(lstMakerCallName.text, lstItemProductNum.text, lstName.text, _
                         strNameList, strProductNumList)
    lstName.list = strNameList
    lstItemProductNum.list = strProductNumList
    
    If Not strId = "" Then txtItemId.text = strId
End Sub
Private Sub btnItemName_Click()
    '品名検索
    Call initItemNameListChenge
End Sub
Private Sub lstName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '品名検索
    Call initItemNameListChenge
End Sub
Sub initItemNameListChenge()
'品名検索
    Dim j As Long
    Dim i As Long
    Dim strMakerList() As String
    Dim strProductNumList() As String
    Dim strName As String
    Dim strCallName As String
    
    j = lstMakerCallName.ListCount - 1
    If j > 0 Then
        ReDim strMakerList(j)
        For i = 0 To j
            strMakerList(i) = lstMakerCallName.list(i)
        Next
    End If
    j = lstItemProductNum.ListCount - 1
    ReDim strProductNumList(j)
    For i = 0 To j
        strProductNumList(i) = lstItemProductNum.list(i)
    Next

    Call itemNameListChenge(lstName.text, lstItemProductNum.text, _
                            lstMakerCallName.text, strMakerList, _
                            strProductNumList)
    txtName.text = lstName.text
    txtMakerCallName.text = ""
    txtItemProductNum.text = ""
    If UBound(strMakerList) = 0 Then txtMakerCallName.text = strMakerList(0)
    lstMakerCallName.list = strMakerList
    lstItemProductNum.list = strProductNumList
End Sub
Private Sub sbtnNumber_SpinDown()
    txtDeliveryNum.text = SpinDownNum(txtDeliveryNum.text)
End Sub

Private Sub sbtnNumber_SpinUp()
    txtDeliveryNum.text = SpinUpNum(txtDeliveryNum.text)
End Sub

Private Sub txtCustomerId_Change()
    Dim strState As String
    Dim shtMy As Worksheet
    
    If txtCustomerId.text = "" Then GoTo ending
    Set shtMy = Sheets("取引先")
    cmbBilltype.text = getTableDatas(shtMy, txtCustomerId.text, _
                                       Customers_id_COL, Customers_bill_type_COL)
    txtCustomerSite.text = getTableDatas(shtMy, txtCustomerId.text, _
                                        Customers_id_COL, Customers_site_COL)
    txtCustomerFloor.text = getTableDatas(shtMy, txtCustomerId.text, _
                                        Customers_id_COL, Customers_floor_COL)
    txtCustomerPlace.text = getTableDatas(shtMy, txtCustomerId.text, _
                                            Customers_id_COL, Customers_place_COL)
ending:
    Set shtMy = Nothing
End Sub

Private Sub txtDeliveryNum_Change()
    
    If Not chkNumber(txtDeliveryNum.text) Like "*ERROR" Then
        If txtItemPrice.text = "" Then txtItemPrice.text = "0"
        labDeliverySum.Caption = CCur(txtItemPrice.text) * CDbl(txtDeliveryNum.text)
        btnOut.Enabled = True
    Else
        btnOut.Enabled = False
        labDeliverySum.Caption = ""
    End If
End Sub

Private Sub setStockData(strItemId As String)
    Dim itemData As Articles
    Dim makerData As makers
    
    Call getItem(strItemId, itemData)
    Call getMaker(itemData.maker_id, makerData)
    Call putZaikoNum(strItemId)
    txtItemId.text = itemData.id
    txtName.text = itemData.name
    txtMakerCallName.text = makerData.call_name
    txtItemProductNum.text = itemData.product_number
    txtItemPrice.text = itemData.price
    txtJan.text = itemData.JAN_CODE
End Sub
Private Sub putZaikoNum(strItemId As String)
'在庫状態を更新する
    Dim dblNum As Double
    Dim strState As String
    
    strState = getZaikoNum(strItemId)
    On Error Resume Next
    dblNum = CDbl(strState)
    On Error GoTo 0
    If dblNum > 0 Then
        labStockNumOfSum.ForeColor = RGB(0, 0, 0)
        labStockNumOfSum.Caption = strState
        btnOut.Enabled = True
    Else
        labStockNumOfSum.ForeColor = RGB(255, 0, 0)
        labStockNumOfSum.Caption = strState
        btnOut.Enabled = False
    End If
End Sub

Private Sub clearForm()
    txtName.text = ""
    txtMakerCallName.text = ""
    txtItemProductNum.text = ""
    txtItemPrice.text = ""
    labStockNumOfSum.Caption = ""
End Sub

Private Sub txtItemId_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Dim item As Articles
    Dim stockItem() As StockArticles
    Dim maker As makers
    Dim strState As String

    If getItem(txtItemId.text, item) Like "*ERROR" Then
        Cancel = True
    Else
        Call setStockData(txtItemId.text)
        txtCustomerId.SetFocus
    End If
End Sub

Private Sub txtJan_AfterUpdate()
    Dim itemID As String
    
    itemID = getItemIdFromItemJanCode(txtJan.text)
    If itemID <> "" Then
        Call setStockData(itemID)
        txtCustomerId.SetFocus
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim shtMy As Worksheet
    Dim rngMy As Range
    Dim strList() As String
    Dim varList As Variant
    
    '取引先初期化
    Call initCustomer
    '品目初期化
    Call initItem
    '請求タイプ初期化
    varList = Range("bill_type_list").Value
    cmbBilltype.list = varList
    
    Set rngMy = Nothing
    Set shtMy = Nothing
End Sub
Private Sub initCustomer()
'取引先を初期化する
    Dim strList() As String
    '物件初期化
    Call initList(strList, "customer_site")
    lstCustomerSite.list = strList
    'フロア初期化
    Call initList(strList, "customer_floor")
    lstCustomerFloor.list = strList
    '場所名初期化
    Call initList(strList, "customer_place")
    lstCustomerPlace.list = strList
End Sub
Sub initItem()
'品目を初期化する
    Dim strList() As String
    '品目初期化
    Call initList(strList, "item_name")
    lstName.list = strList
    'メーカー名初期化
    Call initList(strList, "item_maker_call_name")
    lstMakerCallName.list = strList
    '型番
    Call initList(strList, "item_product_number")
    lstItemProductNum.list = strList
End Sub

