VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JanRegist 
   Caption         =   "JANコード登録申請"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   OleObjectBlob   =   "JanRegist.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "JanRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Submit_Click()
    If Not txtItemId.text = "" Then
        item_id = txtItemId.text
        Call sendJanRegistMail("kawakita@fujibil.co.jp", txtJan.text, txtItemId.text)
'        Call sendJanRegistMail("iwata@fujibil.co.jp", txtJan.text, txtItemId.text)
        Me.Hide
    Else
        MsgBox ("商品が確定してません")
    End If
End Sub

Private Sub UserForm_Initialize()
    textInit
    '品目初期化
    Call initItem
    txtJan.Enabled = False
    txtItemId.Enabled = False
    Submit.Enabled = False
End Sub
Private Function setItemData(strItemId As String, strJanCode As String) As Boolean
    Dim itemData As Articles
    Dim makerData As makers
    setItemData = False
    
    Select Case getItemIdFromItemJanCode(strJanCode)
    Case ""
    Case Else
        setItemData = False
        Exit Function
    End Select
    
    Call getItem(strItemId, itemData)
    Call getMaker(itemData.maker_id, makerData)
    txtItemId.text = itemData.id
    txtName.text = itemData.name
    txtMakerCallName.text = makerData.call_name
    txtItemProductNum.text = itemData.product_number
End Function
Sub initItem()
'品目を初期化する
    Dim strIdsByMaker As String
    Dim strList() As String
    Dim strIdsByProductNum As String
    
    strIdsByMaker = getMakerCallNameFromJanCode(txtJan.text, strList)
    Select Case strIdsByMaker
    Case ""
        '品目初期化
        Call initList(strList, "item_name")
        lstName.list = strList
        'メーカー名初期化
        Call initList(strList, "item_maker_call_name")
        lstMakerCallName.list = strList
        '型番
        Call initList(strList, "item_product_number")
        lstItemProductNum.list = strList
    Case Else
        lstMakerCallName.list = strList
        If UBound(strList) = 0 Then txtMakerCallName.text = strList(0)
        strIdsByProductNum = getProductIDsFromKeys(strList, strIdsByMaker, Articles_maker_id_COL)
        lstItemProductNum.list = strList
        Call getItemNameList(strList, strIdsByProductNum, Articles_id_COL)
        lstName.list = strList
    End Select
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
    ReDim strMakerList(j)
    For i = 0 To j
        strMakerList(i) = lstMakerCallName.list(i)
    Next
    j = lstItemProductNum.ListCount - 1
    ReDim strProductNumList(j)
    For i = 0 To j
        strProductNumList(i) = lstItemProductNum.list(i)
    Next

    Call itemNameListChenge(lstName.text, lstItemProductNum.text, _
                            lstMakerCallName.text, strMakerList, strProductNumList)
    txtName.text = lstName.text
    If UBound(strMakerList) = 0 Then txtMakerCallName.text = strMakerList(0)
    lstMakerCallName.list = strMakerList
    lstItemProductNum.list = strProductNumList
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
        Call setItemData(strId, txtJan.text)
        Submit.Enabled = True
    End If
End Sub
Private Sub btnItemClear_Click()
    Call itemClear
End Sub
Private Sub itemClear()
    textInit
    txtName.text = ""
    txtMakerCallName = ""
    txtItemProductNum = ""
    lstName.Clear
    lstMakerCallName.Clear
    lstItemProductNum.Clear
    Call initItem
End Sub
Private Sub textInit()
    txtJan.text = JAN_CODE
    txtItemId.text = item_id
End Sub
