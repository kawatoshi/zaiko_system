VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockToDelivery 
   Caption         =   "�o��"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   OleObjectBlob   =   "StockToDelivery.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
'���͂���Ă���JAN�R�[�h����o�^�\�����s��
    txtItemId.text = JanRegister(txtJan.text, txtItemId.text)
    Call setStockData(txtItemId.text)
End Sub

Private Sub btnOut_Click()
'�o�ɂ��s��
    Dim dblNum As Double
    Dim strState As String
    
    '�������`�F�b�N
    strState = chkItemOnForm(txtItemId.text, txtName.text, _
                                txtMakerCallName.text, txtItemProductNum.text)
    If Not strState Like "OK" Then _
        Call msgERROR(strState, ""): GoTo ending
        
    If txtDeliveryNum.text = "" Then _
        MsgBox ("���ʂ̓��͂�����܂���"): GoTo ending
    If txtItemPrice.text = "" Or CDbl(txtItemPrice.text) <= 0 Then _
        MsgBox ("�̔����i�̐ݒ肳��Ă��Ȃ����i�͏o�ɏo���܂���"): GoTo ending
    On Error Resume Next
    dblNum = CDbl(txtDeliveryNum.text)
    If Err.number <> 0 Then
        MsgBox ("���͒l�����l�ł͂���܂���")
        On Error GoTo 0
        GoTo ending
    End If
    If chkBillType(cmbBilltype.text) = False Then _
        MsgBox ("�������@�̎w�肪����������܂���"): GoTo ending
    If txtCustomerId.text = "" Then MsgBox ("�o�ɐ�̎w�肪����܂���"): GoTo ending
    If getTableDatas(Sheets("�����"), txtCustomerId.text, Customers_id_COL, Customers_id_COL) = "" Then _
        MsgBox ("�����id�������ł�"): GoTo ending
    '�m�F�_�C�A���O
    strState = MsgBox(txtItemProductNum.text & " �� [ " & txtCustomerPlace.text & " ] ��  " & _
                dblNum & " �o�ɂ��܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�o�Ɋm�F")
    If Not strState Like CStr(vbYes) Then _
        Call MsgBox("�o�ɂ𒆎~���܂���"): GoTo ending
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
    Set shtMy = ActiveWorkbook.Sheets("�����")
    '�t���A���X�g�̍X�V
    strFloorOnSite = getTableDatas(shtMy, strSite, Customers_site_COL, Customers_floor_COL)
    strFloorOnPlace = getTableDatas(shtMy, strPlace, Customers_place_COL, Customers_floor_COL)
    
    strList = Split(Trim(strFloorOnSite & " " & strFloorOnPlace), " ")
    Call DuplicationMerge(strList)
    lstCustomerFloor.list = strList
ending:
    Set shtMy = Nothing
End Sub
Private Sub PlaseChenge(strSite As String, strFloor As String)
'�T�C�g�ƃt���A�̈�������ꏊ�����X�g���X�V����
    Dim strPlaceOnSite As String
    Dim strPlaceOnFloor As String
    Dim strIDs As String
    Dim strId() As String
    Dim i As Long
    Dim shtMy As Worksheet
    Dim strList() As String
    Dim strLists As String
    Dim varId As Variant
    
    Set shtMy = ActiveWorkbook.Sheets("�����")
    
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
'�T�C�g�A�t���A�A�ꏊ������Customer_id���X�V����
    Dim strIdOnSite As String
    Dim strIdOnFloor As String
    Dim strMixedId() As String
    Dim strMixedIds As String
    Dim strIdOnPlace As String
    Dim shtMy As Worksheet
    Dim strId() As String
    Set shtMy = ActiveWorkbook.Sheets("�����")
    
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
    '�^�Ԍ���
    Call initProductListChenge
End Sub
Private Sub lstItemProductNum_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '�^�Ԍ���
    Call initProductListChenge
End Sub
Sub initProductListChenge()
'�^�Ԍ���
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
    '���[�J�[������
    Call initMakerListChenge
End Sub
Private Sub lstMakerCallName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '���[�J�[������
    Call initMakerListChenge
End Sub
Sub initMakerListChenge()
'���[�J�[������
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
    '�i������
    Call initItemNameListChenge
End Sub
Private Sub lstName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '�i������
    Call initItemNameListChenge
End Sub
Sub initItemNameListChenge()
'�i������
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
    Set shtMy = Sheets("�����")
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
'�݌ɏ�Ԃ��X�V����
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
    
    '����揉����
    Call initCustomer
    '�i�ڏ�����
    Call initItem
    '�����^�C�v������
    varList = Range("bill_type_list").Value
    cmbBilltype.list = varList
    
    Set rngMy = Nothing
    Set shtMy = Nothing
End Sub
Private Sub initCustomer()
'����������������
    Dim strList() As String
    '����������
    Call initList(strList, "customer_site")
    lstCustomerSite.list = strList
    '�t���A������
    Call initList(strList, "customer_floor")
    lstCustomerFloor.list = strList
    '�ꏊ��������
    Call initList(strList, "customer_place")
    lstCustomerPlace.list = strList
End Sub
Sub initItem()
'�i�ڂ�����������
    Dim strList() As String
    '�i�ڏ�����
    Call initList(strList, "item_name")
    lstName.list = strList
    '���[�J�[��������
    Call initList(strList, "item_maker_call_name")
    lstMakerCallName.list = strList
    '�^��
    Call initList(strList, "item_product_number")
    lstItemProductNum.list = strList
End Sub

