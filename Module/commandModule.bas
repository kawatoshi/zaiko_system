Attribute VB_Name = "commandModule"
Option Explicit
Sub ���C��()
    Call SheetUnvisible
    MainForm.Show
End Sub
Sub ����()
    BuyForm.Show
    unload BuyForm
End Sub
Function JanRegister(strJan As String, strItemId As String) As String
    Dim strState As String
    
    '��������
    JAN_CODE = strJan
    item_id = strItemId
    'JAN�`�F�b�N
    strState = chkJan(JAN_CODE)
    Select Case strState
    Case ""
        Debug.Print "jan code ok"
    Case Else
        MsgBox (strState)
        ClearJanRegister
        Exit Function
    End Select
    'Item�`�F�b�N
    strState = chkItemHasJanCode(item_id, JAN_CODE)
    Select Case strState
    Case "ERROR"
        item_id = ""
    Case "nomatch"
        Debug.Print "jan code no match"
        item_id = ""
    Case ""
        Debug.Print "item id is ok"
    Case Else
        MsgBox (strState)
        ClearJanRegister
        Exit Function
    End Select
    '�t�H�[������
    JanRegist.Show
    JanRegister = JanRegist.txtItemId.text
    '�㏈��
    JanRegister = item_id
    JAN_CODE = ""
    item_id = ""
    unload JanRegist
End Function
Private Sub ClearJanRegister()
    JAN_CODE = ""
    item_id = ""
End Sub
Sub �o��()
    DeliveryForm.Show
    unload DeliveryForm
End Sub
Sub ����()
    SettleForm.Show
    unload SettleForm
End Sub
Sub �Վ��i�ڒǉ�()
    ExItem.Show
End Sub
Sub �o�[�W����()
    MsgBox (strVer)
End Sub
Private Sub ���X()
    If Not ActiveSheet.name = "�݌Ƀ��X�g" Then _
        MsgBox ("�݌Ƀ��X�g�V�[�g�ōs���Ă�������"): Exit Sub
    StockToLost.Show
    unload StockToLost
End Sub

Private Sub ����ԕi()
    Dim lngRow As Long
    
    If Not ActiveSheet.name = "�o�Ƀ��X�g" Then _
        MsgBox ("�o�Ƀ��X�g�V�[�g�ōs���Ă�������"): Exit Sub
    lngRow = Selection.Row
    If lngRow < DATA_START_ROW Then MsgBox ("�����ȍs�ł�"): Exit Sub
    If Cells(lngRow, DeliveryList_id_COL) = "" Then _
        MsgBox ("�f�[�^���Ȃ��s��I�����Ă��܂�"): Exit Sub
    returnDeliveryItem.Show
    unload StockToLost
End Sub

