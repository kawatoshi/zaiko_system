Attribute VB_Name = "commandModule"
Option Explicit
Sub メイン()
    Call SheetUnvisible
    MainForm.Show
End Sub
Sub 入庫()
    BuyForm.Show
    unload BuyForm
End Sub
Function JanRegister(strJan As String, strItemId As String) As String
    Dim strState As String
    
    '初期処理
    JAN_CODE = strJan
    item_id = strItemId
    'JANチェック
    strState = chkJan(JAN_CODE)
    Select Case strState
    Case ""
        Debug.Print "jan code ok"
    Case Else
        MsgBox (strState)
        ClearJanRegister
        Exit Function
    End Select
    'Itemチェック
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
    'フォーム処理
    JanRegist.Show
    JanRegister = JanRegist.txtItemId.text
    '後処理
    JanRegister = item_id
    JAN_CODE = ""
    item_id = ""
    unload JanRegist
End Function
Private Sub ClearJanRegister()
    JAN_CODE = ""
    item_id = ""
End Sub
Sub 出庫()
    DeliveryForm.Show
    unload DeliveryForm
End Sub
Sub 決済()
    SettleForm.Show
    unload SettleForm
End Sub
Sub 臨時品目追加()
    ExItem.Show
End Sub
Sub バージョン()
    MsgBox (strVer)
End Sub
Private Sub ロス()
    If Not ActiveSheet.name = "在庫リスト" Then _
        MsgBox ("在庫リストシートで行ってください"): Exit Sub
    StockToLost.Show
    unload StockToLost
End Sub

Private Sub 売場返品()
    Dim lngRow As Long
    
    If Not ActiveSheet.name = "出庫リスト" Then _
        MsgBox ("出庫リストシートで行ってください"): Exit Sub
    lngRow = Selection.Row
    If lngRow < DATA_START_ROW Then MsgBox ("無効な行です"): Exit Sub
    If Cells(lngRow, DeliveryList_id_COL) = "" Then _
        MsgBox ("データがない行を選択しています"): Exit Sub
    returnDeliveryItem.Show
    unload StockToLost
End Sub

