Attribute VB_Name = "findModule"
Function findItem(data As Articles, _
                  sorted_item() As Articles, _
                  item_id As String) As Boolean
    Dim min As Long
    Dim max As Long
    Dim middle As Long
    
    findItem = False
    min = LBound(sorted_item)
    max = UBound(sorted_item)
    Do While min < max
        middle = Int((min + max) / 2)
        If CStr(sorted_item(middle).id) < item_id Then
            min = middle + 1
        Else
            max = middle
        End If
    Loop
    If CStr(sorted_item(min).id) = item_id Then
        data = sorted_item(min)
        findItem = True
    End If
End Function
Function findCustomer(data As Customers, customer() As Customers, customer_id As String) As Boolean
'customer_idからカスタマーを抽出する
    Dim i As Long
    
    findCustomer = False
    For i = 0 To UBound(customer)
        If customer(i).id Like customer_id Then
            data = customer(i)
            findCustomer = True
        End If
    Next
End Function
Function findCustomerForDeliveryAccount(customer() As Customers, customer_id As String) As String
'カスタマーidから丸広請求内訳用売場名を返す
    Dim i As Long
    
    For i = 0 To UBound(customer)
        If customer(i).id Like customer_id Then
            findCustomerForDeliveryAccount = customer(i).floor & " " & customer(i).place
            Exit Function
        End If
    Next
    findCustomerForDeliveryAccount = "no data"
End Function
Function findSettleItemsByBillType(data() As SettleArticles, _
                                   settle_item() As SettleArticles, _
                                   strBilltype As String) As Boolean
'決済済データから指定BillTypeのデータを抽出する
    Dim i As Long, j As Long
    Dim varType As Variant
    Dim varBillType As Variant
    
    findSettleItemsByBillType = False
    If UBound(settle_item) <= 0 Then Exit Function
    ReDim Preserve data(UBound(settle_item))
    i = 0
    varType = Split(strBilltype)
    For j = 0 To UBound(settle_item)
        For Each varBillType In varType
            If settle_item(j).bill_type Like varBillType Then
                data(i) = settle_item(j)
              i = i + 1
              Exit For
            End If
        Next
    Next
    If i = 0 Then Exit Function
    ReDim Preserve data(i - 1)
    findSettleItemsByBillType = True
End Function
