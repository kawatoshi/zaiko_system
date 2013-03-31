Attribute VB_Name = "MainModule"
Option Explicit
Public Const strVer As String = "消耗品管理システム" & vbCrLf & "ver1.800"
Public Const DataBaseName As String = "z_system_data*"

Public JAN_CODE As String
Public item_id As String

'全体定数
Public Const DATA_START_ROW As Long = 6     'データシートのスタートrow

'支払いシート
Public Const Payments_id_COL As Long = 1
Public Const Payments_trader_id_COL As Long = 2
Public Const Payments_date_COL As Long = 3
Public Const Payments_sum_COL As Long = 4
Public Const Payments_tax_COL As Long = 5

'品目シート
Public Const Articles_id_COL As Long = 1
Public Const Articles_category_COL As Long = 2
Public Const Articles_name_COL As Long = 3
Public Const Articles_product_number_COL As Long = 4
Public Const Articles_maker_id_COL As Long = 5
Public Const Articles_fujibil_code_COL As Long = 6
Public Const Articles_JAN_code_COL As Long = 7
Public Const Articles_cost_COL As Long = 8
Public Const Articles_price_without_tax_COL As Long = 9
Public Const Articles_tax_COL As Long = 10
Public Const Articles_price_COL As Long = 11
Public Const Articles_trader_id_COL As Long = 12
Public Const Articles_entry_date_COL As Long = 13

'取引業者シート
Public Const Traders_id_COL As Long = 1
Public Const Traders_company_name_COL As Long = 2
Public Const Traders_tel_COL As Long = 3
Public Const Traders_address_COL As Long = 4
Public Const Traders_person_name_COL As Long = 5

'メーカーシート
Public Const Makers_id_COL As Long = 1
Public Const Makers_name_COL As Long = 2
Public Const Makers_call_name_COL As Long = 3

'取引先シート
Public Const Customers_id_COL As Long = 1
Public Const Customers_site_COL As Long = 2
Public Const Customers_floor_COL As Long = 3
Public Const Customers_place_COL As Long = 4
Public Const Customers_claim_name_COL As Long = 5
Public Const Customers_tenant_code_COL As Long = 6
Public Const Customers_A_table_COL As Long = 7
Public Const Customers_bill_type_COL As Long = 8

'入庫シート
Public Const BuyArticles_id_COL As Long = 1
Public Const BuyArticles_payment_id_COL As Long = 2
Public Const BuyArticles_item_id_COL As Long = 3
Public Const BuyArticles_trader_id_COL As Long = 4
Public Const BuyArticles_cost_COL As Long = 5
Public Const BuyArticles_number_COL As Long = 6
Public Const BuyArticles_in_stock_date_COL As Long = 7

'在庫シート
Public Const StockArticles_id_COL As Long = 1
Public Const StockArticles_buy_article_id_COL As Long = 2
Public Const StockArticles_item_id_COL As Long = 3
Public Const StockArticles_cost_COL As Long = 4
Public Const StockArticles_number_COL As Long = 5
Public Const StockArticles_final_delivery_date_COL As Long = 6
Public Const StockArticles_receipt_article_id_COL As Long = 7

'出庫シート
Public Const DeliveryArticles_id_COL As Long = 1
Public Const DeliveryArticles_buy_article_id_COL As Long = 2
Public Const DeliveryArticles_stock_article_id_COL As Long = 3
Public Const DeliveryArticles_item_id_COL As Long = 4
Public Const DeliveryArticles_customer_id_COL As Long = 5
Public Const DeliveryArticles_cost_COL As Long = 6
Public Const DeliveryArticles_item_price_without_tax_COL As Long = 7
Public Const DeliveryArticles_item_price_COL As Long = 8
Public Const DeliveryArticles_number_COL As Long = 9
Public Const DeliveryArticles_sum_COL As Long = 10
Public Const DeliveryArticles_bill_type_COL As Long = 11
Public Const DeliveryArticles_delivery_date_COL As Long = 12


'ロスシート
Public Const LostArticles_id_COL As Long = 1
Public Const LostArticles_buy_article_id_COL As Long = 2
Public Const LostArticles_stock_article_id_COL As Long = 3
Public Const LostArticles_item_id_COL As Long = 4
Public Const LostArticles_cost_COL As Long = 5
Public Const LostArticles_number_COL As Long = 6
Public Const LostArticles_lost_date_COL As Long = 7
Public Const LostArticles_note_COL As Long = 8

'返品履歴シート
Public Const ReturnArticles_id_COL As Long = 1
Public Const ReturnArticles_buy_article_id_COL As Long = 2
Public Const ReturnArticles_stock_article_id_COL As Long = 3
Public Const ReturnArticles_item_id_COL As Long = 4
Public Const ReturnArticles_customer_id_COL As Long = 5
Public Const ReturnArticles_cost_COL As Long = 6
Public Const ReturnArticles_item_price_COL As Long = 7
Public Const ReturnArticles_number_COL As Long = 8
Public Const ReturnArticles_return_date_COL As Long = 9

'出庫リストシート
Public Const DeliveryList_id_COL As Long = 1
Public Const DeliveryList_type_name_COL As Long = 2
Public Const DeliveryList_item_name_COL As Long = 3
Public Const DeliveryList_item_price_COL As Long = 4
Public Const DeliveryList_number_COL As Long = 5
Public Const DeliveryList_sum_COL As Long = 6
Public Const DeliveryList_customer_name_COL As Long = 7
Public Const DeliveryList_delivery_date_COL As Long = 8

'決済済シート
Public Const SettleArticles_id_COL As Long = 1
Public Const SettleArticles_buy_article_id_COL As Long = 2
Public Const SettleArticles_stock_article_id_COL As Long = 3
Public Const SettleArticles_item_id_COL As Long = 4
Public Const SettleArticles_customer_id_COL As Long = 5
Public Const SettleArticles_cost_COL As Long = 6
Public Const SettleArticles_item_price_without_tax_COL As Long = 7
Public Const SettleArticles_item_price_COL As Long = 8
Public Const SettleArticles_number_COL As Long = 9
Public Const SettleArticles_sum_COL As Long = 10
Public Const SettleArticles_bill_type_COL As Long = 11
Public Const SettleArticles_tenant_code_COL As Long = 12
Public Const SettleArticles_delivery_date_COL As Long = 13
Public Const SettleArticles_settle_date_COL As Long = 14
Public Const SettleArticles_bill_date_COL As Long = 15

'Tmp決済シート
Public Const TmpSettleArticles_id_COL As Long = 1
Public Const TmpSettleArticles_delivery_date_COL As Long = 2
Public Const TmpSettleArticles_customer_COL As Long = 3
Public Const TmpSettleArticles_claim_name_COL As Long = 4
Public Const TmpSettleArticles_maker_COL As Long = 5
Public Const TmpSettleArticles_item_name_COL As Long = 6
Public Const TmpSettleArticles_item_COL As Long = 7
Public Const TmpSettleArticles_item_price_COL As Long = 8
Public Const TmpSettleArticles_number_COL As Long = 9
Public Const TmpSettleArticles_sum_COL As Long = 10
Public Const TmpSettleArticles_bill_type_COL As Long = 11

'出庫明細シート
Public Const DeliveryAccount_delivery_date_COL As Long = 1
Public Const DeliveryAccount_customer_COL As Long = 2
Public Const DeliveryAccount_maker_COL As Long = 3
Public Const DeliveryAccount_item_name_COL As Long = 4
Public Const DeliveryAccount_produnt_number_COL As Long = 5
Public Const DeliveryAccount_price_COL As Long = 6
Public Const DeliveryAccount_number_COL As Long = 7
Public Const DeliveryAccount_sum_COL As Long = 8

'在庫リストシート
Public Const StockList_id_COL As Long = 1
Public Const StockList_type_name_COL As Long = 2
Public Const StockList_item_name_COL As Long = 3
Public Const StockList_cost_COL As Long = 4
Public Const StockList_number_COL As Long = 5
Public Const StockList_sum_COL As Long = 6
Public Const StockList_item_price_COL As Long = 7
Public Const StockList_stock_date_COL As Long = 8
Public Const StockList_delivery_date_COL As Long = 9
'入庫リストシート
Public Const BuyList_id_COL As Long = 1
Public Const BuyList_item_id_COL As Long = 2
Public Const BuyList_name_COL As Long = 3
Public Const BuyList_product_number_COL As Long = 4
Public Const BuyList_cost_COL As Long = 5
Public Const BuyList_number_COL As Long = 6
Public Const BuyList_stock_date_COL As Long = 7

'不二ビル明細シート
Public Const DistributerAccount_id_COL As Long = 1
Public Const DistributerAccount_delivery_date_COL As Long = 2
Public Const DistributerAccount_bill_date_COL As Long = 3
Public Const DistributerAccount_item_name_COL As Long = 4
Public Const DistributerAccount_maker_name_COL As Long = 5
Public Const DistributerAccount_product_number_COL As Long = 6
Public Const DistributerAccount_customer_name_COL As Long = 7
Public Const DistributerAccount_floor_COL As Long = 8
Public Const DistributerAccount_claim_name_COL As Long = 9
Public Const DistributerAccount_bill_type_COL As Long = 10
Public Const DistributerAccount_item_price_COL As Long = 11
Public Const DistributerAccount_number_COL As Long = 12
Public Const DistributerAccount_sum_of_price_COL As Long = 13

'テナント請求内訳シート
Public Const TenantAccaunts_delivery_date_COL As Long = 1
Public Const TenantAccaunts_tenant_code_COL As Long = 2
Public Const TenantAccaunts_floor_COL As Long = 3
Public Const TenantAccaunts_place_COL As Long = 4
Public Const TenantAccaunts_maker_COL As Long = 5
Public Const TenantAccaunts_item_name_COL As Long = 6
Public Const TenantAccaunts_product_name_COL As Long = 7
Public Const TenantAccaunts_price_COL As Long = 8
Public Const TenantAccaunts_number_COL As Long = 9
Public Const TenantAccaunts_sum_COL As Long = 10


Public Type Payments
'支払い
    id As String
    trader_id As String
    date As Date
    sum As String
    tax As String
End Type

Public Type Articles
'品目
    id As String
    category As String
    name As String
    product_number As String
    maker_id As String
    fujibil_code As String
    JAN_CODE As String
    cost As String
    price_without_tax As String
    tax As String
    price As String
    trader_id As String
    entry_date As Date
End Type

Public Type Traders
'取引業者
    id As String
    company_name As String
    tel As String
    address As String
    person_name As String
End Type

Public Type makers
'メーカー
    id As String
    name As String
    call_name As String
End Type

Public Type BuyArticles
'入庫
    id As String
    item_id As String
    payment_id As String
    trader_id As String
    cost As String
    number As String
    in_stock_date As Date
End Type
Public Type StockArticles
'在庫
    id As String
    buy_article_id As String
    item_id As String
    cost As String
    number As String
    final_delivery_date As Date
    receipt_article_id As String
End Type
Public Type DeliveryArticles
'出庫
    id As String
    buy_article_id As String
    stock_article_id As String
    item_id As String
    customer_id As String
    cost As String
    item_price_without_tax As String
    item_price As String
    number As String
    sum As String
    bill_type As String
    delivery_date As Date
End Type

Public Type Customers
'取引先
    id As String
    site As String
    floor As String
    place As String
    claim_name As String
    tenant_code As String
    A_table As String
    bill_type As String
End Type
Public Type LostArticles
'ロス
    id As String
    buy_article_id As String
    stock_article_id As String
    item_id As String
    cost As String
    number As String
    lost_date As Date
    note As String
End Type

Public Type ReturnArticles
'返品履歴
    id As String
    buy_article_id As String
    stock_article_id As String
    item_id As String
    customer_id As String
    cost As String
    item_price As String
    number As String
    return_date As Date
End Type

Public Type DeliveryList
'出庫リスト
    id As String
    type_name As String
    item_name As String
    item_price As String
    number As String
    sum As String
    customer_name As String
    delivery_date As Date
End Type

Public Type SettleArticles
'決済済
    id As String
    buy_article_id As String
    stock_article_id As String
    item_id As String
    customer_id As String
    cost As String
    item_price_without_tax As String
    item_price As String
    number As String
    sum As String
    bill_type As String
    tenant_code As String
    delivery_date As Date
    settle_date As Date
    bill_date As String
End Type

Public Type TmpSettleArticles
'Tmp決済
    id As String
    delivery_date As Date
    customer As String
    claim_name As String
    maker As String
    item_name As String
    item As String
    item_price As String
    number As String
    sum As String
    bill_type As String
End Type

Public Type DeliveryAccount
'丸広請求内訳
    delivery_date As Date
    customer As String
    maker As String
    item_name As String
    produnt_number As String
    price As Long
    cost As Long
    number As Long
    sum As Long
End Type

Public Type StockList
'在庫リスト
    id As String
    type_name As String
    item_name As String
    cost As String
    number As String
    sum As String
    item_price As String
    stock_date As Date
    delivery_date As Date
End Type
Public Type BuyList
'入庫リスト
    id As String
    item_id As String
    name As String
    product_number As String
    cost As String
    number As String
    stock_date As String
End Type

Public Type DistributerAccount
'不二ビル明細
    id As String
    delivery_date As Date
    bill_date As String
    item_name As String
    maker_name As String
    product_number As String
    customer_name As String
    floor As String
    claim_name As String
    bill_type As String
    item_price As String
    number As String
    sum_of_price As String
End Type

Public Type TenantAccaunts
'テナント請求内訳
    delivery_date As Date
    tenant_code As String
    floor As String
    place As String
    claim_name As String
    maker As String
    item_name As String
    product_name As String
    price As Long
    cost As Long
    number As Long
    sum As Long
    bill_type As String
End Type

Public Type SumList
'合計一覧用データ
    claim_name As String
    floor As String
    place As String
    tenant_code As String
    price_without_tax As Long
    tax As Long
    price As Long
    cost As Long
    profit As String
    BillType As String
End Type

