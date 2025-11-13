Attribute VB_Name = "modPostingTest"
Option Explicit
'====================================================================
' MODULE : modPostingTest
' PURPOSE: Seed minimal test master data and run PostTransaction("SI",id)
' DEPENDS: modPostingHelpers, modSystemAccounts, modPostingEngine
' UPDATED: 2025-11-11
'====================================================================

Public Sub Test_SalesInvoice_Posting()
    On Error GoTo ErrHandler
    Debug.Print "=== Starting Sales Invoice Posting Test ==="

    Dim custID As Long, prodID As Long, invID As Long

    ' 1. Seed system accounts (SysAcctID assigned inside)
    LoadSystemAccounts

    ' 2. Create Customer
    custID = CreateTestCustomer()
    Debug.Print "Created Customer ID: "; custID

    ' 3. Create Product
    prodID = CreateTestProduct()
    Debug.Print "Created Product ID: "; prodID

    ' 4. Create Sales Invoice
    invID = CreateTestSalesInvoice(custID, prodID)
    Debug.Print "Created Sales Invoice ID: "; invID

    ' 5. Execute Posting
    Debug.Print "Executing PostTransaction('SI'," & invID & ") ..."
    PostTransaction "SI", invID
    Debug.Print "Posting completed."

    ' 6. Verify
    VerifyPosting invID

    Debug.Print "=== Posting Test Completed ==="
    Exit Sub

ErrHandler:
    Debug.Print "Error during test: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    LogPostingError "SI", invID, Err.Number, Err.Description
End Sub


'--- LoadSystemAccounts (uses AssignNextID for SysAcctID) ---
Private Sub LoadSystemAccounts()
    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets("SystemAccounts")
    Set lo = ws.ListObjects("tbl_SystemAccounts")

    If lo.ListRows.Count > 0 Then lo.DataBodyRange.Delete

    Dim accounts As Variant
    accounts = Array( _
        Array("DefaultAR", "Trade Accounts Receivable"), _
        Array("DefaultAP", "Trade Accounts Payable"), _
        Array("DefaultSales", "Product Sales"), _
        Array("DefaultCOGS", "Cost of Goods Sold"), _
        Array("DefaultInventory", "Inventory - Fertilizers"), _
        Array("DefaultTaxPayable", "Sales Tax (GST Payable)"), _
        Array("DiscountAllowed", "Sale Discounts Allowed") _
    )

    Dim i As Long, lr As ListRow, nextID As Long, acctCode As String
    For i = LBound(accounts) To UBound(accounts)
        acctCode = FindAccountCodeByName(CStr(accounts(i)(1)))
        If acctCode <> "" Then
            Set lr = lo.ListRows.Add
            nextID = AssignNextID(lo, "SysAcctID", lr)
            lr.Range(lo.ListColumns("KeyName").Index) = accounts(i)(0)
            lr.Range(lo.ListColumns("AccountCode").Index) = acctCode
            lr.Range(lo.ListColumns("AccountName").Index) = accounts(i)(1)
            lr.Range(lo.ListColumns("Description").Index) = "Auto-seeded for test"
            lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
            lr.Range(lo.ListColumns("CreatedOn").Index) = Now
        End If
    Next i
End Sub

Private Function FindAccountCodeByName(ByVal namePart As String) As String
    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets("ChartOfAccounts")
    Set lo = ws.ListObjects("tbl_ChartOfAccounts")
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If InStr(1, CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("AccountName").Index).Value), namePart, vbTextCompare) > 0 Then
            FindAccountCodeByName = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("AccountCode").Index).Value)
            Exit Function
        End If
    Next i
End Function

Private Function CreateTestCustomer() As Long
    Dim ws As Worksheet, lo As ListObject, lr As ListRow, nextID As Long
    Set ws = ThisWorkbook.Worksheets("Customers")
    Set lo = ws.ListObjects("tbl_Customers")
    Set lr = lo.ListRows.Add
    nextID = AssignNextID(lo, "CustomerID", lr)
    lr.Range(lo.ListColumns("CustomerName").Index) = "TEST CUSTOMER"
    lr.Range(lo.ListColumns("AccountCode").Index) = GetSystemControlAccount("DefaultAR")
    lr.Range(lo.ListColumns("Address").Index) = "N/A"
    lr.Range(lo.ListColumns("Phone").Index) = "000-0000"
    lr.Range(lo.ListColumns("Email").Index) = "test@example.com"
    lr.Range(lo.ListColumns("Status").Index) = "Active"
    lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
    lr.Range(lo.ListColumns("CreatedOn").Index) = Now
    CreateTestCustomer = nextID
End Function

Private Function CreateTestProduct() As Long
    Dim ws As Worksheet, lo As ListObject, lr As ListRow, nextID As Long
    Set ws = ThisWorkbook.Worksheets("Products")
    Set lo = ws.ListObjects("tbl_Products")
    Set lr = lo.ListRows.Add
    nextID = AssignNextID(lo, "ProductID", lr)
    lr.Range(lo.ListColumns("ProductName").Index) = "TEST PRODUCT"
    lr.Range(lo.ListColumns("Category").Index) = "General"
    lr.Range(lo.ListColumns("SubCategory").Index) = "Test"
    lr.Range(lo.ListColumns("Unit").Index) = "PCS"
    lr.Range(lo.ListColumns("InventoryAccount").Index) = GetSystemControlAccount("DefaultInventory")
    lr.Range(lo.ListColumns("COGSAccount").Index) = GetSystemControlAccount("DefaultCOGS")
    lr.Range(lo.ListColumns("SalesAccount").Index) = GetSystemControlAccount("DefaultSales")
    lr.Range(lo.ListColumns("PurchaseAccount").Index) = GetSystemControlAccount("DefaultAP")
    lr.Range(lo.ListColumns("ReorderLevel").Index) = 10
    lr.Range(lo.ListColumns("Status").Index) = "Active"
    lr.Range(lo.ListColumns("Notes").Index) = "Seed product for test"
    lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
    lr.Range(lo.ListColumns("CreatedOn").Index) = Now
    CreateTestProduct = nextID
End Function

Private Function CreateTestSalesInvoice(ByVal custID As Long, ByVal prodID As Long) As Long
    Dim ws As Worksheet, lo As ListObject, lr As ListRow, nextID As Long
    Set ws = ThisWorkbook.Worksheets("SalesInvoices")
    Set lo = ws.ListObjects("tbl_SalesInvoices")
    Set lr = lo.ListRows.Add
    nextID = AssignNextID(lo, "SalesInvoiceID", lr)
    lr.Range(lo.ListColumns("InvoiceNo").Index) = "TEST-" & Format(Now, "yymmdd-hhnnss")
    lr.Range(lo.ListColumns("InvoiceDate").Index) = Date
    lr.Range(lo.ListColumns("CustomerID").Index) = custID
    lr.Range(lo.ListColumns("SubTotal").Index) = 1000
    lr.Range(lo.ListColumns("DiscountAmount").Index) = 0
    lr.Range(lo.ListColumns("TaxAmount").Index) = 0
    lr.Range(lo.ListColumns("TotalAmount").Index) = 1000
    lr.Range(lo.ListColumns("Status").Index) = "Pending"
    lr.Range(lo.ListColumns("IsPosted").Index) = False
    lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
    lr.Range(lo.ListColumns("CreatedOn").Index) = Now

    ' Insert Sales Invoice Line
    Dim loL As ListObject, lrL As ListRow, lineID As Long
    Set loL = ThisWorkbook.Worksheets("SalesInvoiceLines").ListObjects("tbl_SalesInvoiceLines")
    Set lrL = loL.ListRows.Add
    lineID = AssignNextID(loL, "InvoiceLineID", lrL)
    lrL.Range(loL.ListColumns("SalesInvoiceID").Index) = nextID
    lrL.Range(loL.ListColumns("ProductID").Index) = prodID
    lrL.Range(loL.ListColumns("Description").Index) = "TEST PRODUCT SALE"
    lrL.Range(loL.ListColumns("Quantity").Index) = 10
    lrL.Range(loL.ListColumns("Unit").Index) = "PCS"
    lrL.Range(loL.ListColumns("Rate").Index) = 100
    lrL.Range(loL.ListColumns("LineAmount").Index) = 1000
    lrL.Range(loL.ListColumns("NetAmount").Index) = 1000
    lrL.Range(loL.ListColumns("Status").Index) = "Pending"
    lrL.Range(loL.ListColumns("CreatedOn").Index) = Now

    CreateTestSalesInvoice = nextID
End Function

Private Sub VerifyPosting(ByVal SalesInvoiceID As Long)
    Debug.Print "----- Verification Results -----"
    Debug.Print "Transactions Count: "; CountTableRows("tbl_Transactions")
    Debug.Print "TransactionLines Count: "; CountTableRows("tbl_TransactionLines")
    Debug.Print "GeneralLedger Count: "; CountTableRows("tbl_GeneralLedger")
    Debug.Print "InventoryTransactions Count: "; CountTableRows("tbl_InventoryTransactions")
    Debug.Print "PostingErrors Count: "; CountTableRows("tbl_PostingErrors")
End Sub

Private Function CountTableRows(ByVal TableName As String) As Long
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindTableWorksheet(TableName)
    If ws Is Nothing Then CountTableRows = 0: Exit Function
    Set lo = ws.ListObjects(TableName)
    If lo Is Nothing Then CountTableRows = 0: Exit Function
    CountTableRows = lo.ListRows.Count
End Function
