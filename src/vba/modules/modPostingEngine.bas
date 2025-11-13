Attribute VB_Name = "modPostingEngine"
Option Explicit
'====================================================================
' MODULE : modPostingEngine
' PURPOSE: Core posting engine and SI posting flow (balanced entries)
' DEPENDS: modPostingHelpers, modSystemAccounts, modInventoryPosting, modPostingErrorLog
' UPDATED: 2025-11-11
'====================================================================

Public Sub PostTransaction(ByVal SourceType As String, ByVal SourceID As Long)
    On Error GoTo ErrHandler
    Dim currentStep As String: currentStep = "Start"
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If SourceType <> "SI" Then
        Err.Raise vbObjectError + 9000, "PostTransaction", "Only SourceType = 'SI' supported in this implementation."
    End If
    currentStep = "Validate invoice existence"

    Dim invRow As Object
    Set invRow = GetTableRow("tbl_SalesInvoices", "SalesInvoiceID", SourceID)
    If invRow Is Nothing Then Err.Raise vbObjectError + 9001, "PostTransaction", "SalesInvoice not found: " & SourceID
    If invRow.Exists("IsPosted") Then
        If CBool(invRow("IsPosted")) Then Err.Raise vbObjectError + 9002, "PostTransaction", "Invoice already posted: " & SourceID
    End If

    currentStep = "Gather header values"
    Dim refNo As String: refNo = ""
    If invRow.Exists("InvoiceNo") Then refNo = CStr(invRow("InvoiceNo"))
    Dim custID As Long: custID = 0
    If invRow.Exists("CustomerID") Then custID = CLng(invRow("CustomerID"))
    Dim totalAmt As Currency: totalAmt = 0
    If invRow.Exists("TotalAmount") Then totalAmt = CCur(invRow("TotalAmount"))

    currentStep = "Create transaction header"
    Dim transID As Long
    transID = CreateTransactionHeader("SI", refNo, "Sales Invoice Posting - " & refNo, custID, totalAmt)

    Dim totalDr As Currency: totalDr = 0
    Dim totalCr As Currency: totalCr = 0

    currentStep = "Load invoice lines"
    Dim invLines As Collection
    Set invLines = GetTableRows("tbl_SalesInvoiceLines", "SalesInvoiceID", SourceID)

    Dim revMap As Object: Set revMap = CreateObject("Scripting.Dictionary")
    Dim ln As Variant

    For Each ln In invLines
        Dim prodID As Long: prodID = 0
        If ln.Exists("ProductID") Then prodID = CLng(ln("ProductID"))
        Dim qty As Double: qty = 0
        If ln.Exists("Quantity") Then qty = CDbl(ln("Quantity"))
        Dim rate As Currency: rate = 0
        If ln.Exists("Rate") Then rate = CCur(ln("Rate"))
        Dim lineNet As Currency
        If ln.Exists("NetAmount") Then
            lineNet = CCur(ln("NetAmount"))
        Else
            lineNet = qty * rate
        End If

        currentStep = "Resolve sales account for product " & prodID
        Dim salesAcct As String
        salesAcct = GetProductAccount(prodID, "SalesAccount")
        If Len(Trim(salesAcct)) = 0 Then salesAcct = GetSystemControlAccount("DefaultSales")

        If revMap.Exists(salesAcct) Then
            revMap(salesAcct) = revMap(salesAcct) + lineNet
        Else
            revMap.Add salesAcct, lineNet
        End If

        currentStep = "Create inventory line for product " & prodID
        CreateInventoryLine transID, prodID, qty, rate, lineNet, Empty, "Sale of product"

        currentStep = "Attempt COGS for product " & prodID
        Dim prodCost As Currency: prodCost = GetProductCost(prodID)
        Dim cogsAcct As String: cogsAcct = GetProductAccount(prodID, "COGSAccount")
        If Len(Trim(cogsAcct)) > 0 And prodCost > 0 Then
            Dim cogsAmt As Currency: cogsAmt = prodCost * qty
            CreateGLLine transID, cogsAcct, cogsAmt, True, "COGS - ProductID " & prodID, "", "SI"
            totalDr = totalDr + cogsAmt
        End If
    Next ln

    currentStep = "Create revenue GL lines"
    Dim acct As Variant
    For Each acct In revMap.Keys
        Dim amt As Currency: amt = CCur(revMap(acct))
        CreateGLLine transID, CStr(acct), amt, False, "Sales - " & acct, "", "SI"
        totalCr = totalCr + amt
    Next acct

    currentStep = "Process tax and discounts"
    If invRow.Exists("TaxAmount") Then
        Dim taxAmt As Currency: taxAmt = CCur(invRow("TaxAmount"))
        If taxAmt <> 0 Then
            Dim taxAcct As String: taxAcct = GetSystemControlAccount("DefaultTaxPayable")
            CreateGLLine transID, taxAcct, taxAmt, False, "Sales Tax", "", "SI"
            totalCr = totalCr + taxAmt
        End If
    End If
    If invRow.Exists("DiscountAmount") Then
        Dim discAmt As Currency: discAmt = CCur(invRow("DiscountAmount"))
        If discAmt <> 0 Then
            Dim discAcct As String: discAcct = GetSystemControlAccount("DiscountAllowed")
            CreateGLLine transID, discAcct, discAmt, True, "Sales Discount", "", "SI"
            totalDr = totalDr + discAmt
        End If
    End If

    currentStep = "Create AR line"
    Dim custAcctCode As String
    custAcctCode = GetCustomerAccount(custID)
    Dim arAmt As Currency: arAmt = totalCr - totalDr
    If arAmt < 0 Then arAmt = 0
    If arAmt <> 0 Then
        CreateGLLine transID, custAcctCode, arAmt, True, "AR - Invoice " & refNo, "", "SI"
        totalDr = totalDr + arAmt
    End If

    currentStep = "Balance check"
    If Abs(totalDr - totalCr) > 0.005 Then
        Dim roundingAcct As String
        roundingAcct = GetSystemControlAccount("RoundingDiff")
        If Len(Trim(roundingAcct)) > 0 Then
            Dim balAmt As Currency: balAmt = Abs(totalDr - totalCr)
            If totalDr > totalCr Then
                CreateGLLine transID, roundingAcct, balAmt, False, "Auto-balance rounding", "", "SI"
                totalCr = totalCr + balAmt
            Else
                CreateGLLine transID, roundingAcct, balAmt, True, "Auto-balance rounding", "", "SI"
                totalDr = totalDr + balAmt
            End If
        Else
            RollbackTransaction transID
            Err.Raise vbObjectError + 9010, "PostTransaction", "Unbalanced posting (Dr=" & totalDr & " Cr=" & totalCr & ") and no rounding account."
        End If
    End If

    currentStep = "Mark invoice posted"
    MarkSourcePosted "SI", SourceID, transID

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Dim loggedErrNo As Long: loggedErrNo = Err.Number
    Dim loggedErrDesc As String: loggedErrDesc = Err.Description
    If loggedErrNo = 0 Then loggedErrDesc = "Unknown runtime error at step: " & currentStep
    ' Log and rollback
    LogPostingError SourceType, SourceID, loggedErrNo, loggedErrDesc & " [Step: " & currentStep & "]"
    On Error Resume Next
    If transID <> 0 Then RollbackTransaction transID
    Resume CleanExit
End Sub


' Create transaction header -> tbl_Transactions
Public Function CreateTransactionHeader(ByVal TransType As String, ByVal RefNo As String, ByVal Description As String, ByVal CustomerID As Long, ByVal TotalAmount As Currency) As Long
    Dim ws As Worksheet, lo As ListObject, lr As ListRow
    Set ws = ThisWorkbook.Worksheets("Transactions")
    Set lo = ws.ListObjects("tbl_Transactions")
    Set lr = lo.ListRows.Add
    Dim nextID As Long
    nextID = AssignNextID(lo, "TransID", lr)
    If ColumnExistsInTable(lo, "TransType") Then lr.Range(lo.ListColumns("TransType").Index) = TransType
    If ColumnExistsInTable(lo, "RefNo") Then lr.Range(lo.ListColumns("RefNo").Index) = RefNo
    If ColumnExistsInTable(lo, "Description") Then lr.Range(lo.ListColumns("Description").Index) = Description
    If ColumnExistsInTable(lo, "CustomerID") Then lr.Range(lo.ListColumns("CustomerID").Index) = CustomerID
    If ColumnExistsInTable(lo, "TotalAmount") Then lr.Range(lo.ListColumns("TotalAmount").Index) = TotalAmount
    If ColumnExistsInTable(lo, "Status") Then lr.Range(lo.ListColumns("Status").Index) = "Open"
    If ColumnExistsInTable(lo, "CreatedBy") Then lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
    If ColumnExistsInTable(lo, "CreatedOn") Then lr.Range(lo.ListColumns("CreatedOn").Index) = Now
    CreateTransactionHeader = nextID
End Function


' Create GL line -> tbl_GeneralLedger (Debit if IsDebit=True)
Public Function CreateGLLine(ByVal TransID As Long, ByVal AccountCode As String, ByVal Amount As Currency, ByVal IsDebit As Boolean, ByVal Description As String, Optional ByVal CounterAccount As String = "", Optional ByVal Source As String = "") As Long
    Dim ws As Worksheet, lo As ListObject, lr As ListRow
    Set ws = ThisWorkbook.Worksheets("GeneralLedger")
    Set lo = ws.ListObjects("tbl_GeneralLedger")
    Set lr = lo.ListRows.Add
    Dim nextID As Long
    nextID = AssignNextID(lo, "EntryID", lr)
    If ColumnExistsInTable(lo, "TransID") Then lr.Range(lo.ListColumns("TransID").Index) = TransID
    If ColumnExistsInTable(lo, "Date") Then lr.Range(lo.ListColumns("Date").Index) = Now
    If ColumnExistsInTable(lo, "AccountCode") Then lr.Range(lo.ListColumns("AccountCode").Index) = AccountCode
    If ColumnExistsInTable(lo, "Description") Then lr.Range(lo.ListColumns("Description").Index) = Description
    If ColumnExistsInTable(lo, "Debit") Then lr.Range(lo.ListColumns("Debit").Index) = IIf(IsDebit, Amount, 0)
    If ColumnExistsInTable(lo, "Credit") Then lr.Range(lo.ListColumns("Credit").Index) = IIf(Not IsDebit, Amount, 0)
    If ColumnExistsInTable(lo, "CounterAccount") Then lr.Range(lo.ListColumns("CounterAccount").Index) = CounterAccount
    If ColumnExistsInTable(lo, "Source") Then lr.Range(lo.ListColumns("Source").Index) = Source
    If ColumnExistsInTable(lo, "PostedBy") Then lr.Range(lo.ListColumns("PostedBy").Index) = Environ("Username")
    If ColumnExistsInTable(lo, "Timestamp") Then lr.Range(lo.ListColumns("Timestamp").Index) = Now
    CreateGLLine = nextID
End Function

' Create inventory/transaction line (wrapper calls modInventoryPosting or writes into TransactionLines)
Public Sub CreateInventoryLine(ByVal TransID As Long, ByVal ProductID As Long, ByVal QtyOut As Double, ByVal Rate As Currency, ByVal Amount As Currency, Optional ByVal WHID As Variant = Empty, Optional ByVal Remarks As String = "")
    ' Write to TransactionLines (inventory-style table) and let InventoryTransactions be separate
    Dim ws As Worksheet, lo As ListObject, lr As ListRow
    Set ws = ThisWorkbook.Worksheets("TransactionLines")
    Set lo = ws.ListObjects("tbl_TransactionLines")
    Set lr = lo.ListRows.Add
    Dim nextID As Long
    nextID = AssignNextID(lo, "TransLineID", lr)
    If ColumnExistsInTable(lo, "TransID") Then lr.Range(lo.ListColumns("TransID").Index) = TransID
    If ColumnExistsInTable(lo, "ProductID") Then lr.Range(lo.ListColumns("ProductID").Index) = ProductID
    If ColumnExistsInTable(lo, "QtyOut") Then lr.Range(lo.ListColumns("QtyOut").Index) = QtyOut
    If ColumnExistsInTable(lo, "Rate") Then lr.Range(lo.ListColumns("Rate").Index) = Rate
    If ColumnExistsInTable(lo, "Amount") Then lr.Range(lo.ListColumns("Amount").Index) = Amount
    If Not IsMissing(WHID) And ColumnExistsInTable(lo, "WHID") Then lr.Range(lo.ListColumns("WHID").Index) = WHID
    If ColumnExistsInTable(lo, "Remarks") Then lr.Range(lo.ListColumns("Remarks").Index) = Remarks
    If ColumnExistsInTable(lo, "CreatedBy") Then lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
    If ColumnExistsInTable(lo, "CreatedOn") Then lr.Range(lo.ListColumns("CreatedOn").Index) = Now

    ' Optionally also create a simplified InventoryTransactions record
    On Error Resume Next
    PostInventoryMovements "SI", 0, TransID  ' call is harmless if implemented to use real lines
    On Error GoTo 0
End Sub


' Rollback: remove GL entries, TransactionLines and Transactions for a given TransID
Public Sub RollbackTransaction(ByVal TransID As Long)
    On Error Resume Next
    Dim lo As ListObject, i As Long

    Set lo = ThisWorkbook.Worksheets("GeneralLedger").ListObjects("tbl_GeneralLedger")
    For i = lo.ListRows.Count To 1 Step -1
        If CLng(lo.DataBodyRange.Cells(i, lo.ListColumns("TransID").Index).Value) = TransID Then lo.ListRows(i).Delete
    Next i

    Set lo = ThisWorkbook.Worksheets("TransactionLines").ListObjects("tbl_TransactionLines")
    For i = lo.ListRows.Count To 1 Step -1
        If CLng(lo.DataBodyRange.Cells(i, lo.ListColumns("TransID").Index).Value) = TransID Then lo.ListRows(i).Delete
    Next i

    Set lo = ThisWorkbook.Worksheets("Transactions").ListObjects("tbl_Transactions")
    For i = lo.ListRows.Count To 1 Step -1
        If CLng(lo.DataBodyRange.Cells(i, lo.ListColumns("TransID").Index).Value) = TransID Then lo.ListRows(i).Delete
    Next i
End Sub

' Mark source SalesInvoice as posted (set TransactionID, PostedBy, PostedOn, IsPosted)
Public Sub MarkSourcePosted(ByVal SourceType As String, ByVal SourceID As Long, ByVal TransID As Long)
    If SourceType = "SI" Then
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        dict.Add "IsPosted", True
        dict.Add "TransactionID", TransID
        dict.Add "PostedOn", Now
        dict.Add "PostedBy", Environ("Username")
        WriteBackRow "tbl_SalesInvoices", "SalesInvoiceID", SourceID, dict
    End If
End Sub
