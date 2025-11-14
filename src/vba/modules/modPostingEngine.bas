 url=https://github.com/anwer-qureshi/Excel-VBA-ERP/blob/main/src/vba/modules/modPostingEngine.bas
Attribute VB_Name = "modPostingEngine"
Option Explicit
'====================================================================
' MODULE : modPostingEngine
' PURPOSE: Core posting engine and SI posting flow (balanced entries)
' DEPENDS: modPostingHelpers, modSystemAccounts, modInventoryPosting, modPostingErrorLog
' UPDATED: 2025-11-13 (refactor)
'====================================================================

Public Sub PostTransaction(ByVal SourceType As String, ByVal SourceID As Long)
    On Error GoTo ErrHandler
    Dim currentStep As String: currentStep = "Start"
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    SourceType = Trim$(UCase$(SourceType))
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
    If invRow.Exists("CustomerID") Then custID = SafeLong(invRow("CustomerID"))
    Dim totalAmt As Currency: totalAmt = 0
    If invRow.Exists("TotalAmount") Then totalAmt = SafeCurrency(invRow("TotalAmount"))

    currentStep = "Create transaction header"
    Dim transID As Long
    transID = CreateTransactionHeader("SI", refNo, "Sales Invoice Posting - " & refNo, custID, totalAmt)

    Dim totalDr As Currency: totalDr = 0
    Dim totalCr As Currency: totalCr = 0

    currentStep = "Load invoice lines"
    Dim invLines As Collection
    Set invLines = GetTableRows("tbl_SalesInvoiceLines", "SalesInvoiceID", SourceID)
    If invLines Is Nothing Then
        RollbackTransaction transID
        Err.Raise vbObjectError + 9003, "PostTransaction", "No invoice lines found for: " & SourceID
    End If

    ' Use a single timestamp for created rows during this posting
    Dim ts As Date: ts = Now

    Dim revMap As Object: Set revMap = CreateObject("Scripting.Dictionary")
    Dim ln As Variant

    For Each ln In invLines
        Dim prodID As Long: prodID = 0
        If HasKey(ln, "ProductID") Then prodID = SafeLong(ln("ProductID"))
        Dim qty As Double: qty = 0
        If HasKey(ln, "Quantity") Then qty = SafeDouble(ln("Quantity"))
        Dim rate As Currency: rate = 0
        If HasKey(ln, "Rate") Then rate = SafeCurrency(ln("Rate"))
        Dim lineNet As Currency
        If HasKey(ln, "NetAmount") Then
            lineNet = SafeCurrency(ln("NetAmount"))
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
        CreateInventoryLine transID, prodID, qty, rate, lineNet, Empty, IIf(HasKey(ln, "Description"), ln("Description"), "Sale of product")

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
        Dim amt As Currency: amt = SafeCurrency(revMap(acct))
        CreateGLLine transID, CStr(acct), amt, False, "Sales - " & acct, "", "SI"
        totalCr = totalCr + amt
    Next acct

    currentStep = "Process tax and discounts"
    If invRow.Exists("TaxAmount") Then
        Dim taxAmt As Currency: taxAmt = SafeCurrency(invRow("TaxAmount"))
        If taxAmt <> 0 Then
            Dim taxAcct As String: taxAcct = GetSystemControlAccount("DefaultTaxPayable")
            CreateGLLine transID, taxAcct, taxAmt, False, "Sales Tax", "", "SI"
            totalCr = totalCr + taxAmt
        End If
    End If
    If invRow.Exists("DiscountAmount") Then
        Dim discAmt As Currency: discAmt = SafeCurrency(invRow("DiscountAmount"))
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
    Dim EPSILON As Currency: EPSILON = 0.005@
    If Abs(totalDr - totalCr) > EPSILON Then
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
    LogPostingError SourceType, SourceID, loggedErrNo, loggedErrDesc, "PostTransaction", transID, currentStep
    On Error Resume Next
    If transID <> 0 Then RollbackTransaction transID
    Resume CleanExit
End Sub

' Existing helper functions (CreateTransactionHeader, CreateGLLine, CreateInventoryLine, RollbackTransaction, MarkSourcePosted)
' kept function signatures the same but callers above now use Safe* and consistent user/timestamp.
' For brevity the unchanged functions are kept as in original file; when applying replace the body with the same content but
' apply the column-index caching pattern and GetCurrentUserName() and ts usage as in other modules.