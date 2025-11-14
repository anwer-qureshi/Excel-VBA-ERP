 url=https://github.com/anwer-qureshi/Excel-VBA-ERP/blob/main/src/vba/modules/modInventoryPosting.bas
Attribute VB_Name = "modInventoryPosting"
Option Explicit
'====================================================================
' MODULE : modInventoryPosting
' PURPOSE: Create inventory transaction records for stock movement
' DEPENDS: modPostingHelpers (AssignNextID, ColumnExistsInTable), GetTableRows
' UPDATED: 2025-11-13 (refactor)
'====================================================================

Public Sub PostInventoryMovements(ByVal SourceType As String, ByVal SourceID As Long, ByVal transID As Long)
    On Error GoTo ErrHandler

    ' Normalize and validate SourceType
    SourceType = Trim$(UCase$(SourceType))
    If Len(SourceType) = 0 Then
        Debug.Print "PostInventoryMovements: empty SourceType"
        Exit Sub
    End If

    ' Supported types (extend as needed)
    If SourceType <> "SI" And SourceType <> "PI" Then
        Debug.Print "PostInventoryMovements: unsupported SourceType=" & SourceType
        Exit Sub
    End If

    ' Fetch source lines - should return Collection of dictionary-like objects
    Dim lines As Collection
    Set lines = GetTableRows("tbl_SalesInvoiceLines", "SalesInvoiceID", SourceID)
    If lines Is Nothing Then
        Debug.Print "PostInventoryMovements: GetTableRows returned Nothing for SourceID=" & SourceID
        Exit Sub
    End If
    If lines.Count = 0 Then
        Debug.Print "PostInventoryMovements: no lines to post for SourceID=" & SourceID
        Exit Sub
    End If

    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets("InventoryTransactions")
    Set lo = ws.ListObjects("tbl_InventoryTransactions")

    ' Cache column indexes to avoid repeated lookups
    Dim colIdx As Object
    Set colIdx = CreateObject("Scripting.Dictionary")
    Dim cols As Variant
    cols = Array("InventoryTransID", "ProductID", "QuantityOut", "QuantityIn", "SourceType", "SourceID", "TransID", "TransDate", "CreatedOn", "CreatedBy", "Rate", "Remarks")
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If ColumnExistsInTable(lo, cols(i)) Then colIdx(cols(i)) = lo.ListColumns(cols(i)).Index
    Next i

    ' Use consistent timestamp for all rows in this operation
    Dim ts As Date: ts = Now

    Application.ScreenUpdating = False

    Dim ln As Variant
    For Each ln In lines
        Dim lr As ListRow
        Set lr = lo.ListRows.Add

        ' Assign and write InventoryTransID
        Dim nextID As Long
        nextID = AssignNextID(lo, "InventoryTransID", lr)
        If colIdx.Exists("InventoryTransID") Then lr.Range(1, colIdx("InventoryTransID")).Value = nextID

        ' Product
        If colIdx.Exists("ProductID") Then
            On Error Resume Next
            lr.Range(1, colIdx("ProductID")).Value = ln("ProductID")
            On Error GoTo ErrHandler
        End If

        ' Quantity: defensive conversion using helper
        Dim qty As Double: qty = SafeDouble(ln("Quantity"))
        If SourceType = "SI" Then
            If colIdx.Exists("QuantityOut") Then lr.Range(1, colIdx("QuantityOut")).Value = qty
        ElseIf SourceType = "PI" Then
            If colIdx.Exists("QuantityIn") Then lr.Range(1, colIdx("QuantityIn")).Value = qty
        End If

        ' Rate (optional)
        If colIdx.Exists("Rate") Then lr.Range(1, colIdx("Rate")).Value = SafeCurrency(ln("Rate"))

        ' Common metadata
        If colIdx.Exists("SourceType") Then lr.Range(1, colIdx("SourceType")).Value = SourceType
        If colIdx.Exists("SourceID") Then lr.Range(1, colIdx("SourceID")).Value = SourceID
        If colIdx.Exists("TransID") Then lr.Range(1, colIdx("TransID")).Value = transID
        If colIdx.Exists("TransDate") Then lr.Range(1, colIdx("TransDate")).Value = ts
        If colIdx.Exists("CreatedOn") Then lr.Range(1, colIdx("CreatedOn")).Value = ts
        If colIdx.Exists("CreatedBy") Then lr.Range(1, colIdx("CreatedBy")).Value = GetCurrentUserName()
        If colIdx.Exists("Remarks") Then lr.Range(1, colIdx("Remarks")).Value = IIf(HasKey(ln, "Description"), ln("Description"), "")

        Debug.Print "Posted InventoryTransID=" & nextID & " SourceType=" & SourceType & " SourceID=" & SourceID & " ProductID=" & ln("ProductID") & " Qty=" & qty
    Next ln

Done:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Debug.Print "ERROR in PostInventoryMovements: " & Err.Number & " - " & Err.Description
    Resume Done
End Sub