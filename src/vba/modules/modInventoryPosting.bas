Attribute VB_Name = "modInventoryPosting"
Option Explicit
'====================================================================
' MODULE : modInventoryPosting
' PURPOSE: Create inventory transaction records for stock movement
' DEPENDS: modPostingHelpers (AssignNextID, ColumnExistsInTable)
' UPDATED: 2025-11-11
'====================================================================

Public Sub PostInventoryMovements(ByVal SourceType As String, ByVal SourceID As Long, ByVal transID As Long)
    ' Only handle Sales Invoice inventory outflow in current implementation
    If SourceType <> "SI" Then Exit Sub

    Dim lines As Collection
    Set lines = GetTableRows("tbl_SalesInvoiceLines", "SalesInvoiceID", SourceID)

    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets("InventoryTransactions")
    Set lo = ws.ListObjects("tbl_InventoryTransactions")

    Dim ln As Variant
    For Each ln In lines
        ' Add new inventory transaction row and assign InventoryTransID
        Dim lr As ListRow
        Set lr = lo.ListRows.Add
        Dim nextID As Long
        nextID = AssignNextID(lo, "InventoryTransID", lr)

        ' Populate fields (defensive column checks)
        If ColumnExistsInTable(lo, "ProductID") Then lr.Range(lo.ListColumns("ProductID").Index) = ln("ProductID")
        If ColumnExistsInTable(lo, "QuantityOut") Then lr.Range(lo.ListColumns("QuantityOut").Index) = CDbl(ln("Quantity"))
        If ColumnExistsInTable(lo, "SourceType") Then lr.Range(lo.ListColumns("SourceType").Index) = SourceType
        If ColumnExistsInTable(lo, "SourceID") Then lr.Range(lo.ListColumns("SourceID").Index) = SourceID
        If ColumnExistsInTable(lo, "TransID") Then lr.Range(lo.ListColumns("TransID").Index) = transID
        If ColumnExistsInTable(lo, "TransDate") Then lr.Range(lo.ListColumns("TransDate").Index) = Now
        If ColumnExistsInTable(lo, "CreatedOn") Then lr.Range(lo.ListColumns("CreatedOn").Index) = Now
    Next ln
End Sub
