 url=https://github.com/anwer-qureshi/Excel-VBA-ERP/blob/main/src/vba/modules/modPostingTest.bas
Attribute VB_Name = "modPostingTest"
Option Explicit
'====================================================================
' MODULE : modPostingTest
' PURPOSE: Lightweight integration tests for posting flows
' UPDATED: 2025-11-13
'====================================================================

' Test: Post a simple Sales Invoice and verify GL balance and cleanup on rollback
Public Sub Test_PostTransaction_HappyPath()
    On Error GoTo ErrHandler
    Dim testSI As Long
    ' NOTE: This test assumes there is a redacted test SalesInvoice with ID 1 present in workbooks/samples
    testSI = 1
    Dim beforeEntries As Long
    beforeEntries = CountRowsInTable("tbl_GeneralLedger")

    ' Attempt posting
    PostTransaction "SI", testSI

    ' Verify posting created GL entries
    Dim afterEntries As Long
    afterEntries = CountRowsInTable("tbl_GeneralLedger")
    If afterEntries <= beforeEntries Then
        Err.Raise vbObjectError + 1001, "Test_PostTransaction_HappyPath", "Expected GL entries to increase after posting."
    Else
        Debug.Print "Test_PostTransaction_HappyPath: GL entries increased by " & (afterEntries - beforeEntries)
    End If

    Exit Sub
ErrHandler:
    Debug.Print "Test_PostTransaction_HappyPath failed: " & Err.Number & " - " & Err.Description
End Sub

' Test: Create a dummy transaction and call RollbackTransaction then verify cleanup
Public Sub Test_RollbackTransaction_Cleanup()
    On Error GoTo ErrHandler
    Dim transID As Long
    transID = CreateTransactionHeader("TI", "TESTR", "Test Rollback", 0, 0)
    CreateGLLine transID, GetSystemControlAccount("DefaultSales"), 10, True, "TestDr", "", "TEST"
    CreateGLLine transID, GetSystemControlAccount("DefaultSales"), 10, False, "TestCr", "", "TEST"

    ' Ensure created
    Dim countBefore As Long
    countBefore = CountMatchingRows("tbl_GeneralLedger", "TransID", transID)

    If countBefore = 0 Then Err.Raise vbObjectError + 1002, "Test_RollbackTransaction_Cleanup", "No GL entries created for test trans."

    ' Rollback
    RollbackTransaction transID

    Dim countAfter As Long
    countAfter = CountMatchingRows("tbl_GeneralLedger", "TransID", transID)
    If countAfter <> 0 Then
        Err.Raise vbObjectError + 1003, "Test_RollbackTransaction_Cleanup", "Rollback did not remove GL entries."
    Else
        Debug.Print "Test_RollbackTransaction_Cleanup: rollback cleaned up " & countBefore & " entries."
    End If

    Exit Sub
ErrHandler:
    Debug.Print "Test_RollbackTransaction_Cleanup failed: " & Err.Number & " - " & Err.Description
End Sub

' Simple helpers used by tests
Public Function CountRowsInTable(ByVal tblName As String) As Long
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindTableWorksheet(tblName)
    If ws Is Nothing Then CountRowsInTable = 0: Exit Function
    Set lo = ws.ListObjects(tblName)
    If lo.DataBodyRange Is Nothing Then CountRowsInTable = 0 Else CountRowsInTable = lo.ListRows.Count
End Function

Public Function CountMatchingRows(ByVal tblName As String, ByVal keyName As String, ByVal keyValue As Variant) As Long
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindTableWorksheet(tblName)
    If ws Is Nothing Then CountMatchingRows = 0: Exit Function
    Set lo = ws.ListObjects(tblName)
    Dim i As Long, cnt As Long: cnt = 0
    If lo.DataBodyRange Is Nothing Then CountMatchingRows = 0: Exit Function
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(keyName).Index).Value) = CStr(keyValue) Then cnt = cnt + 1
    Next i
    CountMatchingRows = cnt
End Function