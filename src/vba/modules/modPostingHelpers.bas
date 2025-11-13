Attribute VB_Name = "modPostingHelpers"
Option Explicit
'====================================================================
' MODULE : modPostingHelpers
' PURPOSE: Core helpers used across all posting modules:
'          - Table lookups
'          - ID generation and assignment
'          - Safe write-back helpers
' DEPENDS: None (standalone)
' UPDATED: 2025-11-11
'====================================================================

' Return a Dictionary representing a single table row (by key)
Public Function GetTableRow(ByVal TableName As String, ByVal KeyName As String, ByVal KeyValue As Variant) As Object
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindTableWorksheet(TableName)
    If ws Is Nothing Then Exit Function
    Set lo = ws.ListObjects(TableName)
    If lo Is Nothing Then Exit Function

    Dim i As Long, dict As Object
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(KeyName).Index).Value) = CStr(KeyValue) Then
            Set dict = CreateObject("Scripting.Dictionary")
            Dim j As Long
            For j = 1 To lo.ListColumns.Count
                dict(lo.ListColumns(j).Name) = lo.DataBodyRange.Cells(i, j).Value
            Next j
            Set GetTableRow = dict
            Exit Function
        End If
    Next i
End Function

' Return a Collection of row-dictionaries matching KeyName=KeyValue
Public Function GetTableRows(ByVal TableName As String, ByVal KeyName As String, ByVal KeyValue As Variant) As Collection
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindTableWorksheet(TableName)
    If ws Is Nothing Then Exit Function
    Set lo = ws.ListObjects(TableName)
    Dim col As New Collection
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(KeyName).Index).Value) = CStr(KeyValue) Then
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            Dim j As Long
            For j = 1 To lo.ListColumns.Count
                dict(lo.ListColumns(j).Name) = lo.DataBodyRange.Cells(i, j).Value
            Next j
            col.Add dict
        End If
    Next i
    Set GetTableRows = col
End Function

' Find worksheet that hosts a ListObject by table name
Public Function FindTableWorksheet(ByVal TableName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Dim lo As ListObject: Set lo = ws.ListObjects(TableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set FindTableWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

' Get next numeric ID from a table's ID column (safe for empty tables)
Public Function GetNextIDFromTable(ByVal lo As ListObject, ByVal keyColumnName As String) As Long
    On Error Resume Next
    If lo Is Nothing Then GetNextIDFromTable = 1: Exit Function
    Dim idx As Long
    idx = 0
    On Error Resume Next
    idx = lo.ListColumns(keyColumnName).Index
    On Error GoTo 0
    If idx = 0 Then GetNextIDFromTable = 1: Exit Function
    If lo.DataBodyRange Is Nothing Then
        GetNextIDFromTable = 1
        Exit Function
    End If
    Dim maxVal As Variant
    maxVal = Application.Max(lo.ListColumns(keyColumnName).DataBodyRange)
    If IsError(maxVal) Or IsEmpty(maxVal) Then
        GetNextIDFromTable = 1
    Else
        GetNextIDFromTable = CLng(maxVal) + 1
    End If
End Function

' AssignNextID: helper to assign the new ID into the newly added lr and return it
Public Function AssignNextID(ByVal lo As ListObject, ByVal IDColumn As String, ByRef lr As ListRow) As Long
    Dim nextID As Long
    nextID = GetNextIDFromTable(lo, IDColumn)
    lr.Range(lo.ListColumns(IDColumn).Index) = nextID
    AssignNextID = nextID
End Function

' PUBLIC: check whether a named column exists in the ListObject
Public Function ColumnExistsInTable(ByVal lo As ListObject, ByVal colName As String) As Boolean
    On Error Resume Next
    Dim test As ListColumn
    Set test = lo.ListColumns(colName)
    ColumnExistsInTable = (Not test Is Nothing)
End Function

' Generic safe write-back by key using a dictionary of columnName->value
Public Sub WriteBackRow(ByVal TableName As String, ByVal KeyName As String, ByVal KeyValue As Variant, ByVal dict As Object)
    Dim ws As Worksheet, lo As ListObject
    Set ws = FindTableWorksheet(TableName)
    If ws Is Nothing Then Exit Sub
    Set lo = ws.ListObjects(TableName)
    Dim i As Long, j As Long
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(KeyName).Index).Value) = CStr(KeyValue) Then
            For j = 1 To lo.ListColumns.Count
                Dim colName As String: colName = lo.ListColumns(j).Name
                If dict.Exists(colName) Then
                    lo.DataBodyRange.Cells(i, j).Value = dict(colName)
                End If
            Next j
            Exit Sub
        End If
    Next i
End Sub
