Attribute VB_Name = "modPostingErrorLog"
Option Explicit
'====================================================================
' MODULE : modPostingErrorLog
' PURPOSE: Log and maintain posting-related errors into tbl_PostingErrors
' DEPENDS: modPostingHelpers (AssignNextID, GetNextIDFromTable)
' UPDATED: 2025-11-11
'====================================================================

Public Sub LogPostingError(ByVal SourceType As String, ByVal SourceID As Long, ByVal ErrNo As Long, ByVal ErrMsg As String)
    On Error GoTo Fallback
    Dim ws As Worksheet, lo As ListObject, lr As ListRow
    Set ws = ThisWorkbook.Worksheets("SystemPostingErrors")
    Set lo = ws.ListObjects("tbl_PostingErrors")

    ' Ensure next ErrorID is assigned and placed
    Dim nextID As Long
    Set lr = lo.ListRows.Add
    nextID = AssignNextID(lo, "ErrorID", lr)

    If ErrNo = 0 Then ErrNo = -1
    If Len(Trim(ErrMsg)) = 0 Then ErrMsg = "No description provided by caller."

    lr.Range(lo.ListColumns("SourceType").Index) = SourceType
    lr.Range(lo.ListColumns("SourceID").Index) = SourceID
    lr.Range(lo.ListColumns("ErrNo").Index) = ErrNo
    lr.Range(lo.ListColumns("ErrMsg").Index) = ErrMsg
    On Error Resume Next
    lr.Range(lo.ListColumns("ErrProcedure").Index) = "PostTransaction"
    lr.Range(lo.ListColumns("PostedTransID").Index) = ""
    lr.Range(lo.ListColumns("IsResolved").Index) = False
    lr.Range(lo.ListColumns("Remarks").Index) = ""
    lr.Range(lo.ListColumns("CreatedBy").Index) = Environ("Username")
    lr.Range(lo.ListColumns("CreatedOn").Index) = Now
    Exit Sub

Fallback:
    ' If table doesn't exist, write a minimal log to a new worksheet
    On Error Resume Next
    Dim tmpWs As Worksheet
    Set tmpWs = ThisWorkbook.Worksheets.Add
    tmpWs.Name = "PostingErrors_Fallback_" & Format(Now, "hhmmss")
    tmpWs.Range("A1").Value = "SourceType"
    tmpWs.Range("B1").Value = "SourceID"
    tmpWs.Range("C1").Value = "ErrNo"
    tmpWs.Range("D1").Value = "ErrMsg"
    tmpWs.Range("A2").Value = SourceType
    tmpWs.Range("B2").Value = SourceID
    tmpWs.Range("C2").Value = ErrNo
    tmpWs.Range("D2").Value = ErrMsg
End Sub
