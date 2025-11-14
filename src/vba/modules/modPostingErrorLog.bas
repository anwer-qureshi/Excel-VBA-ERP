 url=https://github.com/anwer-qureshi/Excel-VBA-ERP/blob/main/src/vba/modules/modPostingErrorLog.bas
Attribute VB_Name = "modPostingErrorLog"
Option Explicit
'====================================================================
' MODULE : modPostingErrorLog
' PURPOSE: Log and maintain posting-related errors into tbl_PostingErrors
' DEPENDS: modPostingHelpers (AssignNextID, GetNextIDFromTable, GetCurrentUserName)
' UPDATED: 2025-11-13
'====================================================================

' Enhanced logging: includes optional procedure and transID context. Writes safely.
Public Sub LogPostingError(ByVal SourceType As String, ByVal SourceID As Long, ByVal ErrNo As Long, ByVal ErrMsg As String, Optional ByVal ProcName As String = "", Optional ByVal PostedTransID As Long = 0, Optional ByVal StepInfo As String = "")
    On Error GoTo Fallback
    Dim ws As Worksheet, lo As ListObject, lr As ListRow
    Set ws = ThisWorkbook.Worksheets("SystemPostingErrors")
    Set lo = ws.ListObjects("tbl_PostingErrors")

    ' Ensure next ErrorID is assigned and placed
    Set lr = lo.ListRows.Add
    Dim nextID As Long
    nextID = AssignNextID(lo, "ErrorID", lr)

    If ErrNo = 0 Then ErrNo = -1
    If Len(Trim(ErrMsg)) = 0 Then ErrMsg = "No description provided by caller."

    lr.Range(lo.ListColumns("SourceType").Index) = SourceType
    lr.Range(lo.ListColumns("SourceID").Index) = SourceID
    lr.Range(lo.ListColumns("ErrNo").Index) = ErrNo
    lr.Range(lo.ListColumns("ErrMsg").Index) = ErrMsg
    On Error Resume Next
    If ColumnExistsInTable(lo, "ErrProcedure") Then lr.Range(lo.ListColumns("ErrProcedure").Index) = ProcName
    If ColumnExistsInTable(lo, "PostedTransID") Then lr.Range(lo.ListColumns("PostedTransID").Index) = IIf(PostedTransID = 0, "", PostedTransID)
    If ColumnExistsInTable(lo, "Remarks") Then
        lr.Range(lo.ListColumns("Remarks").Index) = StepInfo
    End If
    If ColumnExistsInTable(lo, "CreatedBy") Then lr.Range(lo.ListColumns("CreatedBy").Index) = GetCurrentUserName()
    If ColumnExistsInTable(lo, "CreatedOn") Then lr.Range(lo.ListColumns("CreatedOn").Index) = Now
    Exit Sub

Fallback:
    ' If table doesn't exist, write minimal log to immediate window and a fallback sheet
    On Error Resume Next
    Debug.Print "LogPostingError fallback - SourceType:" & SourceType & " SourceID:" & SourceID & " ErrNo:" & ErrNo & " Msg:" & ErrMsg
    Dim tmpWs As Worksheet
    Set tmpWs = ThisWorkbook.Worksheets.Add
    tmpWs.Name = "PostingErrors_Fallback_" & Format(Now, "hhmmss")
    tmpWs.Range("A1").Value = "SourceType"
    tmpWs.Range("B1").Value = "SourceID"
    tmpWs.Range("C1").Value = "ErrNo"
    tmpWs.Range("D1").Value = "ErrMsg"
    tmpWs.Range("E1").Value = "Procedure"
    tmpWs.Range("F1").Value = "StepInfo"
    tmpWs.Range("A2").Value = SourceType
    tmpWs.Range("B2").Value = SourceID
    tmpWs.Range("C2").Value = ErrNo
    tmpWs.Range("D2").Value = ErrMsg
    tmpWs.Range("E2").Value = ProcName
    tmpWs.Range("F2").Value = StepInfo
End Sub