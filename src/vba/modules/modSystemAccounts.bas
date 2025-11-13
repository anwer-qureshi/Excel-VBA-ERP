Attribute VB_Name = "modSystemAccounts"
Option Explicit
'====================================================================
' MODULE : modSystemAccounts
' PURPOSE: Master data lookups - system and product/customer account resolution
' DEPENDS: modPostingHelpers (GetTableRow, GetTableRows)
' UPDATED: 2025-11-11
'====================================================================

' Return the customer's GL account code (AccountCode), fallback to DefaultAR
Public Function GetCustomerAccount(ByVal CustomerID As Long) As String
    Dim row As Object: Set row = GetTableRow("tbl_Customers", "CustomerID", CustomerID)
    If Not row Is Nothing Then
        If row.Exists("AccountCode") Then
            If Len(Trim(CStr(row("AccountCode")))) > 0 Then
                GetCustomerAccount = CStr(row("AccountCode"))
                Exit Function
            End If
        End If
    End If
    GetCustomerAccount = GetSystemControlAccount("DefaultAR")
End Function

' Return product account (field name is passed like "SalesAccount", "COGSAccount")
Public Function GetProductAccount(ByVal ProductID As Long, ByVal Which As String) As String
    Dim row As Object: Set row = GetTableRow("tbl_Products", "ProductID", ProductID)
    If Not row Is Nothing Then
        If row.Exists(Which) Then
            GetProductAccount = CStr(row(Which))
            Exit Function
        End If
    End If
    GetProductAccount = ""
End Function

' Return the standard cost for a product (0 if not present)
Public Function GetProductCost(ByVal ProductID As Long) As Currency
    Dim row As Object: Set row = GetTableRow("tbl_Products", "ProductID", ProductID)
    If Not row Is Nothing Then
        If row.Exists("StdCost") Then
            If IsNumeric(row("StdCost")) Then GetProductCost = CCur(row("StdCost")): Exit Function
        End If
    End If
    GetProductCost = 0
End Function

' Get control account from SystemAccounts by KeyName
Public Function GetSystemControlAccount(ByVal Key As String) As String
    Dim row As Object: Set row = GetTableRow("tbl_SystemAccounts", "KeyName", Key)
    If Not row Is Nothing Then
        If row.Exists("AccountCode") Then
            GetSystemControlAccount = CStr(row("AccountCode"))
            Exit Function
        End If
    End If
    GetSystemControlAccount = ""
End Function
