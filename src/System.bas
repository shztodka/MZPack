Attribute VB_Name = "System"
'***********************************************
'           MZPack - Excel Utility Pack
'        2012, Katarzyna Zbroinska
'             katarzyna@zbroinski.net
'              www.zbroinski.net
'
' This software is distributed under the term of
'     General Public License (GNU) version 3.
'***********************************************

Option Explicit

Public Sub SaveWithBackup(ByVal control As IRibbonControl)

'Macro saves active workbook with its backup copy
'Author: Katarzyna Zbroiñska,
'2011-11-01, Warszawa.

Dim strFullFileName As String
Dim intAnswer As Integer

Const SUFFIX As String = " (Backup copy)."

    On Error Resume Next
   
    'Save current file
    ActiveWorkbook.Save
    If Err <> 0 Then
        intAnswer = MsgBox("I couldn't save the active workbook.", vbOKOnly, "Error!")
        Exit Sub
    End If
    
    'Save backup copy
    With ActiveWorkbook
        strFullFileName = StrReverse(Replace(StrReverse(.FullName), ".", StrReverse(SUFFIX), , 1, vbTextCompare))
        .SaveCopyAs (strFullFileName)
    End With
    If Err <> 0 Then
        intAnswer = MsgBox("I coulnd't save the copy of the active workbook.", vbOKOnly, "Error!")
        Exit Sub
    End If
    
End Sub
