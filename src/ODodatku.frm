VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ODodatku 
   Caption         =   "About MZPack"
   ClientHeight    =   3288
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5928
   OleObjectBlob   =   "ODodatku.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ODodatku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'           MZPack - Excel Utility Pack
'        2012, Katarzyna Zbroinska
'             katarzyna@zbroinski.net
'              www.zbroinski.net
'
' This software is distributed under the term of
'     General Public License (GNU) version 3.
'***********************************************

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute& Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long)
#Else
    Private Declare Function ShellExecute& Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long)
#End If

Private Sub BZamknij_Click()

    Unload ODodatku

End Sub

Private Sub eMailLink_Click()
     
    Dim sWebAddress As String
    sWebAddress = "mailto:" & Me.eMailLink.Caption
     
    ShellExecute 0&, "open", sWebAddress, vbNullString, vbNullString, 1
    Unload Me


End Sub


Private Sub Label1_Click()

End Sub

Private Sub LLink_Click()
     
    Dim sWebAddress As String
    sWebAddress = Me.LLink.Caption
     
    ShellExecute 0&, "open", sWebAddress, vbNullString, vbNullString, 1
    Unload Me

End Sub

