VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WyborKolorow 
   Caption         =   "Choose color"
   ClientHeight    =   2772
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2640
   OleObjectBlob   =   "WyborKolorow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WyborKolorow"
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

Private Sub CancelButton_Click()
    ColorValue = False
    Unload Me
End Sub
