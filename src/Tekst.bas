Attribute VB_Name = "Tekst"
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

Private Sub DodajTekstNaPoczatek(ByVal control As IRibbonControl)

'Makro dodaje dodatkow¹ treœæ do komórki na pocz¹tku
'ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-13, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim vntAdditionalText As Variant
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Add text at the beginning"
    Const MSG_RANGE As String = "Select cells range(s)::"
    Const MSG_COMMENT As String = "Type the text to be added:"
    Const TYPE_OF_DATA_RELATION As Integer = 8
    Const TYPE_OF_DATA_TEXT As Integer = 2
    
    On Error Resume Next
    
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active woeksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
        
    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    
    'Pytanie o treœæ
    vntAdditionalText = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntAdditionalText = "" Or vntAdditionalText = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie treœci
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = vntAdditionalText & CStr(rngCell.Value)
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub WstawDoKomorki(ByVal control As IRibbonControl)

'Makro dodaje dodatkow¹ treœæ na podan¹ pozycjê
'do ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-13, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim vntText As Variant, strCellText As String
    Dim lngPosition As Long, lngCellLength As Long
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Insert text"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const MSG_COMMENT As String = "Type the text:"
    Const MSG_POSITION As String = "Type the position (character number) where the text will be inserted:"
    Const TYPE_OF_DATA_RELATION As Integer = 8
    Const TYPE_OF_DATA_TEXT As Integer = 2
    Const TYPE_OF_DATA_NUMBER As Integer = 1
    
    On Error Resume Next
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active woeksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
        
    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    
    'Pytanie o treœæ komentarza
    vntText = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntText = "" Or vntText = False Then Exit Sub
    
        
    'Pytanie o pozycjê
    lngPosition = Application.InputBox(prompt:=MSG_POSITION, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_NUMBER)
    If lngPosition = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie komentarza
    Call ZamrozEkran
    
    If lngPosition < 0 Then
        lngPosition = 1
    End If
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                strCellText = CStr(rngCell.Value)
                lngCellLength = Len(strCellText)
                If lngPosition > lngCellLength + 1 Then
                    lngPosition = lngCellLength + 1
                End If
                rngCell.Value = Left(strCellText, lngPosition - 1) & vntText & _
                    Right(strCellText, lngCellLength - lngPosition + 1)
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub DodajTekstNaKoniec(ByVal control As IRibbonControl)

'Makro dodaje dodatkow¹ treœæ na koniec
'do ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-13, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim vntText As Variant
Dim intAnswer As Integer

Const MSG_TITLE As String = "Add text at the end"
Const MSG_RANGE As String = "Select cells range(s):"
Const MSG_COMMENT As String = "Type the text:"
Const TYPE_OF_DATA_RELATION As Integer = 8
Const TYPE_OF_DATA_TEXT As Integer = 2
    
    On Error Resume Next
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
       
    'Check whether active woeksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If

    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    
    'Pytanie o treœæ komentarza
    vntText = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntText = "" Or vntText = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie komentarza
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = CStr(rngCell.Value) & vntText
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub UsunSpacjeNaPoczatku(ByVal control As IRibbonControl)

'Makro usuwa spacje na pocz¹tku
'ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-13, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Remove spaces"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const TYPE_OF_DATA_RELATION As Integer = 8
    
    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active woeksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
        
    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    
    'Usuwanie spacji
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = LTrim(CStr(rngCell.Value))
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub UsunSpacjeNaKoncu(ByVal control As IRibbonControl)

'Makro usuwa spacje na koñcu
'ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-13, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Remove spaces"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const TYPE_OF_DATA_RELATION As Integer = 8
    
    On Error Resume Next
    
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active woeksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
        
    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    
    'Usuwanie spacji
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = RTrim(CStr(rngCell.Value))
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub UsunZnakiPowrotuKaretki(ByVal control As IRibbonControl)

'Makro usuwa znaki powrotu karetki dla
'ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-13, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Remove carriage returns:"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const TYPE_OF_DATA_RELATION As Integer = 8
    
    On Error Resume Next
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active worksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    
    'Usuwanie znakow powrotu karetki
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = Replace(CStr(rngCell.Value), Chr(10), "", , , vbTextCompare)
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub StartFirstWordWithUpperCase(ByVal control As IRibbonControl)

'Macro changes first letter of first word to uppercase
'Version:   1.0
'Author:    Katarzyna  Zbroiñski,
'2011-10-29, Warszawa.

Const MSG_TITLE As String = "Start first word with uppercase:"
Const MSG_RANGE As String = "Select cells range(s):"
Const TYPE_OF_DATA_RELATION As Integer = 8

Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim intAnswer As Integer

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active worksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Find a default range
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    'Ask user for another range
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    'Start first word with uppercase
    Call ZamrozEkran
    On Error Resume Next
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula And Not rngCell.Value = "" Then
                rngCell.Value = UCase(Left(rngCell.Value, 1)) & LCase(Right(rngCell.Value, Len(rngCell.Value) - 1))
            End If
        Next rngCell
    Next rngSubRange
    
Cancelled:
    Call OdmrozEkran

End Sub

Private Sub ChangeToUpperCase(ByVal control As IRibbonControl)

'Macro changes all text in selected range to UPPER case
'Version:   1.0
'Author:    Katarzyna  Zbroiñski,
'2011-10-17, Warszawa.

Const MSG_TITLE As String = "Change to lowercase:"
Const MSG_RANGE As String = "Select cells range(s):"
Const TYPE_OF_DATA_RELATION As Integer = 8

Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim intAnswer As Integer

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active worksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Find a default range
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    'Ask user for another range
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    'Changing to UPPER case
    Call ZamrozEkran
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = UCase(rngCell.Value)
            End If
        Next rngCell
    Next rngSubRange
    
Cancelled:
    Call OdmrozEkran

End Sub


Private Sub ChangeToLowerCase(ByVal control As IRibbonControl)

'Macro changes all text in selected range to lowercase
'Version:   1.0
'Author:    Katarzyna  Zbroiñski,
'2011-10-22, Warszawa.

Const MSG_TITLE As String = "Change to lower case:"
Const MSG_RANGE As String = "Select cells range(s):"
Const TYPE_OF_DATA_RELATION As Integer = 8

Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim intAnswer As Integer
    
    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active worksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Find a default range
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    'Ask user for another range
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    'Changing to UPPER case
    Call ZamrozEkran
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = LCase(rngCell.Value)
            End If
        Next rngCell
    Next rngSubRange
    
Cancelled:
    Call OdmrozEkran

End Sub

Private Sub EachWordWithUppercase(ByVal control As IRibbonControl)

'Macro changes first letter of each word to uppercase
'Version:   1.0
'Author:    Katarzyna  Zbroiñski,
'2011-10-29, Warszawa.

Const MSG_TITLE As String = "Change first letter of each word to capital:"
Const MSG_RANGE As String = "Select cells range(s):"
Const TYPE_OF_DATA_RELATION As Integer = 8

Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim intAnswer As Integer

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether active worksheet is a worksheet
    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Find a default range
    If TypeName(Selection) <> "Range" Then
        Set rngDefaultRange = Nothing
    ElseIf Selection.Count = 1 Then
        Set rngDefaultRange = ActiveCell.CurrentRegion
    Else
        Set rngDefaultRange = Selection
    End If
    
    'Ask user for another range
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    'Changing first letter of each word to capital
    Call ZamrozEkran
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.HasFormula Then
                rngCell.Value = StrConv(rngCell.Value, vbProperCase)
            End If
        Next rngCell
    Next rngSubRange
    
Cancelled:
    Call OdmrozEkran

End Sub

