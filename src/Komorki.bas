Attribute VB_Name = "Komorki"
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

Private Sub ChangeFormulasIntoValues(ByVal control As IRibbonControl)

'Macro changes all formulas in selected range into its values
'Version:   1.0
'Author:    Katarzyna  Zbroiñski,
'2011-10-22, Warszawa.

Const MSG_TITLE As String = "Change into values:"
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
            If rngCell.HasFormula Then
                rngCell.Value = rngCell.Value
            End If
        Next rngCell
    Next rngSubRange
    
Cancelled:
    Call OdmrozEkran

End Sub
Private Sub FormatujListe(ByVal control As IRibbonControl)

'Makro formatuje listê, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-02-04, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim strFont As String
Dim lngFontSize As Long
Dim rngDefaultRange As Range, rngUserRanges As Range, rngSingularList As Range
Dim intAnswer As Integer

Const MSG_TITLE As String = "Range"
Const MSG_CONTENT As String = "Select the range with list to be formatted:"
Const TYPE_OF_DATA As Integer = 8

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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Formatowanie
    Call ZamrozEkran
    
    strFont = Application.StandardFont
    lngFontSize = Application.StandardFontSize
    
    For Each rngSingularList In rngUserRanges.Areas
        With rngSingularList
            .Font.Name = strFont
            .Font.Size = lngFontSize
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .Weight = xlThin
            End With
            
            'Nag³ówek
            With .Range(Cells(1, 1), Cells(1, .Columns.Count))
                .Font.Bold = True
                .Interior.ColorIndex = 1
                .Interior.Pattern = xlSolid
                .Font.ColorIndex = 2
                .HorizontalAlignment = xlCenter
                
                'Szerokoœæ kolumn
                .EntireColumn.AutoFit
            End With
        End With
    Next rngSingularList
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub SelectedRangeIntoNewWorkbook(ByVal control As IRibbonControl)

'The macro creates new workbook and pastes previously selected range into it, version 1.0
'Author: Katarzyna  Zbroiñski,
'2011-11-09, Warsaw.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range, rngUserRanges As Range
Dim intAnswer As Integer
Dim objNewWorkbook As Workbook

Const MSG_TITLE As String = "Range"
Const MSG_CONTENT As String = "Select the range to be copied into new workbook:"
Const TYPE_OF_DATA As Integer = 8

    On Error Resume Next
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

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
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    If rngUserRanges Is Nothing Then Exit Sub
    
    If rngUserRanges.Areas.Count > 1 Then
        intAnswer = MsgBox("You can select only one area. Run the macro once more.", vbOKOnly, "Error!")
        Exit Sub
    End If
    
    'Copying and creating new workbook
    Call ZamrozEkran
    On Error GoTo Cancelled
    
    rngUserRanges.Copy
    Set objNewWorkbook = Workbooks.Add
    With objNewWorkbook.ActiveSheet
        .Paste
        .Cells(1, 1).Select
    End With
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub DeleteHiddenRows(ByVal control As IRibbonControl)

'Macro removes all hidden rows
'Author: Katarzyna  Zbroiñski,
'2011-11-01, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range, rngUserRanges As Range, rngSingularList As Range
Dim lngRow As Long
Dim lngRemovedRows As Long
Dim intAnswer As Integer

Const MSG_TITLE As String = "Range"
Const MSG_CONTENT As String = "Select the range with hidden rows:"
Const TYPE_OF_DATA As Integer = 8

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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Removing hidden rows
    Call ZamrozEkran
    
    lngRemovedRows = 0
    For Each rngSingularList In rngUserRanges.Areas
        lngRow = 1
        With rngSingularList
            Do While lngRow <= rngSingularList.Rows.Count
                With .Cells(lngRow, 1)
                    If .EntireRow.Hidden Then
                        .EntireRow.Delete
                        lngRemovedRows = lngRemovedRows + 1
                    Else
                        lngRow = lngRow + 1
                    End If
                End With
            Loop
        End With
    Next rngSingularList
    intAnswer = MsgBox("Removed: " & lngRemovedRows & " row(s).", vbOKOnly, "Information")
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub DeleteHiddenColumns(ByVal control As IRibbonControl)

'Macro removes all hidden columns
'Author: Katarzyna  Zbroiñski,
'2011-11-01, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range, rngUserRanges As Range, rngSingularList As Range
Dim lngColumn As Long
Dim lngRemovedColumns As Long
Dim intAnswer As Integer

Const MSG_TITLE As String = "Range"
Const MSG_CONTENT As String = "Select the range with hidden columns:"
Const TYPE_OF_DATA As Integer = 8

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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Removing hidden columns
    Call ZamrozEkran
    
    lngRemovedColumns = 0
    For Each rngSingularList In rngUserRanges.Areas
        lngColumn = 1
        With rngSingularList
            Do While lngColumn <= rngSingularList.Columns.Count
                With .Cells(1, lngColumn)
                    If .EntireColumn.Hidden Then
                        .EntireColumn.Delete
                        lngRemovedColumns = lngRemovedColumns + 1
                    Else
                        lngColumn = lngColumn + 1
                    End If
                End With
            Loop
        End With
    Next rngSingularList
    intAnswer = MsgBox("Removed: " & lngRemovedColumns & " column(s).", vbOKOnly, "Information")
    
Cancelled:
    Call OdmrozEkran
End Sub



Private Sub ChangeNumbersSigns(ByVal control As IRibbonControl)

'The macro changes positive numbers to negative and vice versa
'Author: Katarzyna  Zbroiñski,
'2011-10-29, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range, rngUserRanges As Range, rngSingularList As Range, rngCell As Range
Dim intAnswer As Integer

Const MSG_TITLE As String = "Range"
Const MSG_CONTENT As String = "Select the range for changing the numbers' signs:"
Const TYPE_OF_DATA As Integer = 8

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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Changing signs
    Call ZamrozEkran
    
    For Each rngSingularList In rngUserRanges.Areas
        For Each rngCell In rngSingularList
            If Not rngCell.HasFormula And Not rngCell.Value = "" Then
                If IsNumeric(rngCell.Value) Then
                    rngCell.Value = -rngCell.Value
                End If
            End If
        Next rngCell
    Next rngSingularList
    
    
Cancelled:
    Call OdmrozEkran
End Sub


Private Sub ZastosujFormule(ByVal control As IRibbonControl)

'Makro dodaje formu³ê do
'ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-15, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim vntNewOperation As Variant
Dim intAnswer As Integer

Const MSG_TITLE = "Apply formula"
Const MSG_RANGE = "Select cells range(s)::"
Const MSG_FORMULA = "Type the formula:"
Const TYPE_OF_DATA_RELATION = 8
Const TYPE_OF_DATA_FORMULA = 2

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
    
    
    'Pytanie o zakres i formu³ê
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA_RELATION)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_RANGE, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA_RELATION)
    End If
    
    vntNewOperation = Application.InputBox(prompt:=MSG_FORMULA, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_FORMULA)
    
    If vntNewOperation = "" Or vntNewOperation = False Then Exit Sub
    
    On Error Resume Next
    
    'Usuwanie znakow powrotu karetki
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            With rngCell
                If .Value <> 0 And IsNumeric(.Value) Then
                    If .HasFormula Then
                        If .HasArray Then
                            .FormulaArray = "=(" & Mid(.FormulaLocal, 2) & ")" & vntNewOperation
                        Else
                            .FormulaLocal = "=(" & Mid(.FormulaLocal, 2) & ")" & vntNewOperation
                        End If
                    Else
                            .FormulaLocal = "=" & CStr(.Value) & vntNewOperation
                    End If
                End If
            End With
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub KonwertujTekstNaLiczbe(ByVal control As IRibbonControl)

'Makro konwertuje liczby rozpoznawane jako tekst na wartoœci numeryczne, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-03-25, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim intAnswer As Integer

Const MSG_TITLE As String = "Convert text into number"
Const MSG_CONTENT As String = "Select cells range(s) for conversion:"
Const TYPE_OF_DATA As Integer = 8

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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Konwersja
    Call ZamrozEkran
    
    On Error Resume Next
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If rngCell.Value <> "" Then
                rngCell.Value = CDbl(rngCell.Value)
            End If
        Next rngCell
    Next rngSubRange
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub KonwertujLiczbeNaTekst(ByVal control As IRibbonControl)

'Makro konwertuje liczbê na jej reprezentacjê tekstow¹, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-03-23, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim intAnswer As Integer

Const MSG_TITLE As String = "Convert number into text"
Const MSG_CONTENT As String = "Select cells range(s) for conversion:"
Const TYPE_OF_DATA As Integer = 8

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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENT, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Konwersja
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If IsNumeric(rngCell.Value) And rngCell.Value <> "" Then
                rngCell.Value = "'" & CStr(rngCell.Value)
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub



