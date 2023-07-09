Attribute VB_Name = "Arkusze"
'***********************************************
'           MZPack - Excell Utility Pack
'        2012, Katarzyna Zbroinska
'             katarzyna@zbroinski.net
'              www.zbroinski.net
'
' This software is distributed under the term of
'     General Public License (GNU) version 3.
'***********************************************
Option Explicit
Private Sub SortujArkusze(ByVal control As IRibbonControl)

Dim strSheetNames() As String
Dim lngIndex As Long
Dim lngSheets As Long
Dim shtActiveSheet As Object
Dim a As Integer
Dim intAnswer As Integer

    On Error Resume Next


    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Czy struktura jest chroniona?
    If ActiveWorkbook.ProtectStructure Then
        MsgBox ActiveWorkbook.Name & " is protected.", _
            vbCritical, "Sheets cannot be sorted."
        Exit Sub
    End If
        
    Call ZamrozEkran
    
    'Zapamietanie aktywnego arkusza
    Set shtActiveSheet = ActiveSheet
    
    'Wczytanie nazw arkuszy do tablicy
    lngSheets = ActiveWorkbook.Sheets.Count
    ReDim strSheetNames(1 To lngSheets)
    For lngIndex = 1 To lngSheets
        strSheetNames(lngIndex) = ActiveWorkbook.Sheets(lngIndex).Name
    Next lngIndex
    
    'Sortowanie tablicy
    Call pomSortujArkuszeQuicksort(strSheetNames, 1, lngSheets)
    
    'Zmiana kolejnosci arkuszy
    For lngIndex = 1 To lngSheets
        ActiveWorkbook.Sheets(strSheetNames(lngIndex)).Move _
            before:=ActiveWorkbook.Sheets(lngIndex)
    Next lngIndex
    
    'Aktywowanie oryginalnego arkusza
    shtActiveSheet.Activate
    
    'Odmroz ekran
    Call OdmrozEkran

End Sub

Private Sub UtworzArkuszeZNazwamiZKomorek(ByVal control As IRibbonControl)

'Makro tworzy nowe arkusze z nazwami pobranymi z zaznaczonych komórek wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-18, Warszawa.

Dim rngDefaultRange As Range
Dim rngUserRanges As Range
Dim rngSubRange As Range
Dim rngCell As Range
Dim shtStartSheet As Worksheet, shtSheet As Worksheet
Dim strName As String
Dim lngIndex As Long
Dim lngLenghtOfFormerName As Long
Dim intAnswer As Integer

Const MSG_TITLE As String = "Create sheets with names from cells"
Const MSG_CONTENTS As String = "Select the range(s):"
Const TYPE_OF_DATA As Integer = 8

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check whether the structure is protected
    If ActiveWorkbook.ProtectStructure Then
        intAnswer = MsgBox("The workbook is protected. Unprotect the workbook before using the macro.", vbOKOnly, "Warning!")
        Exit Sub
    End If
    
    'Check whether the worksheet is protected
    If ActiveSheet.ProtectContents Then
        intAnswer = MsgBox("The worksheet is protected. Uprotect the worksheet before using the macro.", vbOKOnly, "Warning!")
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
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENTS, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngUserRanges = Application.InputBox(prompt:=MSG_CONTENTS, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Tworzenie arkuszy
    Call ZamrozEkran
    
    Set shtStartSheet = ActiveSheet
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If rngCell.Value <> "" Then
                strName = CStr(rngCell.Value)
                If Len(strName) > 31 Then strName = Left(strName, 31)
                Set shtSheet = ActiveWorkbook.Sheets.Add(after:=ActiveSheet)
                shtSheet.Name = strName
                If shtSheet.Name <> strName Then
                    lngLenghtOfFormerName = Len(strName)
                    lngIndex = 0
                    Do
                        lngIndex = lngIndex + 1
                        strName = Left(strName, lngLenghtOfFormerName) & " " & lngIndex
                        If Len(strName) > 31 Then strName = Left(strName, 30 - Len(CStr(lngIndex))) & " " & lngIndex
                        shtSheet.Name = strName
                    Loop Until shtSheet.Name = strName
                End If
            End If
        Next rngCell
    Next rngSubRange
    
    shtStartSheet.Activate
    
Cancelled:
    Call OdmrozEkran
    
End Sub
Private Sub ZmienNazwyArkuszyWgKomorek(ByVal control As IRibbonControl)

'Makro zmienia nazwy arkuszy wg zaznaczonej komórki, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-18, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range
Dim rngCell As Range
Dim shtSheet As Worksheet
Dim strName As String
Dim lngIndex As Long
Dim lngLenghtOfFormerName As Long
Dim intAnswer As Integer

Const MSG_TITLE As String = "Change sheets' names"
Const MSG_CONTENTS As String = "Select the cell with name:"
Const TYPE_OF_DATA As Integer = 8

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

    'Check whether the structure is protected
    If ActiveWorkbook.ProtectStructure Then
        intAnswer = MsgBox("The workbook is protected. Unprotect the workbook before using the macro.", vbOKOnly, "Warning!")
        GoTo Cancelled
    End If

    If ActiveSheet.Type <> xlWorksheet Then
        intAnswer = MsgBox("The active worksheet is not a sheet. Choose a sheet and run the macro again.", vbOKOnly, "Warning!")
        Exit Sub
    End If

    'Znalezienie zakresu defaultowego
    If TypeName(Selection) <> "Range" Or Selection.Count <> 1 Then
        Set rngDefaultRange = Nothing
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    Do
        If rngDefaultRange Is Nothing Then
            Set rngCell = Application.InputBox(prompt:=MSG_CONTENTS, _
                Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
        Else
            Set rngCell = Application.InputBox(prompt:=MSG_CONTENTS, _
                Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
        End If
        If rngCell.Count <> 1 Then lngIndex = MsgBox("You have to select only one cell.", vbOKOnly, "Error")
    Loop Until rngCell.Count = 1

    On Error Resume Next
    
    'Zmiana nazw arkuszy
    Call ZamrozEkran
    
    For Each shtSheet In ActiveWindow.SelectedSheets
        If shtSheet.Range(rngCell.AddressLocal).Value <> "" And shtSheet.Type = xlWorksheet Then
            strName = CStr(shtSheet.Range(rngCell.AddressLocal).Value)
            If Len(strName) > 31 Then strName = Left(strName, 31)
            shtSheet.Name = strName
            If shtSheet.Name <> strName Then
                lngLenghtOfFormerName = Len(strName)
                lngIndex = 0
                Do
                    lngIndex = lngIndex + 1
                    strName = Left(strName, lngLenghtOfFormerName) & " " & lngIndex
                    If Len(strName) > 31 Then strName = Left(strName, 30 - Len(CStr(lngIndex))) & " " & lngIndex
                    shtSheet.Name = strName
                Loop Until shtSheet.Name = strName
            End If
        End If
    Next shtSheet
        
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub WstawNazwyArkuszyDoKomorek(ByVal control As IRibbonControl)

'Makro wstawia nazwy arkuszy do zaznaczonej komórki, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-18, Warszawa.

'Deklaracje zmiennych oraz sta³ych
Dim rngDefaultRange As Range
Dim rngCell As Range
Dim shtSheet As Object
Dim lngIndex As Long
Dim intAnswer As Integer

Const MSG_TITLE As String = "Insert sheets' names into cells"
Const MSG_CONTENTS As String = "Select cell for name:"
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
    If TypeName(Selection) <> "Range" Or Selection.Count <> 1 Then
        Set rngDefaultRange = Nothing
    Else
        Set rngDefaultRange = Selection
    End If
    
    
    'Pytanie o zakres
    On Error GoTo Cancelled
    Do
        If rngDefaultRange Is Nothing Then
            Set rngCell = Application.InputBox(prompt:=MSG_CONTENTS, _
                Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
        Else
            Set rngCell = Application.InputBox(prompt:=MSG_CONTENTS, _
                Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
        End If
        If rngCell.Count <> 1 Then lngIndex = MsgBox("You have to select only one cell.", vbOKOnly, "Error")
    Loop Until rngCell.Count = 1

    On Error Resume Next
    
    'Kopiowanie nazw
    Call ZamrozEkran
    
    For Each shtSheet In ActiveWindow.SelectedSheets
        If shtSheet.Type = xlWorksheet Then
            shtSheet.Range(rngCell.AddressLocal).Value = shtSheet.Name
        End If
    Next shtSheet
        
Cancelled:
    Call OdmrozEkran
    
End Sub

Private Sub UtworzIndeksArkuszy(ByVal control As IRibbonControl)

'Makro tworzy indeks wszystkich lub zaznaczonych arkuszy, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-04-19, Warszawa.

    
'Deklaracje zmiennych oraz sta³ych
Dim shtSheetsCollection As Sheets
Dim shtIndexSheet As Worksheet, shtWorkingSheet As Object
Dim strName As String
Dim lngIndex As Long
Dim intAnswer As Integer

    On Error Resume Next

   
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Call ZamrozEkran
    
    'Check whether the structure is protected
    If ActiveWorkbook.ProtectStructure Then
        intAnswer = MsgBox("The workbook is protected. Unprotect the workbook before using the macro.", vbOKOnly, "Warning!")
        GoTo DefinedExit
    End If

    'Znalezienie kolekcji arkuszy
    If ActiveWindow.SelectedSheets.Count = 1 Then
        Set shtSheetsCollection = Application.Sheets
    Else
        Set shtSheetsCollection = ActiveWindow.SelectedSheets
    End If
       
    'Dodaj nowy arkusz z indeksem
    ActiveWorkbook.Sheets(1).Select
    Set shtIndexSheet = ActiveWorkbook.Sheets.Add(before:=Application.Sheets(1))
    strName = "Sheets Index"
    shtIndexSheet.Name = strName
    If shtIndexSheet.Name <> strName Then
        lngIndex = 0
        Do
            lngIndex = lngIndex + 1
            If lngIndex > 1 Then
                strName = Left(strName, Len(strName) - Len(CStr(lngIndex)) - 1) & " " & lngIndex
            Else
                strName = strName & " " & lngIndex
            End If
            shtIndexSheet.Name = strName
        Loop Until shtIndexSheet.Name = strName
    End If
    
    
    'Dodaj nazwy arkuszy i utwórz linki
    shtIndexSheet.Cells(1, 1).Activate
    ActiveCell.Value = "Sheets' index"
    lngIndex = 0
    For Each shtWorkingSheet In shtSheetsCollection
        If (Not shtWorkingSheet Is shtIndexSheet) And (shtWorkingSheet.Visible = xlSheetVisible) Then
            lngIndex = lngIndex + 1
            If shtWorkingSheet.Type = xlWorksheet Then
                shtIndexSheet.Hyperlinks.Add Anchor:=ActiveCell.Offset(lngIndex, 0), Address:="", SubAddress:="'" & shtWorkingSheet.Name & "'!A1", TextToDisplay:=shtWorkingSheet.Name
            Else
                ActiveCell.Offset(lngIndex, 0).Value = shtWorkingSheet.Name
            End If
        End If
    Next shtWorkingSheet
    
    
    'Sortowanie indeksu
    shtIndexSheet.Cells(1, 1).CurrentRegion.Sort _
        Key1:=Cells(1, 1), _
        Order1:=xlAscending, _
        Orientation:=xlTopToBottom, _
        Header:=xlYes
    shtIndexSheet.Columns("A:A").EntireColumn.AutoFit

DefinedExit:
    Call OdmrozEkran
    
End Sub

Private Sub OdkryjWszystkieArkusze(ByVal control As IRibbonControl)

Dim shtArkusz As Object
Dim intAnswer As Integer

    On Error Resume Next
    
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

    Call ZamrozEkran
    
    'Check whether the structure is protected
    If ActiveWorkbook.ProtectStructure Then
        intAnswer = MsgBox("The workbook is protected. Unprotect the workbook before using the macro.", vbOKOnly, "Warning!")
        GoTo DefinedExit
    End If
    
    For Each shtArkusz In ActiveWorkbook.Sheets
        If Not shtArkusz.Visible = xlSheetVisible Then
            shtArkusz.Visible = xlSheetVisible
        End If
    Next shtArkusz

DefinedExit:
    Call OdmrozEkran

End Sub

Private Sub HideSelectedSheets(ByVal control As IRibbonControl)

Dim shtArkusz As Object
Dim intAnswer As Integer

    On Error Resume Next
    
    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

    Call ZamrozEkran
    
    'Check whether the structure is protected
    If ActiveWorkbook.ProtectStructure Then
        intAnswer = MsgBox("The workbook is protected. Unprotect the workbook before using the macro.", vbOKOnly, "Warning!")
        GoTo DefinedExit
    End If
    
    For Each shtArkusz In ActiveWindow.SelectedSheets
        shtArkusz.Visible = xlSheetHidden
    Next shtArkusz
    
DefinedExit:
    Call OdmrozEkran

End Sub

Private Sub EmptyHeadersAndFootersInSelectedSheets(ByVal control As IRibbonControl)

Dim shtArkusz As Object
Dim intAnswer As Integer

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

    Call ZamrozEkran
        
    For Each shtArkusz In ActiveWindow.SelectedSheets
        With shtArkusz.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Next shtArkusz
    
DefinedExit:
    Call OdmrozEkran

End Sub

Private Sub HideDeeplySelectedSheets(ByVal control As IRibbonControl)

Dim shtArkusz As Object
Dim intAnswer As Integer

    On Error Resume Next

    
    'Czy jest aktywny skoroszyt?
    If ActiveWorkbook Is Nothing Then Exit Sub

    Call ZamrozEkran
    
    'Check whether the structure is protected
    If ActiveWorkbook.ProtectStructure Then
        intAnswer = MsgBox("The workbook is protected. Unprotect the workbook before using the macro.", vbOKOnly, "Warning!")
        GoTo DefinedExit
    End If
    
    For Each shtArkusz In ActiveWindow.SelectedSheets
        shtArkusz.Visible = xlSheetVeryHidden
    Next shtArkusz
    
DefinedExit:
    Call OdmrozEkran
End Sub
Private Sub pomSortujArkuszeQuicksort(Tablica() As String, LewyIndeks As Long, PrawyIndeks As Long)

Dim lngDivider As Long
Dim lngIndex As Long
Dim strElement As String

    If LewyIndeks < PrawyIndeks Then
        lngDivider = LewyIndeks
        For lngIndex = LewyIndeks + 1 To PrawyIndeks
            If Tablica(lngIndex) < Tablica(LewyIndeks) Then
                lngDivider = lngDivider + 1
                strElement = Tablica(lngDivider)
                Tablica(lngDivider) = Tablica(lngIndex)
                Tablica(lngIndex) = strElement
            End If
        Next lngIndex
            
        strElement = Tablica(lngDivider)
        Tablica(lngDivider) = Tablica(LewyIndeks)
        Tablica(LewyIndeks) = strElement
            
        Call pomSortujArkuszeQuicksort(Tablica, LewyIndeks, lngDivider - 1)
        Call pomSortujArkuszeQuicksort(Tablica, lngDivider + 1, PrawyIndeks)
    End If

End Sub


