Attribute VB_Name = "Komentarze"
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
Private Sub UkryjKomentarze(ByVal control As IRibbonControl)

'Makro ukrywa komentarze w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna Zbroiñska,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Hide comments"
    Const MSG_CONTENT As String = "Select cells range(s):"
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
            If Not rngCell.Comment Is Nothing Then
                rngCell.Comment.Visible = False
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub PokazKomentarze(ByVal control As IRibbonControl)

'Makro pokazuje komentarze w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna Zbroiñska,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Show comments"
    Const MSG_CONTENT As String = "Select cells range(s):"
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
    
    'Wyœwietlanie
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If Not rngCell.Comment Is Nothing Then
                rngCell.Comment.Visible = True
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub DodajKomentarzDoWieluKomorek(ByVal control As IRibbonControl)

'Makro dodaje komentarz do ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna Zbroiñska,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim vntComment As Variant
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Add comment(s)"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const MSG_COMMENT As String = "Add text:"
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
    vntComment = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntComment = "" Or vntComment = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie komentarza
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If rngCell.Comment Is Nothing Then
                rngCell.AddComment vntComment
            Else
                rngCell.Comment.Delete
                rngCell.AddComment vntComment
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub DodajDoKomentarzaNaPoczatku(ByVal control As IRibbonControl)

'Makro dodaje dodatkow¹ treœæ do komentarza na pocz¹tku
'do ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim vntComment As Variant
    Dim intAnswer As Integer
    
    Const MSG_TITLE = "Add comment"
    Const MSG_RANGE = "Select cells range(s)::"
    Const MSG_COMMENT = "Add text:"
    Const TYPE_OF_DATA_RELATION = 8
    Const TYPE_OF_DATA_TEXT = 2
    
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
    vntComment = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntComment = "" Or vntComment = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie komentarza
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If rngCell.Comment Is Nothing Then
                rngCell.AddComment vntComment
            Else
                rngCell.Comment.Text Text:=vntComment, Start:=1, Overwrite:=False
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub WstawDoKomentarza(ByVal control As IRibbonControl)

'Makro dodaje dodatkow¹ treœæ do komentarza na podan¹ pozycjê
'do ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim vntComment As Variant
    Dim lngPosition As Long
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Insert comment"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const MSG_COMMENT As String = "Add text:"
    Const MSG_POSITION As String = "Add position (character number) where to insert:"
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
    vntComment = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntComment = "" Or vntComment = False Then Exit Sub
    
        
    'Pytanie o pozycjê
    lngPosition = Application.InputBox(prompt:=MSG_POSITION, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_NUMBER)
    If lngPosition = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie komentarza
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If rngCell.Comment Is Nothing Then
                rngCell.AddComment vntComment
            Else
                rngCell.Comment.Text Text:=vntComment, _
                    Start:=lngPosition, Overwrite:=False
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub DodajDoKomentarzaNaKoniec(ByVal control As IRibbonControl)

'Makro dodaje dodatkow¹ treœæ do komentarza na koniec
'do ka¿dej komórki w zaznaczonym obszarze, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngUserRanges As Range
    Dim rngSubRange As Range
    Dim rngCell As Range
    Dim vntComment As Variant
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Add comment"
    Const MSG_RANGE As String = "Select cells range(s):"
    Const MSG_COMMENT As String = "Add text:"
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
    vntComment = Application.InputBox(prompt:=MSG_COMMENT, _
        Title:=MSG_TITLE, Type:=TYPE_OF_DATA_TEXT)
    If vntComment = "" Or vntComment = False Then Exit Sub
    
    On Error Resume Next
    
    'Dodawanie komentarza
    Call ZamrozEkran
    
    For Each rngSubRange In rngUserRanges.Areas
        For Each rngCell In rngSubRange
            If rngCell.Comment Is Nothing Then
                rngCell.AddComment vntComment
            Else
                rngCell.Comment.Text Text:=vntComment, _
                    Start:=Len(rngCell.Comment.Text) + 1, Overwrite:=False
            End If
        Next rngCell
    Next rngSubRange
    
    
Cancelled:
    Call OdmrozEkran
End Sub

Private Sub UtworzKomentarzeZKomorek(ByVal control As IRibbonControl)

'Makro tworzy komentarze z tekstów zawartych we wskazanych komórkach, wersja 1.0
'Autor: Katarzyna  Zbroiñski,
'2009-03-28, Warszawa.

    'Deklaracje zmiennych oraz sta³ych
    Dim rngDefaultRange As Range
    Dim rngCommentsRange As Range
    Dim rngRangeOfCells As Range
    Dim vntCommentsTexts() As Variant
    Dim rngCell As Range
    Dim lngColumnIndex As Long, lngRowIndex As Long, lngColumns As Long, lngRows As Long
    Dim intAnswer As Integer
     
    Const MSG_TITLE As String = "Create comments from cells"
    Const MSG_COMMENTS As String = "Select range where the comments should be inserted:"
    Const MSG_OF_CELLS As String = "Select range where the comments' text are:"
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
    
    
    'Pytanie o zakres do wstawiania komentarzy
    On Error GoTo Cancelled
    If rngDefaultRange Is Nothing Then
        Set rngCommentsRange = Application.InputBox(prompt:=MSG_COMMENTS, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngCommentsRange = Application.InputBox(prompt:=MSG_COMMENTS, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    
    'Pytanie o zakres komórek z tekstem
    If rngDefaultRange Is Nothing Then
        Set rngRangeOfCells = Application.InputBox(prompt:=MSG_OF_CELLS, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngRangeOfCells = Application.InputBox(prompt:=MSG_OF_CELLS, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Sprawdzenie czy zakresy s¹ równe i czy nie s¹ wielokrotne
    If rngRangeOfCells.Areas.Count > 1 Or rngCommentsRange.Areas.Count > 1 Then
        Call Blad("You can choose only one range.")
        Exit Sub
    End If
    
    lngColumns = rngRangeOfCells.Columns.Count
    lngRows = rngRangeOfCells.Rows.Count
    If lngColumns <> rngCommentsRange.Columns.Count Or lngRows <> rngCommentsRange.Rows.Count Then
          Call Blad("The range of cells with text is not equal to the range for inserting comments.")
        Exit Sub
    End If
    
    
    'Tworzenie komentarzy
    Call ZamrozEkran
    ReDim vntCommentsTexts(1 To lngRows, 1 To lngColumns)
    vntCommentsTexts = rngRangeOfCells.Value
    For lngRowIndex = 1 To lngRows
        For lngColumnIndex = 1 To lngColumns
            If rngCommentsRange.Cells(lngRowIndex, lngColumnIndex).Comment Is Nothing Then
                rngCommentsRange.Cells(lngRowIndex, lngColumnIndex).AddComment (CStr(vntCommentsTexts(lngRowIndex, lngColumnIndex)))
            Else
                rngCommentsRange.Cells(lngRowIndex, lngColumnIndex).Comment.Delete
                rngCommentsRange.Cells(lngRowIndex, lngColumnIndex).AddComment (CStr(vntCommentsTexts(lngRowIndex, lngColumnIndex)))
            End If
        Next lngColumnIndex
    Next lngRowIndex
    
    
Cancelled:
    Call OdmrozEkran
End Sub


