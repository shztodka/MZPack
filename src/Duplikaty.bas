Attribute VB_Name = "Duplikaty"
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

Type mZaznaczZdublowaneKomorkiPrzetwarzane
    Adres As String
    Wartosc As String
End Type

Public ColorValue As Variant
Dim Buttons(1 To 56) As New ColorButtonClass

Private Sub ZaznaczZdublowaneKomorki(ByVal control As IRibbonControl)

'Makro znajduje i zaznacza wybranym przez u¿ytkownika kolorem
'takie same wartoœci w zaznaczonym zakresie/zakresach, wersja 1.0.
'Autor: Katarzyna  Zbroiñski,
'2009-02-05, Warszawa.
    
    'Deklaracje zmiennych i sta³ych
    Dim rngRange As Range
    Dim strErrorMessage As String
    Dim lngMarkingColor As Long
    Dim udtWorkingTable() As mZaznaczZdublowaneKomorkiPrzetwarzane
    Dim lngIndex As Long
    Dim rngCell As Range
    Dim lngMaxIndexTab As Long, lngMinIndexTab As Long
    Dim blnFlag As Boolean
    Dim rngDefaultRange As Range
    Dim intAnswer As Integer
    
    Const MSG_TITLE As String = "Range"
    Const MSG_CONTENTS As String = "Select the range(s) where the same cells will be marked."
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
        Set rngRange = Application.InputBox(prompt:=MSG_CONTENTS, _
            Title:=MSG_TITLE, Type:=TYPE_OF_DATA)
    Else
        Set rngRange = Application.InputBox(prompt:=MSG_CONTENTS, _
            Title:=MSG_TITLE, Default:=rngDefaultRange.Address(ReferenceStyle:=Application.ReferenceStyle), Type:=TYPE_OF_DATA)
    End If
    
    On Error Resume Next
    
    'Wybór koloru zaznaczenia
    lngMarkingColor = pomGetAColor()
    If lngMarkingColor = False Then
        Exit Sub
    End If
    
    'Wczytanie danych do tablicy
    Call ZamrozEkran
    lngIndex = 1
    ReDim udtWorkingTable(1 To rngRange.Count)
    For Each rngCell In rngRange
        udtWorkingTable(lngIndex).Adres = rngCell.Address
        udtWorkingTable(lngIndex).Wartosc = CStr(rngCell.Value)
        lngIndex = lngIndex + 1
    Next rngCell
    
    'Sortowanie tablicy
    lngMinIndexTab = LBound(udtWorkingTable)
    lngMaxIndexTab = UBound(udtWorkingTable)
    Call pomZaznaczZdublowaneKomorkiQuicksort(udtWorkingTable, lngMinIndexTab, lngMaxIndexTab)
    
    'Zaznaczenie duplikatów
    blnFlag = False
    For lngIndex = lngMinIndexTab + 1 To lngMaxIndexTab
        If udtWorkingTable(lngIndex).Wartosc = udtWorkingTable(lngIndex - 1).Wartosc And udtWorkingTable(lngIndex).Wartosc <> "" Then
            Range(udtWorkingTable(lngIndex - 1).Adres).Interior.Color = lngMarkingColor
            Range(udtWorkingTable(lngIndex).Adres).Interior.Color = lngMarkingColor
            blnFlag = True
        End If
    Next lngIndex
    
    If blnFlag = False Then
        strErrorMessage = "There were no duplicates."
        Call OdmrozEkran
        intAnswer = MsgBox(strErrorMessage, vbInformation, "Warning!")
    End If

Cancelled:
    Call OdmrozEkran
    
End Sub

Private Function pomGetAColor() As Variant
'   Wyœwietla okno dialogowe i zwraca
'   wartoœæ koloru - lub False jeœli ¿aden kolor nie zosta³ wybrany
    Dim ctlMyControl As control
    Dim lngButtonCount As Long
    lngButtonCount = 0
    For Each ctlMyControl In WyborKolorow.Controls
'       56 przycisków kolorów ma ustawionych w³aœciwoœæ Tag na "ColorButton"
        If ctlMyControl.Tag = "ColorButton" Then
            lngButtonCount = lngButtonCount + 1
            Set Buttons(lngButtonCount).ColorButton = ctlMyControl
            If pomWorkbookIsActive Then
'               Pobiera kolor z palety kolorów aktywnego arkusza
                Buttons(lngButtonCount).ColorButton.BackColor = _
                    ActiveWorkbook.Colors(lngButtonCount)
            Else
'               Pobiera kolor z palety kolorów tego arkusza
                Buttons(lngButtonCount).ColorButton.BackColor = _
                    ThisWorkbook.Colors(lngButtonCount)
            End If
        End If
    Next ctlMyControl
    WyborKolorow.Show
    pomGetAColor = ColorValue
End Function

Private Function pomWorkbookIsActive()
'   Zwraca True jeœli istnieje aktywny arkusz
    Dim x As String
    On Error Resume Next
    x = ActiveWorkbook.Name
    If Err = 0 Then
        pomWorkbookIsActive = True
    Else
        pomWorkbookIsActive = False
    End If
    
End Function
Private Sub pomZaznaczZdublowaneKomorkiQuicksort(Tablica() As mZaznaczZdublowaneKomorkiPrzetwarzane, LewyIndeks As Long, PrawyIndeks As Long)

Dim lngDividor As Long
Dim lngIndex As Long
Dim udtElement As mZaznaczZdublowaneKomorkiPrzetwarzane

    If LewyIndeks < PrawyIndeks Then
        lngDividor = LewyIndeks
        For lngIndex = LewyIndeks + 1 To PrawyIndeks
            If Tablica(lngIndex).Wartosc < Tablica(LewyIndeks).Wartosc Then
                lngDividor = lngDividor + 1
                udtElement = Tablica(lngDividor)
                Tablica(lngDividor) = Tablica(lngIndex)
                Tablica(lngIndex) = udtElement
            End If
        Next lngIndex
            
        udtElement = Tablica(lngDividor)
        Tablica(lngDividor) = Tablica(LewyIndeks)
        Tablica(LewyIndeks) = udtElement
            
        Call pomZaznaczZdublowaneKomorkiQuicksort(Tablica, LewyIndeks, lngDividor - 1)
        Call pomZaznaczZdublowaneKomorkiQuicksort(Tablica, lngDividor + 1, PrawyIndeks)
    End If

End Sub




