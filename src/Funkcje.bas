Attribute VB_Name = "Funkcje"
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

Public Function ISBOLDED(Cell As Range) As Boolean
Attribute ISBOLDED.VB_Description = "Returns TRUE if text in a given cell is bolded - otherwise returns FALSE."
Attribute ISBOLDED.VB_ProcData.VB_Invoke_Func = " \n9"


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        ISBOLDED = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Sprawdz czy jest sformatowane zgodnie z zyczeniem
    If Cell.Font.Bold = True Then
        ISBOLDED = True
    Else
        ISBOLDED = False
    End If
    
End Function

Public Function CONCATENATEWITHSEP(CellsRange As Range, Optional Separator As String, Optional SkipBlanks As Boolean = True) As String
Attribute CONCATENATEWITHSEP.VB_Description = "The function returns string wcich is concatenation of the strings from given cells and separated withe the optional Separator. If SkipBlanks is omitted it will be set to TRUE by default."
Attribute CONCATENATEWITHSEP.VB_ProcData.VB_Invoke_Func = " \n7"

Dim rngCell As Range


    'Sprawdz argumenty
    If TypeName(CellsRange) <> "Range" Then
        CONCATENATEWITHSEP = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Oblicz wartosc funkcji
    CONCATENATEWITHSEP = ""
    
    For Each rngCell In CellsRange
        If rngCell.Value <> "" Or Not SkipBlanks Then
            CONCATENATEWITHSEP = CONCATENATEWITHSEP & Separator & rngCell.Value
        End If
    Next rngCell
    
    If CONCATENATEWITHSEP <> "" Then
        CONCATENATEWITHSEP = Mid(CONCATENATEWITHSEP, Len(Separator) + 1, Len(CONCATENATEWITHSEP) - Len(Separator))
    End If

End Function


Public Function EXTRACTNUMBER(Cell As Range, Optional Separator As String) As Double
Attribute EXTRACTNUMBER.VB_Description = "Returns the number which is hidden between letters in a given cell. The optional argument SEPARATOR will be set to default one if omitted."
Attribute EXTRACTNUMBER.VB_ProcData.VB_Invoke_Func = " \n3"

Dim lngCharNumber As Long, lngIndex As Long
Dim strChar As String, strNumber As String, strDefaultSeparator As String


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        EXTRACTNUMBER = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Sprawdz czy podany separator jest jednoznakowy
    If Len(Separator) > 1 Then
        EXTRACTNUMBER = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Okreœl separator dziesiêtny
    If Separator = "" Then
        Separator = Application.DecimalSeparator
    End If
    strDefaultSeparator = Application.DecimalSeparator
    
    'Obliczenie wartoœci funkcji
    strNumber = ""
    
    lngCharNumber = Len(Cell.Value)
    For lngIndex = 1 To lngCharNumber
        strChar = Mid(Cell.Value, lngIndex, 1)
        If IsNumeric(strChar) Then
            strNumber = strNumber & strChar
        ElseIf strChar = Separator And lngIndex > 1 And lngIndex < lngCharNumber Then
            If IsNumeric(Mid(Cell.Value, lngIndex - 1, 1)) And IsNumeric(Mid(Cell.Value, lngIndex + 1, 1)) Then
                strNumber = strNumber & strDefaultSeparator
            End If
        End If
    Next lngIndex
    
    EXTRACTNUMBER = CDbl(strNumber)
    
End Function
Public Function ISITALICS(Cell As Range) As Boolean
Attribute ISITALICS.VB_Description = "Returns TRUE if text in a given cell is in italics - otherwise returns FALSE."
Attribute ISITALICS.VB_ProcData.VB_Invoke_Func = " \n9"

    
    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        ISITALICS = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Sprawdz czy jest sformatowane zgodnie z zyczeniem
    If Cell.Font.Italic = True Then
        ISITALICS = True
    Else
        ISITALICS = False
    End If
    
End Function
Public Function GETCOMMENT(Cell As Range) As String
Attribute GETCOMMENT.VB_Description = "Returns the text from a comment in a given cell."
Attribute GETCOMMENT.VB_ProcData.VB_Invoke_Func = " \n9"

    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETCOMMENT = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Oblicz wartosc komorki
    If Not Cell.Comment Is Nothing Then
        GETCOMMENT = Cell.Comment.Text
    Else
        GETCOMMENT = ""
    End If
    
End Function

Public Function EXTRACTFILENAME(Cell As Range) As String
Attribute EXTRACTFILENAME.VB_Description = "Returns file name from a give path. If the cell does not contain a path the function will return empty string ("")."
Attribute EXTRACTFILENAME.VB_ProcData.VB_Invoke_Func = " \n9"

Dim strCurrentChar As String
Dim lngIndex As Long
Dim blnFlag As Boolean


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        EXTRACTFILENAME = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Obliczanie wartoœci funkcji
    EXTRACTFILENAME = ""
    blnFlag = False
    lngIndex = Len(Cell.Value)
    
    If lngIndex > 0 Then
        strCurrentChar = Mid(Cell.Value, lngIndex, 1)
    End If
    Do Until strCurrentChar = "\" Or lngIndex = 0
        EXTRACTFILENAME = strCurrentChar & EXTRACTFILENAME
        lngIndex = lngIndex - 1
        If lngIndex > 0 Then
            strCurrentChar = Mid(Cell.Value, lngIndex, 1)
        End If
        If strCurrentChar = "\" Then
            blnFlag = True
        End If
    Loop
    
    If blnFlag = False Then
        EXTRACTFILENAME = ""
    End If
    
End Function

Public Function WEEKDAYNAME(Cell As Range) As String

Dim DayNumber As Integer

    'Sprawdz czy zaznaczona jest tylko jedna Cell
'    If Cell.Count <> 1 Then
'        WEEKDAYNAME = CVErr(xlErrRef)
'        Exit Function
'    End If
    
    'Sprawdz czy wartosc komorki to jest data
    If Not IsDate(Cell.Value) Then
        WEEKDAYNAME = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Obliczanie wartosci funkcji
    DayNumber = Weekday(Cell.Value, vbMonday)
    Select Case DayNumber
        Case 1
            WEEKDAYNAME = "Monday"
        Case 2
            WEEKDAYNAME = "Tuesday"
        Case 3
            WEEKDAYNAME = "Wednesday"
        Case 4
            WEEKDAYNAME = "Thursday"
        Case 5
            WEEKDAYNAME = "Friday"
        Case 6
            WEEKDAYNAME = "Saturday"
        Case 7
            WEEKDAYNAME = "Sunday"
    End Select
    
End Function
Public Function EXTRACTFOLDERNAME(Cell As Range) As String
Attribute EXTRACTFOLDERNAME.VB_Description = "Returns full folder name from a cell with a path. If the cell does not contain a path the function will return empty string ("")."
Attribute EXTRACTFOLDERNAME.VB_ProcData.VB_Invoke_Func = " \n9"

Dim lngIndex As Long
Dim strCurrentChar As String


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        EXTRACTFOLDERNAME = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Obliczanie wartoœci funkcji
    EXTRACTFOLDERNAME = ""
    lngIndex = Len(Cell.Value)
    
    If lngIndex > 0 Then
        Do
            strCurrentChar = Mid(Cell.Value, lngIndex, 1)
            lngIndex = lngIndex - 1
        Loop Until lngIndex = 0 Or strCurrentChar = "\"
    End If
    If lngIndex > 0 Then
        EXTRACTFOLDERNAME = Left(Cell.Value, lngIndex)
    End If
    
End Function

Public Function REMOVEDIGITS(Cell As Range) As String
Attribute REMOVEDIGITS.VB_Description = "Removes all numbers from a cell leaving all the other signs (letters, symbols etc.)."
Attribute REMOVEDIGITS.VB_ProcData.VB_Invoke_Func = " \n7"

Dim lngCounter As Long, lngCharNumber As Long
Dim strChar As String


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        REMOVEDIGITS = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Oblicz wartosc komorki
    REMOVEDIGITS = ""
    lngCharNumber = Len(Cell.Value)
    For lngCounter = 1 To lngCharNumber
        strChar = Mid(Cell.Value, lngCounter, 1)
        If Not IsNumeric(strChar) Then
            REMOVEDIGITS = REMOVEDIGITS & strChar
        End If
    Next lngCounter
    
End Function
Public Function ISUNDERLINED(Cell As Range) As Boolean
Attribute ISUNDERLINED.VB_Description = "Returns TRUE if text in a given cell is underlined - otherwise returns FALSE"
Attribute ISUNDERLINED.VB_ProcData.VB_Invoke_Func = " \n9"

    
    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        ISUNDERLINED = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Sprawdz czy jest sformatowane zgodnie z zyczeniem
    If Cell.Font.Underline <> xlUnderlineStyleNone Then
        ISUNDERLINED = True
    Else
        ISUNDERLINED = False
    End If
    
End Function
Public Function GETFONTCOLOR(Cell As Range) As Integer
Attribute GETFONTCOLOR.VB_Description = "Returns font's colour index from the standard colour pallette."
Attribute GETFONTCOLOR.VB_ProcData.VB_Invoke_Func = " \n9"
    
    
    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETFONTCOLOR = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    GETFONTCOLOR = Cell.Font.ColorIndex
    
End Function

Public Function GETBGCOLOR(Cell As Range) As Integer
Attribute GETBGCOLOR.VB_Description = "Returns background's colour index from the standard colour pallette. In case the colour is not set the function will return negative value."
Attribute GETBGCOLOR.VB_ProcData.VB_Invoke_Func = " \n9"


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETBGCOLOR = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    GETBGCOLOR = Cell.Interior.ColorIndex
    
End Function

Public Function GETNUMBERFORMAT(Cell As Range) As Variant
Attribute GETNUMBERFORMAT.VB_Description = "Returns number format which was set for the given cell."
Attribute GETNUMBERFORMAT.VB_ProcData.VB_Invoke_Func = " \n9"


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETNUMBERFORMAT = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    GETNUMBERFORMAT = Cell.NumberFormat
    
End Function
Public Function GETFORMULA(Cell As Range) As String
Attribute GETFORMULA.VB_Description = "Returns formula which is in the given cell. In case the cell does not contain a formula the function will return empty string ("")."
Attribute GETFORMULA.VB_ProcData.VB_Invoke_Func = " \n9"


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETFORMULA = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    If Left(Cell.FormulaLocal, 1) = "=" Then
        GETFORMULA = Cell.FormulaLocal
    Else
        GETFORMULA = ""
    End If
    
End Function
Public Function GETHYPERLINK(Cell As Range) As String
Attribute GETHYPERLINK.VB_Description = "Returns text which is the addres of hyperlink inserted in the given cell. If the cell does not contain a hyperlink the function will return empty string ("")."
Attribute GETHYPERLINK.VB_ProcData.VB_Invoke_Func = " \n9"


    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Or TypeName(Cell) <> "Range" Then
        GETHYPERLINK = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    If Cell.Hyperlinks.Count = 0 Then
        GETHYPERLINK = ""
        Exit Function
    End If
    
    If Cell.Hyperlinks(1).SubAddress = "" Then
        GETHYPERLINK = Cell.Hyperlinks(1).Address
    Else
        GETHYPERLINK = Cell.Hyperlinks(1).SubAddress
    End If
    
End Function

Public Function GETSHEETNAME() As String
Attribute GETSHEETNAME.VB_Description = "Returns the sheet name where the function is used."
Attribute GETSHEETNAME.VB_ProcData.VB_Invoke_Func = " \n9"


    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Oblicz wartoœæ funkcji
    GETSHEETNAME = Application.ActiveSheet.Name

End Function


Public Function GETFONTNAME(Cell As Range) As String
Attribute GETFONTNAME.VB_Description = "Returns font's name used in a given cell."
Attribute GETFONTNAME.VB_ProcData.VB_Invoke_Func = " \n9"
    
    
    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETFONTNAME = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    GETFONTNAME = Cell.Font.Name

End Function

Public Function GETWBKFULLNAME() As String
Attribute GETWBKFULLNAME.VB_Description = "Returns the full workbook's name (together with the path) where the function is used."
Attribute GETWBKFULLNAME.VB_ProcData.VB_Invoke_Func = " \n9"

    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    GETWBKFULLNAME = Application.ActiveWorkbook.FullName
    
End Function

Public Function GETPATH() As String
Attribute GETPATH.VB_Description = "Returns the path to the workbook where the function is used but without the workbook's name."
Attribute GETPATH.VB_ProcData.VB_Invoke_Func = " \n9"

    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    GETPATH = Application.ActiveWorkbook.Path
    
End Function

Public Function GETWKBNAME() As String
Attribute GETWKBNAME.VB_Description = "Returns the workbook's name where the function is used (without the path)."
Attribute GETWKBNAME.VB_ProcData.VB_Invoke_Func = " \n9"

    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    GETWKBNAME = Application.ActiveWorkbook.Name
    
End Function


Public Function GETFONTSIZE(Cell As Range) As Variant
Attribute GETFONTSIZE.VB_Description = "Returns the font's size used in a given cell."
Attribute GETFONTSIZE.VB_ProcData.VB_Invoke_Func = " \n9"

    
    'Sprawdz czy zaznaczona jest tylko jedna Cell
    If Cell.Count <> 1 Then
        GETFONTSIZE = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Przypisz wartoœæ funkcji
    GETFONTSIZE = Cell.Font.Size

End Function


Public Function SUMIFFONTCOLOR(CellsRange As Range, ColourNumber As Integer) As Variant
Attribute SUMIFFONTCOLOR.VB_Description = "Sums values from given cells where font's colour is the same as ColourNumber."
Attribute SUMIFFONTCOLOR.VB_ProcData.VB_Invoke_Func = " \n3"
   
'Deklaracja zmiennych
Dim rngCell As Range

    
    'Sprawdzenie poprawnoœci argumentów
    If TypeName(CellsRange) <> "Range" Then
        SUMIFFONTCOLOR = CVErr(xlErrRef)
        Exit Function
    End If
    
    If ColourNumber > 56 Then
        SUMIFFONTCOLOR = CVErr(xlErrNum)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    SUMIFFONTCOLOR = 0
    For Each rngCell In CellsRange.Cells
        If IsNumeric(rngCell.Value) Then
            If rngCell.Font.ColorIndex = ColourNumber Then
                SUMIFFONTCOLOR = SUMIFFONTCOLOR + rngCell.Value
            End If
        End If
    Next rngCell

End Function
Public Function COUNTIFFONTCOLOR(CellsRange As Range, ColourNumber As Integer) As Variant
Attribute COUNTIFFONTCOLOR.VB_Description = "Counts cells in selected range where font colour is the same as given ColourNumber."
Attribute COUNTIFFONTCOLOR.VB_ProcData.VB_Invoke_Func = " \n3"
   
'Deklaracja zmiennych
Dim rngCell As Range

    
    'Sprawdzenie poprawnoœci argumentów
    If TypeName(CellsRange) <> "Range" Then
        COUNTIFFONTCOLOR = CVErr(xlErrRef)
        Exit Function
    End If
    
    If ColourNumber > 56 Then
        COUNTIFFONTCOLOR = CVErr(xlErrNum)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    COUNTIFFONTCOLOR = 0
    For Each rngCell In CellsRange.Cells
        If rngCell.Font.ColorIndex = ColourNumber Then
            COUNTIFFONTCOLOR = COUNTIFFONTCOLOR + 1
        End If
    Next rngCell

End Function

Public Function SUMIFBGCOLOR(CellsRange As Range, ColourNumber As Integer) As Variant
Attribute SUMIFBGCOLOR.VB_Description = "Sums values from given cells where background's colour is the same as ColourNumber."
Attribute SUMIFBGCOLOR.VB_ProcData.VB_Invoke_Func = " \n3"
   
'Deklaracja zmiennych
Dim rngCell As Range

    
    'Sprawdzenie poprawnoœci argumentów
    If TypeName(CellsRange) <> "Range" Then
        SUMIFBGCOLOR = CVErr(xlErrRef)
        Exit Function
    End If
    
    If ColourNumber > 56 Then
        SUMIFBGCOLOR = CVErr(xlErrNum)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    SUMIFBGCOLOR = 0
    For Each rngCell In CellsRange.Cells
        If IsNumeric(rngCell.Value) Then
            If rngCell.Interior.ColorIndex = ColourNumber Then
                SUMIFBGCOLOR = SUMIFBGCOLOR + rngCell.Value
            End If
        End If
    Next rngCell

End Function
Public Function COUNTIFBGCOLOR(CellsRange As Range, ColourNumber As Integer) As Variant
Attribute COUNTIFBGCOLOR.VB_Description = "Counts cells from selected range where background colour is the same as given ColourNumber."
Attribute COUNTIFBGCOLOR.VB_ProcData.VB_Invoke_Func = " \n3"
   
'Deklaracja zmiennych
Dim rngCell As Range
    
    
    'Sprawdzenie poprawnoœci argumentów
    If TypeName(CellsRange) <> "Range" Then
        COUNTIFBGCOLOR = CVErr(xlErrRef)
        Exit Function
    End If
    
    If ColourNumber > 56 Then
        COUNTIFBGCOLOR = CVErr(xlErrNum)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    COUNTIFBGCOLOR = 0
    For Each rngCell In CellsRange.Cells
        If rngCell.Interior.ColorIndex = ColourNumber Then
            COUNTIFBGCOLOR = COUNTIFBGCOLOR + 1
        End If
    Next rngCell

End Function
Public Function COUNTCOLOUREDBG(CellsRange As Range) As Variant
Attribute COUNTCOLOUREDBG.VB_Description = "Counts cells from selected range where background colour is different than default (no colour)."
Attribute COUNTCOLOUREDBG.VB_ProcData.VB_Invoke_Func = " \n3"
   
'Deklaracja zmiennych
Dim rngCell As Range


    'Sprawdzenie poprawnoœci argumentów
    If TypeName(CellsRange) <> "Range" Then
        COUNTCOLOUREDBG = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    COUNTCOLOUREDBG = 0
    For Each rngCell In CellsRange.Cells
        If rngCell.Interior.ColorIndex > 0 Then
            COUNTCOLOUREDBG = COUNTCOLOUREDBG + 1
        End If
    Next rngCell

End Function
Public Function COUNTCOLOUREDFONTS(CellsRange As Range) As Variant
Attribute COUNTCOLOUREDFONTS.VB_Description = "Counts cells from selected range where font colour is different than default one (black)."
Attribute COUNTCOLOUREDFONTS.VB_ProcData.VB_Invoke_Func = " \n3"
   
'Deklaracja zmiennych
Dim rngCell As Range
    
    
    'Sprawdzenie poprawnoœci argumentów
    If TypeName(CellsRange) <> "Range" Then
        COUNTCOLOUREDFONTS = CVErr(xlErrRef)
        Exit Function
    End If
    
    'Wlaczenie automatycznego przeliczenia przy kazdorazowej aktualizacji danych
    Application.Volatile True
    
    'Obliczenie wartoœci funkcji
    COUNTCOLOUREDFONTS = 0
    For Each rngCell In CellsRange.Cells
        If rngCell.Font.ColorIndex <> 1 And rngCell.Font.ColorIndex <> -4105 Then
            COUNTCOLOUREDFONTS = COUNTCOLOUREDFONTS + 1
        End If
    Next rngCell

End Function

Public Function SHEETSCOUNT() As Long
Attribute SHEETSCOUNT.VB_Description = "Returns the number of sheets in the active workbook."
Attribute SHEETSCOUNT.VB_ProcData.VB_Invoke_Func = " \n9"

    SHEETSCOUNT = ActiveWorkbook.Sheets.Count

End Function

Public Function GETUSERNAME() As String
Attribute GETUSERNAME.VB_Description = "Returns the name of the current user."
Attribute GETUSERNAME.VB_ProcData.VB_Invoke_Func = " \n9"

    GETUSERNAME = Application.UserName

End Function

Public Function STATICRND() As Single
Attribute STATICRND.VB_Description = "Returns a random number that doesn't change when the worksheet is recalculated."
Attribute STATICRND.VB_ProcData.VB_Invoke_Func = " \n3"

    Application.Volatile (False)
    Randomize
    STATICRND = Rnd(1)
    
End Function

Public Function STATICRNDBETWEEN(LowerBound As Long, UpperBound As Long) As Long
Attribute STATICRNDBETWEEN.VB_Description = "Returns a random number from a given range that doesn't change when the worksheet is recalculated."
Attribute STATICRNDBETWEEN.VB_ProcData.VB_Invoke_Func = " \n3"

    Application.Volatile (False)
    Randomize
    STATICRNDBETWEEN = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
    
End Function

Public Function FILEEXISTS(Path As String) As Boolean
Attribute FILEEXISTS.VB_Description = "Returns TRUE if a specified file exists. Otherwise, returns FALSE."
Attribute FILEEXISTS.VB_ProcData.VB_Invoke_Func = " \n9"
    
    'Function's body
    On Error Resume Next
    FILEEXISTS = (Dir(Path) <> "")
    If Err <> 0 Then
        FILEEXISTS = False
    End If

    
End Function

Public Function PATHEXISTS(Path As String) As Boolean
Attribute PATHEXISTS.VB_Description = "Returns TRUE if a specified path exists. Othwerwise, returns FALSE."
Attribute PATHEXISTS.VB_ProcData.VB_Invoke_Func = " \n9"
    
    'Function's body
    On Error Resume Next
    If Dir(Path, vbDirectory) = "" Then
        PATHEXISTS = False
    Else
        PATHEXISTS = (GetAttr(Path) And vbDirectory) = vbDirectory
    End If
    If Err <> 0 Then
        PATHEXISTS = False
    End If
    
End Function

Public Function DAYSINMONTH(GivenDate As Date) As Integer
Attribute DAYSINMONTH.VB_Description = "Returns the number of days in the month for a given date."
Attribute DAYSINMONTH.VB_ProcData.VB_Invoke_Func = " \n2"

Dim datMyDate As Date

    'Function's body
    datMyDate = CDate(GivenDate)
    DAYSINMONTH = Day(DateSerial(Year(datMyDate), Month(datMyDate) + 1, 0))

End Function

