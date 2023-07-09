Attribute VB_Name = "pomWspolne"
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

Private Sub WyczyscRejestr()

    DeleteSetting MYAPPNAME, VERSION

End Sub

Sub ZamrozEkran()

    'Remember current settings
    With Application
        blnScreenUpdatingStatus = .ScreenUpdating
        enmCalculationStatus = .Calculation
        blnEnableEventsStatus = .EnableEvents
    End With
    
    'Switch them off for better performance
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .Cursor = xlWait
        .StatusBar = "Wait..."
    End With
    
End Sub

Sub Blad(strMessage As String)

Dim intAnswer As Integer
    
    strMessage = strMessage & vbCrLf & "Run the macro again."
    Call OdmrozEkran
    intAnswer = MsgBox(strMessage, vbInformation, "Warning!")

End Sub



Sub OdmrozEkran()

    With Application
        .ScreenUpdating = blnScreenUpdatingStatus
        .Calculation = enmCalculationStatus
        .EnableEvents = blnEnableEventsStatus
        .Cursor = xlDefault
        .StatusBar = False
    End With

End Sub

Sub PokazODodatku(ByVal control As IRibbonControl)

    ODodatku.Show
    
End Sub

Private Sub PokazRejestr()

Dim MySettings As Variant, intSettings As Integer

    ' Retrieve the settings.
    MySettings = GetAllSettings(appname:=MYAPPNAME, section:=VERSION)
    For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
        Debug.Print MySettings(intSettings, 0), MySettings(intSettings, 1)
    Next intSettings

End Sub

