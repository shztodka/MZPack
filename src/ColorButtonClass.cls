VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorButtonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ColorButton As MSForms.CommandButton
Attribute ColorButton.VB_VarHelpID = -1

Private Sub ColorButton_MouseDown(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    'Przypisuje kolor i zamyka okno dialogowe
    ColorValue = ColorButton.BackColor
    Unload WyborKolorow
    
End Sub


