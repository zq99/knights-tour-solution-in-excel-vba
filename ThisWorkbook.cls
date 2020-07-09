VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim oFont As New clsFontInstall

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
    Set oFont = Nothing
End Sub

Private Sub Workbook_Open()
On Error GoTo ERR_HANDLER:
    oFont.FontName = "Chess Alpha"
    oFont.FontFileName = ThisWorkbook.Path & "\Alpha.ttf"
    If oFont.InstallFonts = False Then
        MsgBox "Could not install the font(s): " & oFont.FontName
    End If
EXIT_HERE:
    Exit Sub
ERR_HANDLER:
    GoTo EXIT_HERE
End Sub
