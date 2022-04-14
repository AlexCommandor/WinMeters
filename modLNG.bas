Attribute VB_Name = "modLNG"
Option Explicit

Public Const LNG_MAX_ENTRIES = 28
Public strControls() As String, strCaptions() As String

Public Sub InitLanguage(Optional ByVal strLanguage As String = "ENU")
    Dim objForm As Form, objControl As Control, objControls As Controls, i As Long

    ReDim strControls(1 To LNG_MAX_ENTRIES):             ReDim strCaptions(1 To LNG_MAX_ENTRIES)
    
    strControls(1) = "wmSettings:frameCPU":              strCaptions(1) = "CPU indicator parameters"
    strControls(2) = "wmSettings:checkCPU":              strCaptions(2) = "Show indicator"
    strControls(3) = "wmSettings:checkOneTotalCPU":      strCaptions(3) = ""
    strControls(4) = "wmSettings:lblRefreshInterval":    strCaptions(4) = ""
    strControls(5) = "wmSettings:checkDigits":           strCaptions(5) = ""
    strControls(6) = "wmSettings:cmdSelectCPUFont":      strCaptions(6) = ""
    strControls(7) = "wmSettings:chkSolidColor":         strCaptions(7) = ""
    strControls(8) = "wmSettings:lblIndicatorColors":    strCaptions(8) = ""
    strControls(9) = "wmSettings:cmdResetColors":        strCaptions(9) = ""
    strControls(10) = "wmSettings:lblMainUsage":         strCaptions(10) = ""
    strControls(11) = "wmSettings:lblKernelUsage":       strCaptions(11) = ""
    strControls(12) = "wmSettings:frameMEM":             strCaptions(12) = ""
    strControls(13) = "wmSettings:checkMEM":             strCaptions(13) = strCaptions(2)
    strControls(14) = "wmSettings:checkAntiAliasedMem":  strCaptions(14) = ""
    strControls(15) = "wmSettings:chkShowKernelMem":     strCaptions(15) = ""
    strControls(16) = "wmSettings:frameHDD":             strCaptions(16) = ""
    strControls(17) = "wmSettings:checkHDD":             strCaptions(17) = strCaptions(2)
    strControls(18) = "wmSettings:checkExtendedHDDInfo": strCaptions(18) = ""
    strControls(19) = "wmSettings:frameNET":             strCaptions(19) = ""
    strControls(20) = "wmSettings:checkNET":             strCaptions(20) = strCaptions(2)
    strControls(21) = "wmSettings:lblNetInterface":      strCaptions(21) = ""
    strControls(22) = "wmSettings:checkShowTooltips":    strCaptions(22) = ""
    strControls(23) = "wmSettings:checkShowSplash":      strCaptions(23) = ""
    strControls(24) = "wmSettings:checkAutostart":       strCaptions(24) = ""
    strControls(25) = "wmSettings:cmdCloseMe":           strCaptions(25) = ""
    strControls(26) = "wmSettings:cmdExit":              strCaptions(26) = ""
    strControls(27) = "wmSettings:frameAbout":           strCaptions(27) = ""
    strControls(28) = "wmSettings:txtAbout":             strCaptions(28) = ""

    For i = 1 To Forms.Count
        MsgBox Forms(i).Caption
    Next i
End Sub
