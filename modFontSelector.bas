Attribute VB_Name = "modFontSelector"
Option Explicit

'**************************************
' Name: Font Selection Via API CAll
' Description:To Select the font via API Call, this decreases the burden of OCX in the Visual Basic
' By: Sanjay Gupta
'
' Inputs:In the input various parameter have to be passed so that the values of selection are passed into this variables
'
' Returns:input parameters arec changed to set as per the user selection
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=24277&lngWId=1'for details.'**************************************

Private Declare Function ChooseFont Lib "comdlg32.dll" _
        Alias "ChooseFontA" (lpChooseFont As udtCHOOSEFONT) As Long
Public fNm As String
Public fBld As Boolean
Public fSz As Single
Public fItl As Boolean
Public fUnl As Boolean
Public Fclr As Long

Type udtCHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type udtLogFont
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharset As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 31
End Type

Private Const ITALIC_FONTTYPE = &H200
Private Const BOLD_FONTTYPE = &H100
Private Const REGULAR_FONTTYPE = &H400
Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_ROMAN = 16

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

'Private Const CF_ANSIONLY = &H400&
'Private Const CF_APPLY = &H200&
Private Const CF_SCREENFONTS = &H1&
Private Const CF_PRINTERFONTS = &H2&
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&

Public Function ShowFont(Optional FontName As String = "Times New Roman", Optional Size As Integer, _
            Optional Bold As Boolean, Optional Italic As Boolean, Optional hwnd As Long) As String
    Dim cf As udtCHOOSEFONT, lfont As udtLogFont, hMem As Long, pMem As Long
    Dim retval As Long 'fontName As String
    With lfont
        .lfHeight = 0 ' determine default height
        .lfWidth = 0 ' determine default width
        .lfEscapement = 0 ' angle between baseline and escapement vector
        .lfOrientation = 0 ' angle between baseline and orientation vector
        .lfWeight = FW_NORMAL ' normal weight i.e. not bold
        .lfCharset = DEFAULT_CHARSET ' use default character set
        .lfOutPrecision = OUT_DEFAULT_PRECIS ' default precision mapping
        .lfClipPrecision = CLIP_DEFAULT_PRECIS ' default clipping precision
        .lfQuality = DEFAULT_QUALITY ' default quality setting
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN ' default pitch, proportional with serifs
        .lfFaceName = FontName & vbNullChar ' string must be null-terminated
        '.lfItalic = Italic
    End With
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem) ' lock and get pointer
    CopyMemory ByVal pMem, lfont, Len(lfont) ' copy structure's contents into block
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    With cf
        .lStructSize = Len(cf) ' size of structure
        If IsNumeric(hwnd) Then .hwndOwner = hwnd
        'cf.hwndOwner = Form1.hwnd ' window Form1 is opening this dialog box
        .hDC = Printer.hDC ' device context of default printer (using VB's mechanism)
        .lpLogFont = pMem ' pointer to LOGFONT memory block buffer
        If IsNumeric(Size) Then .iPointSize = Size * 10 Else .iPointSize = 120 ' 12 point font (in units of 1/10 point)
        .flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .rgbColors = RGB(0, 0, 0) ' black
        .nFontType = REGULAR_FONTTYPE ' regular font type i.e. not bold or anything
        If Italic Then .nFontType = .nFontType + ITALIC_FONTTYPE
        If Bold Then .nFontType = .nFontType + BOLD_FONTTYPE
        .nSizeMin = 6 ' minimum point size
        .nSizeMax = 20 ' maximum point size
    End With
    ' Now, call the function. If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    retval = ChooseFont(cf) ' open the dialog box
    If retval <> 0 Then ' success
        CopyMemory lfont, ByVal pMem, Len(lfont) ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        ShowFont = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        'Debug.Print ' end the line
    End If

    ' Deallocate the memory block we created earlier. Note that this must
    ' be done whether the function succeeded or not.
    retval = GlobalUnlock(hMem) ' destroy pointer, unlock block
    retval = GlobalFree(hMem) ' free the allocated memory
End Function

Public Function GetFont(Optional FontName As String, Optional Size As Integer, _
            Optional Bold As Boolean, Optional Italic As Boolean, _
            Optional UnderLine As Boolean, Optional Strikeout As Boolean, _
            Optional color As Long, Optional hwnd) As Long
    Dim rc As Long
    Dim pChooseFont As udtCHOOSEFONT
    Dim pLogFont As udtLogFont, hMem As Long, pMem As Long
    'Initialize the buffer
    With pLogFont
        .lfFaceName = FontName & Chr$(0)
        .lfItalic = Italic
        .lfUnderline = UnderLine
        .lfStrikeOut = Strikeout
        
    End With
    
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(pLogFont))
    pMem = GlobalLock(hMem) ' lock and get pointer
    CopyMemory ByVal pMem, pLogFont, Len(pLogFont) ' copy structure's contents into block
    
    'Initialize the structure
    With pChooseFont
        .hInstance = App.hInstance
        If IsNumeric(hwnd) Then .hwndOwner = hwnd
        '.flags = &H100 Or &H3&
        .flags = CF_BOTH Or CF_NOSCRIPTSEL Or CF_INITTOLOGFONTSTRUCT Or CF_FORCEFONTEXIST Or CF_LIMITSIZE
        'If IsNumeric(Size) Then .iPointSize = -(Size * 10)
        If IsNumeric(Size) Then .iPointSize = Size * 10
        'If Bold Then .nFontType = .nFontType + ITALIC_FONTTYPE
        .nFontType = REGULAR_FONTTYPE
        If Italic Then .nFontType = .nFontType + ITALIC_FONTTYPE
        If Bold Then .nFontType = .nFontType + BOLD_FONTTYPE
        'If IsNumeric(color) Then .rgbColors = color
        .lStructSize = Len(pChooseFont)
        '.lpLogFont = VarPtr(pLogFont)
        .lpLogFont = pMem
        .nSizeMin = 6
        .nSizeMax = 20
    End With
    
    
    'Call the API
    rc = ChooseFont(pChooseFont)
    If rc Then
        'Success!
        CopyMemory pLogFont, ByVal pMem, Len(pLogFont)
        'fontName = StrConv(pLogFont.lfFaceName, vbUnicode)
        'fontName = Left$(fontName, InStr(fontName, vbNullChar) - 1)
        FontName = Left$(pLogFont.lfFaceName, InStr(pLogFont.lfFaceName, vbNullChar) - 1)
        fNm = FontName
        'Return it's properties
        With pChooseFont
            Size = .iPointSize / 10
            Bold = (.nFontType And BOLD_FONTTYPE)
            Italic = (.nFontType And ITALIC_FONTTYPE)
            'UnderLine = (pLogFont.lfUnderline)
            'Strikeout = (pLogFont.lfStrikeOut)
            'color = (.rgbColors)
            fSz = Size
            fBld = Bold
            fItl = Italic
            'fUnl = UnderLine
            'Fclr = color
        End With
        'Return the Font Name
        GetFont = rc
    Else
        'The User clicked cancel
        GetFont = 0
    End If
    ' Deallocate the memory block we created earlier. Note that this must
    ' be done whether the function succeeded or not.
    rc = GlobalUnlock(hMem) ' destroy pointer, unlock block
    rc = GlobalFree(hMem) ' free the allocated memory
    
End Function

