VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNA         As Long = 8
Private Const TRANSPARENT       As Long = 1

Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400

Private Const BS_SOLID = 0
Private Const HS_CROSS = 4

Private Const LF_FACESIZE As Long = 32

Private Const FW_NORMAL = 400
Private Const FW_HEAVY = 900
Private Const FW_SEMIBOLD = 600
Private Const FW_BLACK = FW_HEAVY
Private Const FW_BOLD = 700
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_DONTCARE = 0
Private Const FW_EXTRABOLD = 800
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_MEDIUM = 500
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_THIN = 100
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_ULTRALIGHT = FW_EXTRALIGHT


Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private mudtPS      As PAINTSTRUCT
Private mudtRC      As RECT
Private mlnghWnd    As Long

Public Property Get hDC() As Long
    hDC = mudtPS.hDC
End Property

Public Property Let hWnd(ByVal Value As Long)
    mlnghWnd = Value
    GetWindowRect mlnghWnd, mudtRC
End Property

Public Property Get hWnd() As Long
    hWnd = mlnghWnd
End Property

Public Property Get Left() As Long
    Left = mudtRC.Left
End Property

Public Property Get Top() As Long
    Top = mudtRC.Top
End Property

Public Property Get Width() As Long
    Width = mudtRC.Right - mudtRC.Left
End Property

Public Property Get Height() As Long
    Height = mudtRC.Bottom - mudtRC.Top
End Property

Public Sub StartPainting()
    BeginPaint mlnghWnd, mudtPS
    SetBkMode mudtPS.hDC, TRANSPARENT
End Sub

Public Sub Finish()
    EndPaint mlnghWnd, mudtPS
    DeleteDC mudtPS.hDC
End Sub

Public Sub GetTextSize(ByVal pstrText As String, ByVal pstrFont As String, ByVal plngSize As Long, ByVal pblnBold As Boolean, ByVal pblnItalic As Boolean, ByVal pblnUnderline As Boolean, ByRef plngWidth As Long, ByRef plngHeight As Long)
Dim udtRC       As RECT
Dim lnghDC      As Long
Dim lngFont     As Long
Dim lngOldFont  As Long
    lnghDC = CreateCompatibleDC(mudtPS.hDC)
    lngFont = CreateFont(pstrFont, plngSize, pblnBold, pblnItalic, pblnUnderline)
    lngOldFont = SelectObject(lnghDC, lngFont)
    DrawText lnghDC, pstrText, Len(pstrText), udtRC, DT_VCENTER + DT_NOCLIP + DT_CALCRECT
    plngWidth = udtRC.Right - udtRC.Left
    plngHeight = udtRC.Bottom - udtRC.Top
    SelectObject lnghDC, lngOldFont
    DeleteObject lngFont
    DeleteDC lnghDC
End Sub

Public Sub Move(ByVal plngLeft As Long, ByVal plngTop As Long, ByVal plngWidth As Long, ByVal plngHeight As Long)
    With mudtRC
        .Left = plngLeft
        .Top = plngTop
        .Right = plngLeft + plngWidth
        .Bottom = plngTop + Height
    End With
    MoveWindow mlnghWnd, plngLeft, plngTop, plngWidth, plngHeight, True
End Sub

Public Sub SetRegionBackgroundColor(ByVal plngRegion As Long, ByVal plngColor As Long)
Dim lngBrush    As Long
Dim typLB       As LOGBRUSH
    With typLB
        .lbColor = plngColor
        .lbStyle = BS_SOLID
        .lbHatch = HS_CROSS
    End With
    lngBrush = CreateBrushIndirect(typLB)
    FillRgn mudtPS.hDC, plngRegion, lngBrush
    DeleteObject lngBrush
End Sub

Public Sub SetRegionBackgroundImage(ByVal plngRegion As Long, ByVal plngPicture As Long)
Dim lngBrush    As Long
    lngBrush = CreatePatternBrush(plngPicture)
    FillRgn mudtPS.hDC, plngRegion, lngBrush
    DeleteObject lngBrush
End Sub

Public Sub FrameRegion(ByVal plngRegion As Long, ByVal plngColor As Long)
    FrameRgn mudtPS.hDC, plngRegion, plngColor, 1, 1
End Sub

Public Sub DrawCustomText(ByVal pstrText As String, ByVal plngColor As Long, ByVal pstrFont As String, ByVal plngSize As Long, ByVal pblnBold As Boolean, ByVal pblnItalic As Boolean, ByVal pblnUnderline As Boolean, ByVal plngLeft As Long, ByVal plngTop As Long)
Dim udtRC       As RECT
Dim lngFont     As Long
Dim lngOldFont  As Long
    With udtRC
        .Left = plngLeft
        .Top = plngTop
        .Right = plngLeft
        .Bottom = plngTop
    End With
    lngFont = CreateFont(pstrFont, plngSize, pblnBold, pblnItalic, pblnUnderline)
    lngOldFont = SelectObject(mudtPS.hDC, lngFont)
    SetTextColor mudtPS.hDC, plngColor
    DrawText mudtPS.hDC, pstrText, Len(pstrText), udtRC, DT_VCENTER + DT_NOCLIP
    SelectObject mudtPS.hDC, lngOldFont
    DeleteObject lngFont
End Sub

Public Function CreateRegion(ByVal plngLeft As Long, ByVal plngTop As Long, ByVal plngWidth As Long, ByVal plngHeight As Long, Optional ByVal plngCurveSize As Long) As Long
Dim lngRegion   As Long
    If plngCurveSize > 0 Then
        lngRegion = CreateRoundRectRgn(plngLeft, plngTop, plngLeft + plngWidth, plngTop + plngHeight, plngCurveSize, plngCurveSize)
    Else
        lngRegion = CreateRectRgn(plngLeft, plngTop, plngLeft + plngWidth, plngTop + plngHeight)
    End If
    CreateRegion = lngRegion
End Function

Private Function CreateFont(ByVal pstrName As String, ByVal plngSize As Long, ByVal pblnBold As Boolean, ByVal pblnItalic As Boolean, ByVal pblnUnderline As Boolean) As Long
Dim lngIndex    As Long
Dim udtLF       As LOGFONT
    With udtLF
        .lfHeight = plngSize
        If pblnBold Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfItalic = pblnItalic
        .lfUnderline = pblnUnderline
        For lngIndex = 1 To Len(pstrName)
            .lfFaceName(lngIndex) = Asc(Mid(pstrName, lngIndex, 1))
        Next lngIndex
    End With
    CreateFont = CreateFontIndirect(udtLF)
End Function
