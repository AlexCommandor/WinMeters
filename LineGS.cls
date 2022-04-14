VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SmoothLineDIB"
Option Explicit

'Original TMT Pascal/Asm code by Jonas Widarsson
'
'Implemented in Vb6 by Dana Seaman
'Send comments/bug reports to dseaman@ieg.com.br
'
'REVISION HISTORY
'26-Jan-2002 Created LineGP  Method
'28-Jan-2002 Created LineDIB Method
'........... Created Class
'........... Optimized code
'05-Apr-2002 First Release to PSC *(Deleted by Hacker)
'09-Apr-2002 Improved speed, more comments and error handling
'11-Apr-2002 Added 3D clock hands
'12-Apr-2002 Improved clock timer event handler
'14-Apr-2002 Added Circle/Ellipse draw
'16-Apr-2002 Circle/Ellipse draw by Quadrants
'19-Apr-2002 Changed to RGBQuad (easier to understand DIBits)
'22-Apr-2002 Added circle thick/thin
'26-Apr-2002 Pass hDC together with LineGP/CircleGp
'........... Simplified blending code
'........... Common SetRGBComponents Sub
'28-Apr-2002 Added Arc drawing
'........... Several tweaks/speedups
'30-May-2002 Added Rounded Rectangle
'01-Jun-2002 Bevel/3D Rounded Rectangle
Public Enum cThickness
   Thin
   Thick
End Enum

Private Type RGBQUAD
   Blue                 As Byte
   Green                As Byte
   Red                  As Byte
   Reserved             As Byte
End Type

Private Type BITMAPINFOHEADER
   biSize               As Long
   biWidth              As Long
   biHeight             As Long
   biPlanes             As Integer
   biBitCount           As Integer
   biCompression        As Long
   biSizeImage          As Long
   biXPelsPerMeter      As Long
   biYPelsPerMeter      As Long
   biClrUsed            As Long
   biClrImportant       As Long
End Type

Private Type BITMAPINFO
   bmiHeader            As BITMAPINFOHEADER
End Type

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const Pi        As Single = 3.141592
Private Const HalfPi    As Single = Pi / 2
Private Const cThin     As Single = Pi * 0.34
Private Const cThick    As Single = Pi * 0.17
Private Const Rads      As Single = Pi / 180
Private Const PS_SOLID  As Long = 0
Private Const ICRESVER As Long = &H30000
Private Const GCL_HICON = (-14)

Private Binfo           As BITMAPINFO
Private buf()           As RGBQUAD
Private InDIBits        As Boolean
Private Red             As Long
Private Green           As Long
Private Blue            As Long
Private m_Color         As Long
Private m_hDC           As Long
Private m_W1            As Long
Private m_H1            As Long
Private m_Handle        As Long

Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Function TranslateColour(ByVal clr As OLE_COLOR, _
   Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColour) Then
      TranslateColour = vbBlack 'CLR_INVALID
   End If
End Function

Public Sub DIB(ByVal hDC As Long, ByVal handle As Long, ByVal W1 As Long, ByVal H1 As Long)
   m_hDC = hDC
   m_Handle = handle
   m_W1 = W1
   m_H1 = H1
   Pic2Array
End Sub

Private Sub Pic2Array()
   ReDim buf(0 To (m_W1 - 1), m_H1 - 1) As RGBQUAD
   With Binfo.bmiHeader
      .biSize = 40
      .biWidth = m_W1
      .biHeight = -m_H1
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = 0
      .biClrUsed = 0
      .biClrImportant = 0
      .biSizeImage = m_W1 * m_H1
   End With
   'Copy hDC to Array
   GetDIBits m_hDC, m_Handle, 0, m_H1, buf(0, 0), Binfo, DIB_RGB_COLORS
   'Set local flag
   InDIBits = True
End Sub

Public Sub CreateEmptyDIB(ByVal dibWidth As Long, ByVal dibHeight As Long)
   m_W1 = dibWidth
   m_H1 = dibHeight
   Erase buf()
   ReDim buf(0 To (m_W1 - 1), m_H1 - 1) As RGBQUAD
   With Binfo.bmiHeader
      .biSize = 40
      .biWidth = m_W1
      .biHeight = m_H1 * 2
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = 0
      .biClrUsed = 0
      .biClrImportant = 0
      .biSizeImage = (.biWidth + 1) * .biHeight * 2 'm_W1 * m_H1
      '.biSizeImage = .biWidth * .biHeight * 2  'm_W1 * m_H1
      '.biSizeImage = m_W1 * m_H1
      '.biXPelsPerMeter = &HB12 ' 72 dpi
      '.biYPelsPerMeter = &HB12 ' 72 dpi
   End With
   'Set all pixels to WHITE with full transparency
   'ClearAllToTransparency
   'Set local flag
   InDIBits = True
End Sub

Public Function GetDIB(ByRef outDIB() As Byte) As Long
    Dim i As Long, j As Long, k As Long
    ReDim outDIB(0 To (UBound(buf, 1) + 1) * (UBound(buf, 2) + 1) * 4 - 1)
    k = 0
    For i = 0 To UBound(buf, 1)
        For j = 0 To UBound(buf, 2)
            outDIB(k) = buf(i, j).Blue
            outDIB(k + 1) = buf(i, j).Green
            outDIB(k + 2) = buf(i, j).Red
            outDIB(k + 3) = buf(i, j).Reserved
            k = k + 4
        Next j
    Next i
    'CopyMemory ByVal VarPtr(outDIB(0)), ByVal VarPtr(buf(0, 0)), UBound(outDIB) + 1
    GetDIB = UBound(buf, 1) + 1
End Function

Public Function GetIconArray(ByRef outArray() As Byte, Optional ByRef alsoGetIconWidth As Long) As Long
    Dim arrDIB() As Byte, lDIBSize As Long, lAllSize As Long
    lDIBSize = GetDIB(arrDIB)
    If Not IsMissing(alsoGetIconWidth) Then alsoGetIconWidth = lDIBSize
    lDIBSize = UBound(arrDIB) + 1
    lAllSize = lDIBSize + 40
    ReDim outArray(0 To lDIBSize + 39)
    CopyMemory outArray(0), Binfo.bmiHeader, LenB(Binfo.bmiHeader)
    CopyMemory outArray(40), arrDIB(0), lDIBSize - 1
    ReDim Preserve outArray(0 To Binfo.bmiHeader.biSizeImage - 1)
    lAllSize = Binfo.bmiHeader.biSizeImage
    GetIconArray = lAllSize
End Function

Public Function GetIconFromDIB() As Long
    Dim arrIcon() As Byte, lIconBufSize As Long, lIcon As Long, lWidth As Long
    lIconBufSize = GetIconArray(arrIcon, lWidth)
    'lIcon = CreateIconFromResource(arrIcon(0), lIconBufSize, True, ICRESVER)
    lIcon = CreateIconFromResourceEx(arrIcon(0), lIconBufSize, True, ICRESVER, lWidth, lWidth, LR_DEFAULTCOLOR)
    GetIconFromDIB = lIcon
End Function

Private Function GetPointSpacing(Radius As Long, Thickness As cThickness) As Single
   Dim sngLS            As Single

   If Thickness = Thick Then
      sngLS = cThick
   Else
      sngLS = cThin
   End If

   If Radius < 0 Then
      GetPointSpacing = -sngLS / Radius
   ElseIf Radius = 0 Then
      GetPointSpacing = sngLS
   Else
      GetPointSpacing = sngLS / Radius
   End If

End Function

Public Sub CircleDIB(ByVal x1 As Long, _
   ByVal y1 As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal color As OLE_COLOR, _
   Optional ByVal Thickness As cThickness = Thick)

   Dim Bbg              As Byte
   Dim Gbg              As Byte
   Dim Rbg              As Byte
   Dim Abg              As Byte
   Dim savAlpha(1 To 4) As Byte
   Dim Cl               As Long
   Dim Strength         As Long
   Dim StrengthI        As Long
   Dim Quadrant         As Long
   Dim Radius           As Long
   Dim RX1              As Long
   Dim RX2              As Long
   Dim RY1              As Long
   Dim RY2              As Long
   Dim savX(1 To 4)     As Long
   Dim savY(1 To 4)     As Long
   Dim X4               As Long
   Dim Y4               As Long
   Dim Ax               As Single
   Dim Ay               As Single
   Dim Bx               As Single
   Dim By               As Single
   Dim L1               As Single
   Dim l2               As Single
   Dim L3               As Single
   Dim L4               As Single
   Dim sngAngle         As Single
   Dim sngPointSpacing  As Single
   Dim X2               As Single
   Dim Xp5              As Single
   Dim Y2               As Single

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling CircleDIB."
      Exit Sub
   End If

   SetRGBComponents color

   Radius = RadiusX
   If RadiusY > RadiusX Then
      Radius = RadiusY
   End If

   sngPointSpacing = GetPointSpacing(Radius, Thickness)

   For sngAngle = 0 To HalfPi Step sngPointSpacing
      X2 = RadiusX * Cos(sngAngle)
      Y2 = RadiusY * Sin(sngAngle)
      'Prevents error when vb rounds .5 down
      If X2 = Int(X2) Then X2 = X2 + 0.001
      If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
      For Quadrant = 0 To 3
         Select Case Quadrant
            Case 0 '0-90°
               Ax = X2 + x1 - 0.5
               Ay = -Y2 + y1 - 0.5
            Case 1 '90-180°
               Ax = X2 + x1 - 0.5
               Ay = Y2 + y1 - 0.5
            Case 2 '180-270°
               Ax = -X2 + x1 - 0.5
               Ay = Y2 + y1 - 0.5
            Case 3 '270-360°
               Ax = -X2 + x1 - 0.5
               Ay = -Y2 + y1 - 0.5
         End Select

         Bx = Ax + 1
         By = Ay + 1
         RX1 = Ax
         RX2 = RX1 + 1
         Xp5 = RX1 + 0.5
         RY1 = Ay
         RY2 = By
         L1 = RY1 + 0.5 - Ay
         l2 = 256 * (Xp5 - Ax) - Xp5 + Ax
         L3 = 255 - l2
         L4 = By - RY2 + 0.5
         savX(1) = RX1
         savY(1) = RY1
         savX(2) = RX2
         savY(2) = RY1
         savY(3) = RY2
         savX(3) = RX1
         savY(4) = RY2
         savX(4) = RX2
         savAlpha(1) = L1 * l2
         savAlpha(2) = L1 * L3
         savAlpha(3) = L4 * l2
         savAlpha(4) = L4 * L3

         For Cl = 1 To 4
            Strength = savAlpha(Cl)
            X4 = savX(Cl)
            Y4 = savY(Cl)
            'Check if in bounds
            If X4 >= 0 And X4 < m_W1 And Y4 >= 0 And Y4 <= m_H1 Then
               If Strength > 252 Then  '> 99%
                  'No blending
                  buf(X4, Y4).Blue = Blue
                  buf(X4, Y4).Green = Green
                  buf(X4, Y4).Red = Red
                  buf(X4, Y4).Reserved = &HFF
               Else
                If Not (buf(X4, Y4).Blue = &H0 And buf(X4, Y4).Green = &H0 And buf(X4, Y4).Red = &H0 And buf(X4, Y4).Reserved = &H0) Then
                  Bbg = buf(X4, Y4).Blue
                  Gbg = buf(X4, Y4).Green
                  Rbg = buf(X4, Y4).Red
                  Abg = buf(X4, Y4).Reserved
                  'Blend % of bgColor with % of Color
                  StrengthI = 255 - Strength
                  buf(X4, Y4).Red = (StrengthI * Rbg + Strength * Red) \ 256
                  buf(X4, Y4).Green = (StrengthI * Gbg + Strength * Green) \ 256
                  buf(X4, Y4).Blue = (StrengthI * Bbg + Strength * Blue) \ 256
                  buf(X4, Y4).Reserved = (StrengthI * Abg + Strength * &HFF) \ 256
                Else
                  buf(X4, Y4).Blue = Blue
                  buf(X4, Y4).Green = Green
                  buf(X4, Y4).Red = Red
                  buf(X4, Y4).Reserved = Strength
                End If
               End If
            End If
         Next
      Next
   Next
End Sub

Public Sub ArcDIB(ByVal x1 As Long, _
   ByVal y1 As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal startAngle As Single, _
   ByVal StopAngle As Single, _
   ByVal color As OLE_COLOR, _
   Optional ByVal Thickness As cThickness = Thick)

   Dim Bbg              As Byte
   Dim Gbg              As Byte
   Dim Rbg              As Byte
   Dim savAlpha(1 To 4) As Byte
   Dim Cl               As Long
   Dim Strength         As Long
   Dim StrengthI        As Long
   Dim Radius           As Long
   Dim RX1              As Long
   Dim RX2              As Long
   Dim RY1              As Long
   Dim RY2              As Long
   Dim savX(1 To 4)     As Long
   Dim savY(1 To 4)     As Long
   Dim X4               As Long
   Dim Y4               As Long
   Dim Ax               As Single
   Dim Ay               As Single
   Dim Bx               As Single
   Dim By               As Single
   Dim L1               As Single
   Dim l2               As Single
   Dim L3               As Single
   Dim L4               As Single
   Dim sngAngle         As Single
   Dim sngPointSpacing  As Single
   Dim X2               As Single
   Dim Xp5              As Single
   Dim Y2               As Single

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling CircleDIB."
      Exit Sub
   End If

   SetRGBComponents color

   Radius = RadiusX
   If RadiusY > RadiusX Then
      Radius = RadiusY
   End If

   sngPointSpacing = GetPointSpacing(Radius, Thickness)

   If startAngle > StopAngle Then
      StopAngle = StopAngle + 360
   End If
   'Convert to Radians
   startAngle = startAngle * Rads
   StopAngle = StopAngle * Rads

   For sngAngle = startAngle To StopAngle Step sngPointSpacing
      X2 = RadiusX * Cos(sngAngle - HalfPi)
      Y2 = RadiusY * Sin(sngAngle + HalfPi)
      'Prevents error when vb rounds .5 down
      If X2 = Int(X2) Then X2 = X2 + 0.001
      If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
      Ax = X2 + x1 - 0.5
      Ay = Y2 + y1 - 0.5
      Bx = Ax + 1
      By = Ay + 1
      RX1 = Ax
      RX2 = RX1 + 1
      Xp5 = RX1 + 0.5
      RY1 = Ay
      RY2 = By
      L1 = RY1 + 0.5 - Ay
      l2 = 256 * (Xp5 - Ax) - Xp5 + Ax
      L3 = 255 - l2
      L4 = By - RY2 + 0.5
      savX(1) = RX1
      savY(1) = RY1
      savX(2) = RX2
      savY(2) = RY1
      savY(3) = RY2
      savX(3) = RX1
      savY(4) = RY2
      savX(4) = RX2
      savAlpha(1) = L1 * l2
      savAlpha(2) = L1 * L3
      savAlpha(3) = L4 * l2
      savAlpha(4) = L4 * L3

      For Cl = 1 To 4
         Strength = savAlpha(Cl)
         X4 = savX(Cl)
         Y4 = savY(Cl)
         'Check if in bounds
         If X4 >= 0 And X4 < m_W1 And Y4 >= 0 And Y4 <= m_H1 Then
            If Strength > 252 Then  '> 99%
               'No blending
               buf(X4, Y4).Blue = Blue
               buf(X4, Y4).Green = Green
               buf(X4, Y4).Red = Red
            Else
               Bbg = buf(X4, Y4).Blue
               Gbg = buf(X4, Y4).Green
               Rbg = buf(X4, Y4).Red
               'Blend % of bgColor with % of Color
               StrengthI = 255 - Strength
               buf(X4, Y4).Red = (StrengthI * Rbg + Strength * Red) \ 256
               buf(X4, Y4).Green = (StrengthI * Gbg + Strength * Green) \ 256
               buf(X4, Y4).Blue = (StrengthI * Bbg + Strength * Blue) \ 256
            End If
         End If
      Next
   Next
End Sub

Public Sub SetRGBComponents(ByVal color As OLE_COLOR)

   color = TranslateColour(color)
   m_Color = color 'make available global
   If color Then
      Red = color And &HFF&
      Green = color \ 256 And &HFF
      Blue = color \ 65536
   Else 'Color is Black
      Red = 0
      Green = 0
      Blue = 0
   End If
End Sub

Public Sub LineDIB(ByVal x1 As Integer, _
   ByVal y1 As Integer, _
   ByVal X2 As Integer, _
   ByVal Y2 As Integer, _
   ByVal color As OLE_COLOR, Optional ByVal Alfa As Long = -1)

   'Blended lines, maximum blend at transition,
   'tapering off too minimum.
   '
   '            ----------
   'Min    Max/Max     Min
   '----------
   '

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling LineDIB."
      Exit Sub
   End If

   Dim XScope           As Long
   Dim YScope           As Long
   Dim XDir             As Long
   Dim YDir             As Long
   Dim LinearDeviance   As Long
   Dim Counter          As Long
   Dim AntiAliasStrength As Long
   Dim EndPointIntensity As Long
   Dim HalfIntensity As Long
   
   If Alfa = -1 Then
    HalfIntensity = 127
   Else
    HalfIntensity = Alfa \ 2
   End If

   XScope = X2 - x1
   YScope = Y2 - y1

   If XScope < 0 Then
      XScope = Abs(XScope)
      XDir = -1
   Else
      XDir = 1
   End If

   If YScope < 0 Then
      YScope = Abs(YScope)
      YDir = -1
   Else
      YDir = 1
   End If

   If XScope + YScope = 0 Then
      'Exit if line length is 0
      Exit Sub
   End If

   SetRGBComponents color

   If XScope > YScope Then
      'Output EndPoints outside of main loop.
      'If Alfa = -1 Then
      '  EndPointIntensity = (85 * YScope) \ XScope
      'Else
      '  EndPointIntensity = Alfa
      'End If
      'PutPixelDIB x1 - XDir, y1 - YDir, EndPointIntensity
      'PutPixelDIB x1 - XDir, y1, HalfIntensity
      'PutPixelDIB X2 + XDir, Y2 + YDir, EndPointIntensity
      'PutPixelDIB X2 + XDir, Y2, HalfIntensity
      '-----
        LinearDeviance = XScope \ 2
      For Counter = 0 To XScope
         'Main line, output full strength direct to DIB array.
         If Alfa = -1 Then
            PutPixelDIB x1, y1, 255
         Else
            PutPixelDIB x1, y1, Alfa
         End If
         'Output the blended lines for anti-alias effect.
         AntiAliasStrength = (LinearDeviance * 255) \ XScope
         PutPixelDIB x1, y1 - YDir, 255 - AntiAliasStrength
         PutPixelDIB x1, y1 + YDir, AntiAliasStrength
         LinearDeviance = (LinearDeviance + YScope)
         If LinearDeviance >= XScope Then
            LinearDeviance = LinearDeviance - XScope
            y1 = y1 + YDir
         End If
         x1 = x1 + XDir
      Next
   Else
      'Output EndPoints outside of main loop.
      'If Alfa = -1 Then
      '  EndPointIntensity = (85 * XScope) \ YScope
      'Else
      '  EndPointIntensity = Alfa
      'End If
      'PutPixelDIB x1 - XDir, y1 - YDir, EndPointIntensity
      'PutPixelDIB x1, y1 - YDir, HalfIntensity
      'PutPixelDIB X2 + XDir, Y2 + YDir, EndPointIntensity
      'PutPixelDIB X2, Y2 + YDir, HalfIntensity
      '-----
        LinearDeviance = YScope \ 2
      For Counter = 0 To YScope
         'Main line, output full strength direct to DIB array.
         If Alfa = -1 Then
            PutPixelDIB x1, y1, 255
         Else
            PutPixelDIB x1, y1, Alfa
         End If
         'Output the blended lines for anti-alias effect.
         AntiAliasStrength = (LinearDeviance * 255) \ YScope
         PutPixelDIB x1 - XDir, y1, 255 - AntiAliasStrength
         PutPixelDIB x1 + XDir, y1, AntiAliasStrength
         LinearDeviance = LinearDeviance + XScope
         If (LinearDeviance >= YScope) Then
            LinearDeviance = LinearDeviance - YScope
            x1 = x1 + XDir
         End If
         y1 = y1 + YDir
      Next
   End If

End Sub


Public Sub LineDIBwithoutAntialiasing(ByVal x1 As Integer, _
   ByVal y1 As Integer, _
   ByVal X2 As Integer, _
   ByVal Y2 As Integer, _
   ByVal color As OLE_COLOR, Optional ByVal Alfa As Long = 255)

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling LineDIB."
      Exit Sub
   End If

   Dim XScope           As Long
   Dim YScope           As Long
   Dim XDir             As Long
   Dim YDir             As Long
   Dim LinearDeviance   As Long
   Dim Counter          As Long
   Dim EndPointIntensity As Long
   Dim HalfIntensity As Long
   
   If Alfa = -1 Then
    HalfIntensity = 127
   Else
    HalfIntensity = Alfa \ 2
   End If

   XScope = X2 - x1
   YScope = Y2 - y1

   If XScope < 0 Then
      XScope = Abs(XScope)
      XDir = -1
   Else
      XDir = 1
   End If

   If YScope < 0 Then
      YScope = Abs(YScope)
      YDir = -1
   Else
      YDir = 1
   End If

   If XScope + YScope = 0 Then
      'Exit if line length is 0
      Exit Sub
   End If

   SetRGBComponents color

   If XScope > YScope Then
      ''Output EndPoints outside of main loop.
      'If Alfa = 255 Then
      '  EndPointIntensity = (85 * YScope) \ XScope
      'Else
      '  EndPointIntensity = Alfa
      'End If
      'PutPixelDIB x1 - XDir, y1 - YDir, EndPointIntensity
      'PutPixelDIB x1 - XDir, y1, HalfIntensity
      'PutPixelDIB X2 + XDir, Y2 + YDir, EndPointIntensity
      'PutPixelDIB X2 + XDir, Y2, HalfIntensity
      '-----
      LinearDeviance = XScope \ 2
      For Counter = 0 To XScope
         
         PutPixelDIB x1, y1, Alfa, False

         LinearDeviance = (LinearDeviance + YScope)
         If LinearDeviance >= XScope Then
            LinearDeviance = LinearDeviance - XScope
            y1 = y1 + YDir
         End If
         x1 = x1 + XDir
      Next
   Else
      ''Output EndPoints outside of main loop.
      'If Alfa = 255 Then
      '  EndPointIntensity = (85 * XScope) \ YScope
      'Else
      '  EndPointIntensity = Alfa
      'End If
      'PutPixelDIB x1 - XDir, y1 - YDir, EndPointIntensity
      'PutPixelDIB x1, y1 - YDir, HalfIntensity
      'PutPixelDIB X2 + XDir, Y2 + YDir, EndPointIntensity
      'PutPixelDIB X2, Y2 + YDir, HalfIntensity
      '-----
      LinearDeviance = YScope \ 2
      For Counter = 0 To YScope
         'Main line, output full strength direct to DIB array.
         PutPixelDIB x1, y1, Alfa, False
         LinearDeviance = LinearDeviance + XScope
         If (LinearDeviance >= YScope) Then
            LinearDeviance = LinearDeviance - YScope
            x1 = x1 + XDir
         End If
         y1 = y1 + YDir
      Next
   End If

End Sub

Public Sub Array2Pic()
   'If we have an array copy back to hDC
   If InDIBits Then
      SetDIBits m_hDC, m_Handle, 0, m_H1, buf(0, 0), Binfo, DIB_RGB_COLORS
      InDIBits = False
      'Erase buf '* Moved to Class_Terminate
   End If
End Sub

Private Sub PutPixelDIB(ByVal x As Long, _
   ByVal y As Long, _
   ByVal Strength As Long, Optional ByVal MakeBlend As Boolean = True)

   Dim Rbg              As Long
   Dim Gbg              As Long
   Dim Bbg              As Long
   Dim Abg              As Long
   Dim StrengthI        As Long

   'Check if in bounds
   If x < 0 Or x >= m_W1 Or y < 0 Or y > m_H1 Then
      Exit Sub
   End If
   If Strength > 252 Then '99%
      buf(x, y).Blue = Blue
      buf(x, y).Green = Green
      buf(x, y).Red = Red
      buf(x, y).Reserved = &HFF
   Else
     If MakeBlend Then
        If Not (buf(x, y).Blue = &H0 And buf(x, y).Green = &H0 And buf(x, y).Red = &H0 And buf(x, y).Reserved = &H0) Then
          '##### Get Background Pixel components
          Bbg = buf(x, y).Blue
          Gbg = buf(x, y).Green
          Rbg = buf(x, y).Red
          Abg = buf(x, y).Reserved
          '##### Blend % of bgColor with % of m_Color
          StrengthI = 255 - Strength
          buf(x, y).Red = (StrengthI * Rbg + Strength * Red) \ 256
          buf(x, y).Green = (StrengthI * Gbg + Strength * Green) \ 256
          buf(x, y).Blue = (StrengthI * Bbg + Strength * Blue) \ 256
          buf(x, y).Reserved = (StrengthI * Abg + Strength * &HFF) \ 256
        Else
          buf(x, y).Blue = Blue
          buf(x, y).Green = Green
          buf(x, y).Red = Red
          buf(x, y).Reserved = Strength
        End If
     Else
        buf(x, y).Blue = Blue
        buf(x, y).Green = Green
        buf(x, y).Red = Red
        buf(x, y).Reserved = Strength
     End If
   End If
End Sub

Public Sub ClearAllToTransparency()
    Erase buf
    ReDim buf(0 To (m_W1 - 1), m_H1 - 1) As RGBQUAD
    'Dim i As Long, j As Long
    'For i = LBound(buf, 1) To UBound(buf, 1)
    '    For j = LBound(buf, 2) To UBound(buf, 2)
    '        buf(i, j).Blue = &HFF
    '        buf(i, j).Green = &HFF
    '        buf(i, j).Red = &HFF
    '        buf(i, j).Reserved = &H0
    '    Next j
    'Next i
End Sub

Private Sub Class_Terminate()
   Erase buf()
End Sub

Public Sub PieDIB(ByVal centerX As Long, ByVal centerY As Long, _
            ByVal Radius As Long, ByVal startAngle As Long, ByVal endAngle As Long, _
            ByVal cColor As OLE_COLOR, Optional ByVal Alfa As Byte = 255)
    Dim i As Long, newX As Double, newY As Double, prevX As Double, prevY As Double
    For i = 90 + startAngle To 90 + endAngle ' Step 2
        newX = Round(Radius * Sin(i * Pi / 180) + Radius + 2)
        newY = Round(Radius * Cos(i * Pi / 180) + Radius + 2)
        If prevX <> newX Or prevY <> newY Then
            LineDIBwithoutAntialiasing centerX, centerY, newX, newY, cColor, Alfa
            prevX = newX: prevY = newY
        End If
    Next i
        newX = Round(Radius * Sin((90 + endAngle) * Pi / 180) + Radius + 2)
        newY = Round(Radius * Cos((90 + endAngle) * Pi / 180) + Radius + 2)
        LineDIB centerX, centerY, newX, newY, cColor, Alfa
End Sub


