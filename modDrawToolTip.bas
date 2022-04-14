Attribute VB_Name = "modDrawToolTip"
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type mtypSize
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const CURSOR_OFFSET_X As Long = 12
Private Const CURSOR_OFFSET_Y As Long = 20

Public Sub DrawToolTip(ByRef pobjTip As ToolTip)
Dim udtPT           As POINTAPI
Dim objDraw         As clsDraw
Dim lngRgnBalloon   As Long
Dim lngRgnShadow    As Long
Dim objStyle        As ToolTipStyle
Dim lngTxtLeft      As Long
Dim lngTxtTop       As Long
Dim lngTxtShdLeft   As Long
Dim lngTxtShdTop    As Long
Dim lngWidth        As Long
Dim lngHeight       As Long
Dim udtBalloon      As mtypSize
Dim udtShadow       As mtypSize
Dim lngOffset       As Long
    Set objStyle = pobjTip.Style
    If objStyle Is Nothing Then
        Set objStyle = New ToolTipStyle
    End If
    If Len(objStyle.WavFile) > 0 Then
        PlayWavFile objStyle.WavFile
    End If
    GetCursorPos udtPT
    Set objDraw = New clsDraw
    objDraw.hWnd = pobjTip.hWnd
    
    With objStyle.TextStyle
        objDraw.GetTextSize pobjTip.Text, .Font, .FontSize, .Bold, .Italic, .Underline, lngWidth, lngHeight
        lngWidth = lngWidth + .LeftMargin + .RightMargin
        lngHeight = lngHeight + .TopMargin + .BottomMargin
    End With
    If objStyle.BalloonStyle.Shadow.Visible Then
        lngWidth = lngWidth + Abs(objStyle.BalloonStyle.Shadow.OffsetX)
        lngHeight = lngHeight + Abs(objStyle.BalloonStyle.Shadow.OffsetY)
    End If
    If objStyle.TextStyle.Shadow.Visible Then
        lngWidth = lngWidth + Abs(objStyle.TextStyle.Shadow.OffsetX)
        lngHeight = lngHeight + Abs(objStyle.TextStyle.Shadow.OffsetY)
    End If

    'resize window
    objDraw.Move udtPT.x + CURSOR_OFFSET_X, udtPT.y + CURSOR_OFFSET_Y, lngWidth, lngHeight
    
    objDraw.StartPainting
    
    'calculate position to draw things
    With udtBalloon
        .Left = 0
        .Top = 0
        .Width = lngWidth
        .Height = lngHeight
    End With
    LSet udtShadow = udtBalloon
    If objStyle.BalloonStyle.Shadow.Visible Then
        lngOffset = objStyle.BalloonStyle.Shadow.OffsetX
        If lngOffset > 0 Then
            udtShadow.Left = lngOffset
            udtShadow.Width = udtShadow.Width - lngOffset
            udtBalloon.Width = udtBalloon.Width - lngOffset
        Else
            udtShadow.Width = udtShadow.Width + lngOffset
            udtBalloon.Left = Abs(lngOffset)
            udtBalloon.Width = udtBalloon.Width + lngOffset
        End If
        lngOffset = objStyle.BalloonStyle.Shadow.OffsetY
        If lngOffset > 0 Then
            udtShadow.Top = lngOffset
            udtShadow.Height = udtShadow.Height - lngOffset
            udtBalloon.Height = udtBalloon.Height - lngOffset
        Else
            udtShadow.Height = udtShadow.Height + lngOffset
            udtBalloon.Top = Abs(lngOffset)
            udtBalloon.Height = udtBalloon.Height + lngOffset
        End If
        With udtShadow
            lngRgnShadow = objDraw.CreateRegion(.Left, .Top, .Width, .Height, objStyle.BalloonStyle.CurveIndex)
        End With
        objDraw.SetRegionBackgroundColor lngRgnShadow, objStyle.BalloonStyle.Shadow.Color
    End If
    With udtBalloon
        lngRgnBalloon = objDraw.CreateRegion(.Left, .Top, .Width, .Height, objStyle.BalloonStyle.CurveIndex)
    End With
    
    If objStyle.BalloonStyle.Image <> 0 Then
        objDraw.SetRegionBackgroundImage lngRgnBalloon, objStyle.BalloonStyle.Image
    Else
        objDraw.SetRegionBackgroundColor lngRgnBalloon, objStyle.BalloonStyle.Color
    End If
    objDraw.FrameRegion lngRgnBalloon, GetDefaultBorderColor
    
    lngTxtLeft = udtBalloon.Left + objStyle.TextStyle.LeftMargin
    lngTxtTop = udtBalloon.Top + objStyle.TextStyle.TopMargin
    lngTxtShdLeft = lngTxtLeft
    lngTxtShdTop = lngTxtTop
    
    
    'draw text and text shadow
    With objStyle.TextStyle
        If .Shadow.Visible Then
            lngOffset = .Shadow.OffsetX
            If lngOffset > 0 Then
                lngTxtShdLeft = lngTxtShdLeft + lngOffset
            Else
                lngTxtLeft = lngTxtLeft + Abs(lngOffset)
            End If
            lngOffset = .Shadow.OffsetY
            If lngOffset > 0 Then
                lngTxtShdTop = lngTxtShdTop + lngOffset
            Else
                lngTxtTop = lngTxtTop + Abs(lngOffset)
            End If
            objDraw.DrawCustomText pobjTip.Text, .Shadow.Color, .Font, .FontSize, .Bold, .Italic, .Underline, lngTxtShdLeft, lngTxtShdTop
        End If
        objDraw.DrawCustomText pobjTip.Text, .Color, .Font, .FontSize, .Bold, .Italic, .Underline, lngTxtLeft, lngTxtTop
    End With
    
    'close off drawing and release resources
    objDraw.Finish
    Set objDraw = Nothing
End Sub




