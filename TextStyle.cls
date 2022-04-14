VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngColor           As Long
Private mstrFont            As String
Private mlngFontSize        As Long
Private mblnBold            As Boolean
Private mblnItalic          As Boolean
Private mblnUnderline       As Boolean

Private mlngLeftMargin      As Long
Private mlngRightMargin     As Long
Private mlngTopMargin       As Long
Private mlngBottomMargin    As Long
Private mobjShadow          As Shadow

Private Sub Class_Initialize()
    Set mobjShadow = New Shadow
    SetupDefaultStyle
End Sub

Private Sub Class_Terminate()
    Set mobjShadow = Nothing
End Sub

Private Sub SetupDefaultStyle()
    mlngColor = GetDefaultTextColor
    mstrFont = "Tahoma"
    mlngFontSize = 13
    mlngLeftMargin = 5
    mlngRightMargin = 5
    mlngTopMargin = 5
    mlngBottomMargin = 5
End Sub

Public Property Let Color(ByVal Value As Long)
    mlngColor = Value
End Property

Public Property Get Color() As Long
    Color = mlngColor
End Property

Public Property Let Font(ByVal Value As String)
    mstrFont = Value
End Property

Public Property Get Font() As String
    Font = mstrFont
End Property

Public Property Let FontSize(ByVal Value As Long)
    mlngFontSize = Value
End Property

Public Property Get FontSize() As Long
    FontSize = mlngFontSize
End Property

Public Property Let Bold(ByVal Value As Boolean)
    mblnBold = Value
End Property

Public Property Get Bold() As Boolean
    Bold = mblnBold
End Property

Public Property Let Italic(ByVal Value As Boolean)
    mblnItalic = Value
End Property

Public Property Get Italic() As Boolean
    Italic = mblnItalic
End Property

Public Property Let Underline(ByVal Value As Boolean)
    mblnUnderline = Value
End Property

Public Property Get Underline() As Boolean
    Underline = mblnUnderline
End Property

Public Property Let LeftMargin(ByVal Value As Long)
    mlngLeftMargin = Value
End Property

Public Property Get LeftMargin() As Long
    LeftMargin = mlngLeftMargin
End Property

Public Property Let RightMargin(ByVal Value As Long)
    mlngRightMargin = Value
End Property

Public Property Get RightMargin() As Long
    RightMargin = mlngRightMargin
End Property

Public Property Let TopMargin(ByVal Value As Long)
    mlngTopMargin = Value
End Property

Public Property Get TopMargin() As Long
    TopMargin = mlngTopMargin
End Property

Public Property Let BottomMargin(ByVal Value As Long)
    mlngBottomMargin = Value
End Property

Public Property Get BottomMargin() As Long
    BottomMargin = mlngBottomMargin
End Property

Public Property Get Shadow() As Shadow
    Set Shadow = mobjShadow
End Property
