VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalloonStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngColor           As Long
Private mlngImage           As Long
Private mlngCurveIndex      As Long
Private mobjShadow          As Shadow

Private Sub Class_Initialize()
    Set mobjShadow = New Shadow
    SetupDefaultStyle
End Sub

Private Sub Class_Terminate()
    Set mobjShadow = Nothing
End Sub

Private Sub SetupDefaultStyle()
    mlngColor = GetDefaultBalloonColor
    mlngCurveIndex = 7
End Sub

Public Property Let Color(ByVal Value As Long)
    mlngColor = Value
End Property

Public Property Get Color() As Long
    Color = mlngColor
End Property

Public Property Let Image(ByVal Value As Long)
    mlngImage = Value
End Property

Public Property Get Image() As Long
    Image = mlngImage
End Property

Public Property Let CurveIndex(ByVal Value As Long)
    mlngCurveIndex = Value
End Property

Public Property Get CurveIndex() As Long
    CurveIndex = mlngCurveIndex
End Property

Public Property Get Shadow() As Shadow
    Set Shadow = mobjShadow
End Property
