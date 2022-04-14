VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolTipStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrWavFile         As String

Private mobjTextStyle       As TextStyle
Private mobjBalloonStyle    As BalloonStyle

Private Sub Class_Initialize()
    Set mobjTextStyle = New TextStyle
    Set mobjBalloonStyle = New BalloonStyle
End Sub

Private Sub Class_Terminate()
    Set mobjTextStyle = Nothing
    Set mobjBalloonStyle = Nothing
End Sub

Public Property Let WavFile(ByVal Value As String)
    mstrWavFile = Value
End Property

Public Property Get WavFile() As String
    WavFile = mstrWavFile
End Property

Public Property Get TextStyle() As TextStyle
    Set TextStyle = mobjTextStyle
End Property

Public Property Get BalloonStyle() As BalloonStyle
    Set BalloonStyle = mobjBalloonStyle
End Property


