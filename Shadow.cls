VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngColor       As Long
Private mlngOffsetX     As Long
Private mlngOffsetY     As Long
Private mblnVisible     As Boolean

Private Sub Class_Initialize()
    SetupDefaultStyle
End Sub

Private Sub Class_Terminate()
    '
End Sub

Private Sub SetupDefaultStyle()
    mlngColor = GetDefaultShadowColor
    mlngOffsetX = 2
    mlngOffsetY = 2
    mblnVisible = False
End Sub

Public Property Let Color(ByVal Value As Long)
    mlngColor = Value
End Property

Public Property Get Color() As Long
    Color = mlngColor
End Property

Public Property Let OffsetX(ByVal Value As Long)
    mlngOffsetX = Value
End Property

Public Property Get OffsetX() As Long
    OffsetX = mlngOffsetX
End Property

Public Property Let OffsetY(ByVal Value As Long)
    mlngOffsetY = Value
End Property

Public Property Get OffsetY() As Long
    OffsetY = mlngOffsetY
End Property

Public Property Let Visible(ByVal Value As Boolean)
    mblnVisible = Value
End Property

Public Property Get Visible() As Boolean
    Visible = mblnVisible
End Property
