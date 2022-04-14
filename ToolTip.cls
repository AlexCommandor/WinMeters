VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrKey             As String
Private mlngFormhWnd        As Long
Private mlngControlhWnd     As Long

Private mlnghWnd            As Long
Private mlngOldWndProc      As Long

Private mstrText            As String
Private mobjStyle           As ToolTipStyle

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    DestroyToolTip Me
End Sub

Friend Sub Initialise(ByVal plngFormhWnd As Long, ByRef pctlControl As Control, ByVal pstrText As String)
    mstrKey = pctlControl.Name
    mlngFormhWnd = plngFormhWnd
    mlngControlhWnd = pctlControl.hWnd
    mstrText = pstrText
    CreateToolTip Me
End Sub

Friend Property Get Key() As String
    Key = mstrKey
End Property

Friend Property Get FormhWnd() As Long
    FormhWnd = mlngFormhWnd
End Property

Friend Property Get ControlhWnd() As Long
    ControlhWnd = mlngControlhWnd
End Property

Friend Property Let hWnd(ByVal Value As Long)
    mlnghWnd = Value
End Property

Friend Property Get hWnd() As Long
    hWnd = mlnghWnd
End Property

Friend Property Let OldWndProc(ByVal Value As Long)
    mlngOldWndProc = Value
End Property

Friend Property Get OldWndProc() As Long
    OldWndProc = mlngOldWndProc
End Property

Public Property Let Text(ByVal Value As String)
    If Not (Value = mstrText) Then
        DestroyToolTip Me
        mstrText = Value
        CreateToolTip Me
    End If
End Property

Public Property Get Text() As String
    Text = mstrText
End Property

Public Property Set Style(ByRef Value As ToolTipStyle)
    Set mobjStyle = Value
End Property

Public Property Get Style() As ToolTipStyle
    Set Style = mobjStyle
End Property
