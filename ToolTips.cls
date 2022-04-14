VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngFormhWnd    As Long

Private mcolItems       As Collection

Private Sub Class_Initialize()
    Set mcolItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolItems = Nothing
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Property Let FormhWnd(ByVal Value As Long)
    mlngFormhWnd = Value
End Property

Public Property Get FormhWnd() As Long
    FormhWnd = mlngFormhWnd
End Property

Public Property Get Item(ByRef Index As Variant) As ToolTip
Attribute Item.VB_UserMemId = 0
Dim varIndex    As Variant
    If TypeOf Index Is Control Then
        varIndex = Index.Name
    Else
        varIndex = Index
    End If
    Set Item = mcolItems.Item(varIndex)
End Property

Public Function Exists(ByRef Index As Variant) As Boolean
Dim objTip  As ToolTip
On Error Resume Next
    Set objTip = Item(Index)
On Error GoTo 0
    If Not (objTip Is Nothing) Then
        Exists = True
        Set objTip = Nothing
    End If
End Function

Public Property Get Count() As Long
    Count = mcolItems.Count
End Property

Public Function Add(ByRef Control As Object, ByVal Text As String, Optional ByRef Style As ToolTipStyle) As ToolTip
Dim objNew  As ToolTip
    Set objNew = New ToolTip
    objNew.Initialise mlngFormhWnd, Control, Text
    Set objNew.Style = Style
    AddToolTip objNew
    Set Add = objNew
    Set objNew = Nothing
End Function

Private Sub AddToolTip(ByRef pobjTip As ToolTip)
    mcolItems.Add pobjTip, pobjTip.Key
End Sub

Public Sub Remove(ByRef Index As Variant)
    mcolItems.Remove Item(Index).Key
End Sub
