Attribute VB_Name = "modMsgHook"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const GWL_WNDPROC = -4

Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317

Private mcolItems       As Collection

Public Sub HookWindow(ByRef pobjToolTip As ToolTip)
Dim lngOldWndProc       As Long
    pobjToolTip.OldWndProc = SetWindowLong(pobjToolTip.hWnd, GWL_WNDPROC, AddressOf WndProc)
    If mcolItems Is Nothing Then
        Set mcolItems = New Collection
    End If
    mcolItems.Add ObjPtr(pobjToolTip), CreateKey(pobjToolTip.hWnd)
End Sub

Public Sub UnhookWindow(ByRef pobjToolTip As ToolTip)
    mcolItems.Remove CreateKey(pobjToolTip.hWnd)
    Call SetWindowLong(pobjToolTip.hWnd, GWL_WNDPROC, pobjToolTip.OldWndProc)
    If mcolItems.Count = 0 Then
        Set mcolItems = Nothing
    End If
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim objTip          As ToolTip
'On Error GoTo ErrHandler
    Set objTip = PtrObj(mcolItems.Item(CreateKey(hWnd)))
    Select Case uMsg
        Case WM_PRINT
            WndProc = 1
        Case WM_PAINT
            DrawToolTip objTip
            WndProc = 0
        Case Else
            WndProc = CallWindowProc(objTip.OldWndProc, hWnd, uMsg, wParam, lParam)
    End Select
    Set objTip = Nothing
'    Exit Function
'ErrHandler:
'    If Not (objTip Is Nothing) Then
'        Set objTip = Nothing
'    End If
'    Debug.Print "ERROR: " & Err.Description
End Function

Private Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject   As Object
    CopyMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    CopyMemory objObject, 0&, 4&
End Function

Private Function CreateKey(ByVal plnghWnd As Long) As String
    CreateKey = CStr(plnghWnd) & "K"
End Function
