Attribute VB_Name = "modToolTip"
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    R As RECT
    hinst As Long
    lpszText As String
End Type

Private Const WM_USER = &H400

Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOL = (WM_USER + 4)
Private Const TTM_DELTOOL = WM_USER + 5
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTN_FIRST = (-520)
Private Const TTN_SHOW = (TTN_FIRST - 1)

Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000
Private Const CW_USEDEFAULT = &H80000000

Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_3DLIGHT = 22
Private Const COLOR_BTNSHADOW As Long = 16



Private Const DKGRAY_BRUSH = 3

Public Sub CreateToolTip(ByRef pobjTip As ToolTip)
Dim lngStyle    As Long
Dim typTI       As TOOLINFO
    With typTI
        .cbSize = Len(typTI)
        .uFlags = TTF_IDISHWND + TTF_SUBCLASS
        .hWnd = pobjTip.ControlhWnd
        .uId = pobjTip.ControlhWnd
        .lpszText = pobjTip.Text
    End With
    pobjTip.hWnd = CreateWindowEx(0&, "tooltips_class32", "", TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, pobjTip.FormhWnd, 0&, App.hInstance, 0&)
    SendMessage pobjTip.hWnd, TTM_ADDTOOL, 0&, typTI
    lngStyle = GetWindowLong(pobjTip.hWnd, GWL_STYLE)
    lngStyle = lngStyle And (Not WS_BORDER)
    SetWindowLong pobjTip.hWnd, GWL_STYLE, lngStyle
    HookWindow pobjTip
End Sub

Public Sub DestroyToolTip(ByRef pobjTip As ToolTip)
Dim typTI       As TOOLINFO
    If pobjTip.hWnd > 0 Then
        UnhookWindow pobjTip
        With typTI
            .cbSize = Len(typTI)
            .uFlags = TTF_IDISHWND + TTF_SUBCLASS
            .hWnd = pobjTip.ControlhWnd
            .uId = pobjTip.ControlhWnd
            .lpszText = pobjTip.Text
        End With
        SendMessage pobjTip.hWnd, TTM_DELTOOL, 0&, typTI
        DestroyWindow pobjTip.hWnd
        pobjTip.hWnd = 0
    End If
End Sub

Public Function GetDefaultBalloonColor() As Long
    GetDefaultBalloonColor = GetSysColor(COLOR_INFOBK)
End Function

Public Function GetDefaultTextColor() As Long
    GetDefaultTextColor = GetSysColor(COLOR_INFOTEXT)
End Function

Public Function GetDefaultBorderColor() As Long
    GetDefaultBorderColor = GetStockObject(DKGRAY_BRUSH)
End Function

Public Function GetDefaultShadowColor() As Long
    GetDefaultShadowColor = GetSysColor(COLOR_BTNSHADOW)
End Function
