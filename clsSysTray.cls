VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Written by Ellis Dee of vbforums.com
'
' Thanks to...
'   Edgemeal   http://www.vbforums.com/showthread.php?t=595193
'   Mike D. Sutton   http://www.mvps.org/EDais/
'   Steve McMahon   http://www.vbaccelerator.com/home/VB/Code/Libraries/Shell_Projects/SysTray_-_The_Easy_Way/article.asp
'   Ben Baird for the original code
Option Explicit

' Tray Icon
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIF_GUID = &H20
Private Const NIF_REALTIME = &H40
Private Const NIF_SHOWTIP = &H80
    
    
Private Const NIIF_LARGE_ICON = &H20

'    Const NIF_TIP = &H4
'    Const NIM_MODIFY = &H1
    
'Private Const WM_MOUSEMOVE = &H200

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 4

Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

'Private Type GUID
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(7) As Byte
'End Type

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
    guidItem As GUID
    'hBalloonIcon As Long
End Type
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwHandle As Long, _
    ByVal dwLen As Long, lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" _
    (pBlock As Any, ByVal lpSubBlock As String, _
    lpBuffer As Any, nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, ByVal Source As Long, ByVal length As Long)

'Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID)
'Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal rguid As Long) As Long

' Bitmap to Icon
Private Type IconInfo
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (ByRef pIconInfo As IconInfo) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, _
        ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
        lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, _
        ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, ByVal lpAddress As Long, _
        ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, _
        ByVal lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type

' Tray Events (fire in WinMetersTray.frm)
Public Event LeftClick()
Public Event RightClick()
Public Event DoubleClick()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()

Private frmOwner As Form

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400
Private Const WM_DESTROY = &H2


Private Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved1 As Byte
    bReserved2 As Byte
    dwData As Long
    iString As Long
End Type

Private Const TB_BUTTONCOUNT = WM_USER + 24
Private Const TB_GETBUTTON = WM_USER + 23
Private Const TB_GETITEMRECT = WM_USER + 29

Private Const MEM_COMMIT = &H1000
Private Const MEM_RELEASE = &H8000
Private Const PAGE_READWRITE = &H4
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM = PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE

Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' Ballon tips
Public Enum BalloonIconEnum
   beNone = 0
   beInformation = 1
   beWarning = 2
   beError = 3
   beNoSound = &H10
End Enum

Private Type TrayData
    hwnd As Long
    ID As Long
End Type

Private trayRECT As RECT
Public trayRECT_Top As Long
Public trayRECT_Bottom As Long
Public trayRECT_Left As Long
Public trayRECT_Right As Long

Private mstrTooltip As String

Private mtypIcon As NOTIFYICONDATA, snWinVer As Single
'Public frmTip As frmTooltip


' METHODS


' Set tray icon to contents of picturebox
Public Sub DrawIcon(ByRef ppic As PictureBox, Optional ByVal plngTransparentColor As Long = -1, Optional ByVal hndlIcon As Long = 0)
    If hndlIcon = 0 Then
        ppic.AutoRedraw = True
        ppic.Picture = ppic.Image
        If plngTransparentColor < 0 Then
            mtypIcon.hIcon = BitmapToIcon(ppic.Picture.handle)
        Else
            mtypIcon.hIcon = BitmapToIconTransparent(ppic.Picture.handle, plngTransparentColor, ppic.Width, ppic.Height)
        End If
    Else
        mtypIcon.hIcon = hndlIcon
    End If
    If mtypIcon.hIcon <> 0 Then
        RefreshIcon
        If hndlIcon = 0 Then DestroyIcon mtypIcon.hIcon
    End If
End Sub

Public Function Init(pfrm As Form, pstrTooltip As String, w7GUID As String, Optional ByVal hndlIcon As Long = 0) As Boolean
    
    GetWindowsVersion snWinVer
    
    With mtypIcon
        '.cbSize = Len(mtypIcon)
        .cbSize = GetIconDataSize
        .hwnd = pfrm.hwnd
        If hndlIcon = 0 Then
            .hIcon = pfrm.Icon
        Else
            .hIcon = hndlIcon
        End If
        .uId = vbNull
        .uFlags = NIF_MESSAGE
        '.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        'Call CoCreateGuid(.guidItem)
        'If snWinVer >= 6# Then .dwInfoFlags = .dwInfoFlags Or NIIF_LARGE_ICON ' on Vista and above trying to use large icons
        If snWinVer >= 6.1 Then ' on Win7 we must use GUID for every icons created
            'Call CoCreateGuid(.guidItem)
            GetGUIDfromString w7GUID, .guidItem
            .uFlags = .uFlags Or NIF_GUID
        End If
        .uTimeoutOrVersion = NOTIFYICON_VERSION
    End With
    Shell_NotifyIcon NIM_ADD, mtypIcon
    'Shell_NotifyIcon NIM_SETVERSION, mtypIcon
    SetIcon pfrm
    Me.TooltipText = pstrTooltip
    'Set frmTip = New frmTooltip
    'frmTip.lblTooltipText = pstrTooltip
    Set frmOwner = pfrm
End Function

Public Sub MouseMove(Button As Integer, ByVal x As Long, pfrm As Form)
   Select Case x
        Case WM_LBUTTONDBLCLK: RaiseEvent DoubleClick
        Case WM_LBUTTONUP: RaiseEvent LeftClick
        Case WM_RBUTTONUP: RaiseEvent RightClick
        Case NIN_BALLOONHIDE: RaiseEvent BalloonHide
        Case NIN_BALLOONTIMEOUT: RaiseEvent BalloonTimeOut
        Case NIN_BALLOONUSERCLICK: RaiseEvent BalloonClicked
   End Select
End Sub

' Set tray icon to a form icon
Public Sub SetIcon(pfrm As Form, Optional ByVal hndlIcon As Long = 0)
    If hndlIcon = 0 Then
        mtypIcon.hIcon = pfrm.Icon
    Else
        mtypIcon.hIcon = hndlIcon
    End If
    RefreshIcon
End Sub

Public Function GetIcon() As Long
    GetIcon = mtypIcon.hIcon
End Function

'Public Sub ShowBalloonTip(ByVal pstrMessage As String, Optional ByVal penIcon As BalloonIconEnum, Optional ByVal pstrTitle As String, Optional ByVal plngTimeout = 30000)
'    mtypIcon.szInfo = pstrMessage & vbNullChar
'    mtypIcon.szInfoTitle = pstrTitle & vbNullChar
'    mtypIcon.uTimeoutOrVersion = plngTimeout
'    mtypIcon.dwInfoFlags = penIcon
'    mtypIcon.uFlags = NIF_INFO 'Or NIF_REALTIME
'    'If snWinVer >= 6# Then mtypIcon.uFlags = mtypIcon.uFlags Or NIF_REALTIME
'    If snWinVer >= 6.1 Then mtypIcon.uFlags = mtypIcon.uFlags Or NIF_GUID
'    Shell_NotifyIcon NIM_MODIFY, mtypIcon
'End Sub

Public Sub GetTrayIconRect(ByVal pX As Long, ByVal pY As Long)
    trayRECT = GetSystemTrayItemRect(mtypIcon.hwnd, 1, pX, pY)
    trayRECT_Bottom = trayRECT.Bottom
    trayRECT_Left = trayRECT.Left
    trayRECT_Right = trayRECT.Right
    trayRECT_Top = trayRECT.Top
End Sub


' PROPERTIES

Public Property Get TooltipText() As String
    TooltipText = mstrTooltip
End Property

Public Property Let TooltipText(ByVal pstrTooltip As String)
    mstrTooltip = pstrTooltip
    mtypIcon.szTip = mstrTooltip & vbNullChar
    mtypIcon.uFlags = NIF_TIP
    'mtypIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    If snWinVer >= 6# Then mtypIcon.uFlags = mtypIcon.uFlags Or NIF_SHOWTIP
    If snWinVer >= 6.1 Then mtypIcon.uFlags = mtypIcon.uFlags Or NIF_GUID
    Shell_NotifyIcon NIM_MODIFY, mtypIcon
    'frmTip.lblTooltipText = mstrTooltip
End Property


' INTERNAL ROUTINES


Public Sub Class_Terminate()
    'Const NIM_DELETE = &H2
    
    Shell_NotifyIcon NIM_DELETE, mtypIcon
    'RefreshIcon
    Set frmOwner = Nothing
    'Set frmTip = Nothing
End Sub

Private Sub RefreshIcon()
    'Const NIF_ICON = &H2
    'Const NIM_MODIFY = &H1
    
    mtypIcon.uFlags = NIF_ICON
    'mtypIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    If snWinVer >= 6.1 Then mtypIcon.uFlags = mtypIcon.uFlags Or NIF_GUID
    Shell_NotifyIcon NIM_MODIFY, mtypIcon
End Sub

' Birmap to icon functions provided by Edgemeal
' Special Thanks to Mike D. Sutton - Http://www.mvps.org/EDais/
Private Function BitmapToIcon(ByVal inBMP As Long) As Long
    Dim IconInf As IconInfo
    Dim BMInf As BITMAP
    Dim hMask As Long
    
    ' Get some information about this Bitmap and create a mask the same size
    If (GetObject(inBMP, Len(BMInf), BMInf) = 0) Then Exit Function
    hMask = CreateBitmap(BMInf.bmWidth, BMInf.bmHeight, 0, 0, ByVal 0&)
    With IconInf ' Set some information about the icon
        .fIcon = True
        .hbmMask = hMask
        .hbmColor = inBMP
    End With
    ' Create the icon and destroy the temp mask
    BitmapToIcon = CreateIconIndirect(IconInf)
    Call DeleteObject(hMask)
End Function

' Take a HBITMAP and return an HICON
Private Function BitmapToIconTransparent(ByVal hSrcBMP As Long, Optional ByVal inTransCol As Long = -1, _
            Optional ByVal inPictWidth As Integer = 32, Optional ByVal intPictHeght As Integer = 32) As Long
    Dim IconInf As IconInfo
    Dim hSrcDC As Long ', hSrcBMP As Long
    Dim hSrcOldBMP As Long
    Dim hMaskDC As Long
    Dim hMaskBMP As Long
    Dim hMaskOldBMP As Long
    
    ' Create DC's and select source copy
    hSrcDC = CreateCompatibleDC(0)
    hMaskDC = CreateCompatibleDC(0)
    hSrcOldBMP = SelectObject(hSrcDC, hSrcBMP)
    If (hSrcOldBMP) Then ' Extract a colour mask from source copy
        hMaskBMP = GetColMask(hSrcDC, inTransCol, inPictWidth, intPictHeght)
        hMaskOldBMP = SelectObject(hMaskDC, hMaskBMP)
        If (hMaskOldBMP) Then ' Overlay inverted mask over source
            Call SetTextColor(hSrcDC, vbWhite)
            Call SetBkColor(hSrcDC, vbBlack)
            Call BitBlt(hSrcDC, 0, 0, inPictWidth, intPictHeght, hMaskDC, 0, 0, vbSrcAnd)
            Call SelectObject(hMaskDC, hMaskOldBMP) ' De-select mask
        End If
        ' De-select source copy
        Call SelectObject(hSrcDC, hSrcOldBMP)
    End If
    ' Destroy DC's
    Call DeleteDC(hMaskDC)
    Call DeleteDC(hSrcDC)
    With IconInf ' Set some information about the icon
        .fIcon = True
        .hbmMask = hMaskBMP
        .hbmColor = hSrcBMP
    End With
    ' Create the icon and destroy the temp mask
    BitmapToIconTransparent = CreateIconIndirect(IconInf)
    ' Destroy interim Bitmaps
    Call DeleteObject(hMaskBMP)
    Call DeleteObject(hSrcBMP)
End Function

Private Function GetColMask(ByVal inDC As Long, ByVal inMaskCol As Long, _
                Optional ByVal inPictWidth As Integer = 32, Optional ByVal intPictHeght As Integer = 32) As Long
    Dim MaskDC As Long, MaskBMP As Long, OldMask As Long, OldBack As Long
    ' Create a new DC
    MaskDC = CreateCompatibleDC(inDC)
    If (MaskDC) Then ' Create a new 1-bpp Bitmap (DDB)
        MaskBMP = CreateBitmap(inPictWidth, intPictHeght, 1, 1, ByVal 0&)
        If (MaskBMP) Then ' Select Bitmap into DC
            OldMask = SelectObject(MaskDC, MaskBMP)
            If (OldMask) Then ' Set mask colour
                OldBack = SetBkColor(inDC, inMaskCol)
                ' Generate mask image
                If (BitBlt(MaskDC, 0, 0, inPictWidth, intPictHeght, inDC, 0, 0, vbSrcCopy) <> 0) Then GetColMask = MaskBMP
                ' Clean up
                Call SetBkColor(inDC, OldBack)
                Call SelectObject(MaskDC, OldMask)
            End If
            ' Something went wrong, destroy mask Bitmap
            If (GetColMask = 0) Then Call DeleteObject(MaskBMP)
        End If
        ' Destroy temporary DC
        Call DeleteDC(MaskDC)
    End If
End Function


Private Function GetIconDataSize() As Long
   Dim nVerMajor As Integer
   Dim nVerMinor As Integer
   Dim nVerDelta As Integer

   GetShellVersion nVerMajor, nVerMinor, nVerDelta
   Select Case nVerMajor
      Case 6
         If nVerMinor = 0 And nVerDelta < 5000 Then
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
         Else
            NOTIFYICONDATA_SIZE = Len(mtypIcon)
         End If
      Case 5
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      Case Is <= 4
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select
   GetIconDataSize = NOTIFYICONDATA_SIZE
End Function


Private Function GetShellVersion(ByRef versionMajor As Integer, ByRef versionMinor As Integer, ByRef versionDelta As Integer) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim bBuffer() As Byte
   Dim udtVerBuffer As VS_FIXEDFILEINFO ', FileVer As String, Prodver As String
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\\", lpBuffer, nUnused) = 1 Then
         
         'CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
         'CopyMemory udtVerBuffer, lpBuffer, Len(udtVerBuffer)
         MoveMemory udtVerBuffer, lpBuffer, Len(udtVerBuffer)
        
        versionMajor = udtVerBuffer.dwFileVersionMSh
        versionMinor = udtVerBuffer.dwFileVersionMSl
        versionDelta = udtVerBuffer.dwFileVersionLSh
        
        '**** Determine File Version number ****
        'FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & _
            "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
        
        '**** Determine Product Version number ****
        'Prodver = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) _
            & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." _
            & Format$(udtVerBuffer.dwProductVersionLSl)
    
         GetShellVersion = True
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

'' Returns the version of Windows that the user is running
'Private Function GetWindowsVersion(Optional ByRef snVersion As Single) As String
'    Dim osv As OSVERSIONINFO
'    osv.OSVSize = Len(osv)
'
'    If GetVersionEx(osv) = 1 Then
'        Select Case osv.PlatformID
'            Case VER_PLATFORM_WIN32s
'                GetWindowsVersion = "Win32s on Windows 3.1"
'            Case VER_PLATFORM_WIN32_NT
'                GetWindowsVersion = "Windows NT"
'                If Not IsMissing(snVersion) Then snVersion = CSng(CStr(osv.dwVerMajor) & GetDecimalSeparator() & CStr(osv.dwVerMinor))
'                Select Case osv.dwVerMajor
'                    Case 3
'                        GetWindowsVersion = "Windows NT 3.5"
'                    Case 4
'                        GetWindowsVersion = "Windows NT 4.0"
'                    Case 5
'                        Select Case osv.dwVerMinor
'                            Case 0
'                                GetWindowsVersion = "Windows 2000"
'                            Case 1
'                                GetWindowsVersion = "Windows XP"
'                            Case 2
'                                GetWindowsVersion = "Windows Server 2003"
'                        End Select
'                    Case 6
'                        Select Case osv.dwVerMinor
'                            Case 0
'                                GetWindowsVersion = "Windows Vista/Server 2008"
'                            Case 1
'                                GetWindowsVersion = "Windows 7/Server 2008 R2"
'                            Case 2
'                                GetWindowsVersion = "Windows 8/Server 2012"
'                        End Select
'                End Select
'
'            Case VER_PLATFORM_WIN32_WINDOWS:
'                Select Case osv.dwVerMinor
'                    Case 0
'                        GetWindowsVersion = "Windows 95"
'                    Case 90
'                        GetWindowsVersion = "Windows Me"
'                    Case Else
'                        GetWindowsVersion = "Windows 98"
'                End Select
'        End Select
'    Else
'        GetWindowsVersion = "Unable to identify your version of Windows."
'    End If
'End Function

'Private Function GetDecimalSeparator() As String
'    Dim iLocale As Integer, sTmpStr As String, lRes As Long, aLen As Long
'    On Error Resume Next
'    sTmpStr = String$(255, " ") & Chr$(0)
'    aLen = 1
'    iLocale = GetUserDefaultLCID()
'    lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, aLen)
'    GetDecimalSeparator = Left$(sTmpStr, aLen)
'    Err.Clear
'    On Error GoTo 0
'End Function

'Private Function GetGUID() As String
'    Dim MyGUID As GUID
'    Dim GUIDByte() As Byte
'    Dim GuidLen As Long
'
'    CoCreateGuid MyGUID
'
'    ReDim GUIDByte(80)
'    GuidLen = StringFromGUID2(VarPtr(MyGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
'
'    GetGUID = Left(GUIDByte, GuidLen)
'End Function

Private Function GetGUIDfromString(ByVal sGUID As String, ByRef Result As GUID) As Boolean
    Dim NewGUID As GUID
    Dim GuidRes As Long
    GuidRes = CLSIDFromString(StrPtr(sGUID), VarPtr(NewGUID.Data1))
    If GuidRes = 0 Then
        Result = NewGUID
        GetGUIDfromString = True
    Else
        GetGUIDfromString = False
    End If
End Function

 
Private Function GetSystemTrayItemRect(ByVal hwnd As Long, ByVal ID As Long, Optional ByVal posX As Long, Optional ByVal posY As Long) As RECT
    Dim hWndTray As Long
    Dim hProc As Long
    Dim pid As Long
    Dim vaPtr As Long
    Dim ret As Long
    Dim tbut As TBBUTTON
    Dim cButtons As Long
    Dim td As TrayData
    Dim i As Long
    Dim rc As RECT
     
    hWndTray = modMain.GethWndTray
    If hWndTray = 0 Then Exit Function
     
    Call GetWindowThreadProcessId(hWndTray, pid)
    If pid = 0 Then Exit Function
    
    hProc = OpenProcess(PROCESS_VM, 0, pid)
    If hProc = 0 Then Exit Function
     
    vaPtr = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)
    If vaPtr = 0 Then GoTo cleanup
     
    cButtons = SendMessage(hWndTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
    
    On Error GoTo cleanup
    
    For i = 0 To cButtons - 1
        Call SendMessage(hWndTray, TB_GETBUTTON, i, ByVal vaPtr)
        Call ReadProcessMemory(hProc, ByVal vaPtr, tbut, Len(tbut), ret)
         
        If Not tbut.dwData = 0 Then
            Call ReadProcessMemory(hProc, ByVal tbut.dwData, td, Len(td), ret)
            If hwnd = td.hwnd Then
                If ID = td.ID Then
                    Call SendMessage(hWndTray, TB_GETITEMRECT, i, ByVal vaPtr)
                    Call ReadProcessMemory(hProc, ByVal vaPtr, rc, Len(rc), ret)
                    Exit For
                End If
            End If
        End If
    Next i
    Call MapWindowPoints(hWndTray, 0&, rc, 2)
    If i = cButtons Then 'item not found or data is wrong (mostly in Win7 (sucks win7!!!))
        For i = 0 To cButtons - 1
            Call SendMessage(hWndTray, TB_GETITEMRECT, i, ByVal vaPtr)
            Call ReadProcessMemory(hProc, ByVal vaPtr, rc, Len(rc), ret)
            Call MapWindowPoints(hWndTray, 0&, rc, 2)
            If rc.Left <= posX And rc.Right >= posX _
                    And rc.Top <= posY And rc.Bottom >= posY Then Exit For
        Next i
    End If
    
    GetSystemTrayItemRect = rc
    
cleanup:
    If hProc Then
        If vaPtr Then
            Call VirtualFreeEx(hProc, ByVal vaPtr, 0&, MEM_RELEASE)
        End If
    End If
    If hProc Then CloseHandle (hProc)
End Function


Public Function GetSysTrayIconWidth(Optional ByRef AlsoGetSysTrayIconHeight As Long) As Long
    Dim hWndTray As Long
    Dim hProc As Long
    Dim pid As Long
    Dim vaPtr As Long
    Dim ret As Long
    Dim tbut As TBBUTTON
    Dim cButtons As Long
    Dim td As TrayData
    Dim i As Long
    Dim rc As RECT
     
    hWndTray = modMain.GethWndTray
    If hWndTray = 0 Then Exit Function
     
    Call GetWindowThreadProcessId(hWndTray, pid)
    If pid = 0 Then Exit Function
    
    hProc = OpenProcess(PROCESS_VM, 0, pid)
    If hProc = 0 Then Exit Function
     
    vaPtr = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)
    If vaPtr = 0 Then GoTo cleanup
     
    cButtons = SendMessage(hWndTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
    
    On Error GoTo cleanup
    
    For i = 0 To cButtons - 1
        Call SendMessage(hWndTray, TB_GETITEMRECT, i, ByVal vaPtr)
        Call ReadProcessMemory(hProc, ByVal vaPtr, rc, Len(rc), ret)
        Call MapWindowPoints(hWndTray, 0&, rc, 2)
        If (rc.Left <> rc.Right) Or (rc.Top <> rc.Bottom) Then Exit For
    Next i
    
    GetSysTrayIconWidth = Abs(rc.Right - rc.Left)
    If Not IsMissing(AlsoGetSysTrayIconHeight) Then AlsoGetSysTrayIconHeight = Abs(rc.Bottom - rc.Top)
    
cleanup:
    If hProc Then
        If vaPtr Then
            Call VirtualFreeEx(hProc, ByVal vaPtr, 0&, MEM_RELEASE)
        End If
    End If
    If hProc Then CloseHandle (hProc)
End Function

