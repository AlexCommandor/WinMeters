Attribute VB_Name = "modMain"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public IsFirstRun As Integer
Public nCPUs As Integer, lMem As Long, vTrays() As Form, i As Integer
Public sExtraCPUsInfo As String, sExtraMemoryInfo As String, sExtraHDDsInfo As String
Public Const Pi As Double = 3.14159265
Public ShowOnlyTotalCPUUsage As Boolean
Public ShowKernelMemory As Boolean
Public ShowSplashScreen As Boolean
Public strCPUsNames() As String
Public CoresPerCPU() As Integer
Public strCPUsCoresNames() As String
Public intCPUs As Integer, intMEM As Integer, intHDD As Integer, intNet As Integer ', arrBlackCircle(0 To 31, 0 To 31) As Byte
Public rgbCPUsUser As OLE_COLOR, rgbCPUsKernel As OLE_COLOR, rgbCPUsUser2 As OLE_COLOR, rgbCPUsKernel2 As OLE_COLOR
Public rgbMEM As OLE_COLOR, rgbMEMKernel As OLE_COLOR
Public rgbdevRead As OLE_COLOR, rgbdevWrite As OLE_COLOR, rgbTemp As OLE_COLOR
Public sNetworks() As String, NetworkPresent As Boolean, sActiveNetwork As String, NetworkInterfaceIsMissing As Boolean
Public ExtendedHDDInfo As Boolean
Public hwIcon(0 To 3) As Long, netIcon(0 To 3) As Long, lIconIndex As Integer, memIcon As Long
Public IndicatorEnabled(1 To 4) As Boolean, AntialiasedMEMIndicator As Boolean
'Icon index is: 0 - 128x128, 1 - 64x64, 2 - 48x48, 3 - 32x32, 4 - 28x28, 5 - 24x24, 6 - 20x20, 7 - 16x16
Public Const colDefaultCPUUser As Long = 13132800 ' -RGB(0, 100, 200)    '13123584 '&HC84000 'RGB(0, 64, 200)
Public Const colDefaultCPUUser2 As Long = 16776960 '&HFFFF00 'RGB(0, 255, 255)
Public Const colDefaultCPUKernel As Long = 220 '&HDC     'RGB(220, 0, 0)
Public Const colDefaultCPUKernel2 As Long = 65535 '&HFFFF   'RGB(255, 255, 0)
Public Const colDefaultMEMUsage As Long = 13132800 ' -RGB(0, 100, 200)    '13123584 '&HC84000 'RGB(0, 64, 200)
Public Const colDefaultMEMKernel As Long = 220 '&HDC 'RGB(220, 0, 0)
        
Public Const fontDefNameCPU As String = "Tahoma"
'Public Const fontDefNameCPU As String = "Small Fonts"
Public Const fontDefSizeCPU As Integer = 8
Public Const fontDefBoldCPU As Boolean = False
Public Const fontDefItalicCPU As Boolean = False
Public FontNameCPU As String, FontSizeCPU As Integer, FontBoldCPU As Boolean, FontItalicCPU As Boolean
'Public bFontUnderlineCPU as Boolean,
'Public iconMinWidth As Long, iconMinHeight As Long

Public ShowDigitsInsteadThermometer As Boolean, ShowAdvancedTooltips As Boolean
Public ShowSolidColors As Boolean, ShowOnlyTotalCPULoad As Boolean

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum tDeviceStatus
    devIdle = 0
    devRead = 1
    devWrite = 2
    devReadWrite = 3
End Enum

Public Type NOTIFYICONDATA
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
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwflags As Long
    szExeFile As String * 260
End Type

'Private Type core_temp_shared_data
'    uiLoad(1 To 256) As Long
'    uiTjMax(1 To 128) As Long
'    uiCoreCnt  As Long
'    uiCPUCnt  As Long
'    fTemp(1 To 256) As Single
'    fVID  As Single
'    fCPUSpeed  As Single
'    fFSBSpeed  As Single
'    fMultiplier  As Single
'    sCPUName(1 To 100) As Byte
'    ucFahrenheit As Byte 'Boolean
'    ucDeltaToTjMax As Byte 'Boolean
'End Type

Public Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" _
    (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'"R:\!_Work\Alex_K\Progs\WinMeters\GetCoreTempInfo.dll"
'Private Declare Function fnGetCoreTempInfoAlt Lib "GetCoreTempInfo.dll" (ByRef pData As core_temp_shared_data) As Byte
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const WM_NCACTIVATE = &H86
Private Const SC_CLOSE = &HF060
Private Const MF_BYCOMMAND = &H0&

'Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID)
'Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
'Private Declare Function StringFromCLSID Lib "ole32.dll" (ByVal rguid As Long, ByRef lpsz As Long) As Long
'Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal rguid As Long) As Long
Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
        ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const flags = SWP_NOSIZE Or SWP_NOMOVE

Public Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type

Public Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Public Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Public Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Public Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Public Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Public Enum tGradientType
    gradHorizontal = GRADIENT_FILL_RECT_H
    gradVertical = GRADIENT_FILL_RECT_V
    gradTriangle = GRADIENT_FILL_TRIANGLE
End Enum

Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
            (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
            pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
            
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, _
            ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
            
Public Declare Function AngleArc Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
            ByVal dwRadius As Long, ByVal eStartAngle As Single, ByVal eSweepAngle As Single) As Long
            
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, _
            ByVal x As Long, ByVal y As Long) As Long
            
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, _
            ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
            
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Public Const FLOODFILLBORDER = 0  ' Fill until crColor& color encountered.
Public Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not encountered.


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
 
Public Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID)
Public Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal rguid As Long) As Long

Public Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
 
Public Declare Function CreateIconFromResourceEx Lib "user32" (pbIconBits As Byte, ByVal cbIconBits As Long, _
            ByVal fIcon As Long, ByVal dwVersion As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long
Public Declare Function CreateIconFromResource Lib "user32" (pbIconBits As Byte, ByVal cbIconBits As Long, _
            ByVal fIcon As Long, ByVal dwVersion As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PictDesc, riid As GUID, _
    ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_SHARED = &H8000
 
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Public Const DT_WORDBREAK = &H10
'Public Const DT_CENTER = &H1
'Public Const DT_VCENTER = &H4
'Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
 
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const LOCALE_SISO639LANGNAME = &H59               ' ISO abbreviated language name
Public Const LOCALE_SDECIMAL = &HE         '  decimal separator
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetOEMCP Lib "kernel32" () As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetACP Lib "kernel32" () As Long


Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
    (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


Private Type LUID
  UsedPart As Long
  IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  TheLuid As LUID
  Attributes As Long
End Type

      'The GetCurrentProcess function returns a pseudohandle for the
      'current process.
      Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

      'The OpenProcessToken function opens the access token associated with
      'a process.
      Private Declare Function OpenProcessToken Lib "advapi32" _
         (ByVal ProcessHandle As Long, _
          ByVal DesiredAccess As Long, _
          TokenHandle As Long) As Long

      'The LookupPrivilegeValue function retrieves the locally unique
      'identifier (LUID) used on a specified system to locally represent
      'the specified privilege name.
      Private Declare Function LookupPrivilegeValue Lib "advapi32" _
         Alias "LookupPrivilegeValueA" _
         (ByVal lpSystemName As String, _
          ByVal lpName As String, _
          lpLuid As LUID) As Long

      'The AdjustTokenPrivileges function enables or disables privileges
      'in the specified access token. Enabling or disabling privileges
      'in an access token requires TOKEN_ADJUST_PRIVILEGES access.
      Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
         (ByVal TokenHandle As Long, _
          ByVal DisableAllPrivileges As Long, _
          NewState As TOKEN_PRIVILEGES, _
          ByVal BufferLength As Long, _
          PreviousState As TOKEN_PRIVILEGES, _
          ReturnLength As Long) As Long
      
      Private Declare Sub SetLastError Lib "kernel32" _
         (ByVal dwErrCode As Long)

Public IsIDE As Boolean

'Public objDraw As New LineGS
'Public objWMIServiceForCPU As Object
'Public objWMIServiceForMEM As Object
'Public objWMIServiceForHDD As Object
'Public objWMIServiceForNET As Object
'Public objRefresherCPU As Object
'Public objRefresherMEM As Object
'Public objRefresherHDD As Object
'Public objRefresherNET As Object
'Public objParentCPU As Object
'Public objParentMEM As Object
'Public objParentHDD As Object
'Public objParentNET As Object
'Public objHDD As Object ', objChild As Object
'Public objParent2 As Object, objParent3 As Object


Public Function LongToUShort(Unsigned As Long) As Integer
    'A small function to convert from long to unsigned short
    LongToUShort = CInt(Unsigned - &H10000)
End Function

Public Sub DoGradient(ByRef Objhdc As Object, ByVal colorFrom As OLE_COLOR, ByVal colorTo As OLE_COLOR, _
                XpointFrom As Long, YpointFrom As Long, XpointTo As Long, YpointTo As Long, gradientType As tGradientType)
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    'Dim iScaleMode As Long
    'Objhdc.AutoRedraw = True
    'iScaleMode = Objhdc.ScaleMode
    'Objhdc.ScaleMode = vbPixels

    With vert(0)
        .x = XpointFrom
        .y = YpointFrom
    End With
    GradientFillColor vert(0), colorFrom

    With vert(1)
        .x = XpointTo + 1
        .y = YpointTo + 1
    End With
    GradientFillColor vert(1), colorTo

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    GradientFillRect Objhdc.hDC, vert(0), 2, gRect, 1, gradientType
    'Objhdc.ScaleMode = iScaleMode
End Sub

Private Sub GradientFillColor(ByRef tTV As TRIVERTEX, ByVal iColor As OLE_COLOR)
  Dim iRed   As Long
  Dim iGreen As Long
  Dim iBlue  As Long

    '/* Separate color into RGB
    iRed = (iColor And &HFF&) * &H100&
    iGreen = (iColor And &HFF00&)
    iBlue = (iColor And &HFF0000) \ &H100&
    
    '/* Make Red color a UShort
    If (iRed And &H8000&) = &H8000& Then
       tTV.Red = (iRed And &H7F00&)
       tTV.Red = tTV.Red Or &H8000
    Else
       tTV.Red = iRed
    End If
    '/* Make Green color a UShort
    If (iGreen And &H8000&) = &H8000& Then
       tTV.Green = (iGreen And &H7F00&)
       tTV.Green = tTV.Green Or &H8000
    Else
       tTV.Green = iGreen
    End If
    '/* Make Blue color a UShort
    If (iBlue And &H8000&) = &H8000& Then
       tTV.Blue = (iBlue And &H7F00&)
       tTV.Blue = tTV.Blue Or &H8000
    Else
       tTV.Blue = iBlue
    End If
    
    tTV.Alpha = 0

End Sub

Private Function GetRed(ByVal color As OLE_COLOR) As Integer
    GetRed = color And 255
End Function

Private Function GetGreen(ByVal color As OLE_COLOR) As Integer
    GetGreen = (color And 65280) \ 256
End Function

Private Function GetBlue(ByVal color As OLE_COLOR) As Integer
    GetBlue = (color And 16711680) \ 65535
End Function

Public Function GenerateRandomGUID() As String
    Dim myGUID As GUID, NewGUID As GUID
    Dim GUIDByte() As Byte, sGUID As String
    Dim GuidLen As Long
    
    CoCreateGuid myGUID
    
    ReDim GUIDByte(80)
    GuidLen = StringFromGUID2(VarPtr(myGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
    
    sGUID = Left(GUIDByte, GuidLen)
    
    
    GuidLen = CLSIDFromString(StrPtr(sGUID), VarPtr(NewGUID.Data1))
    
    If Asc(Right$(sGUID, 1)) = 0 Then sGUID = Left$(sGUID, Len(sGUID) - 1)
    
    GenerateRandomGUID = sGUID
End Function

Public Function GetGUIDfromString(ByVal sGUID As String, ByRef Result As GUID) As Boolean
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

Public Function GetStringFromGUID(ByRef tGUID As GUID) As String
    Dim myGUID As GUID
    Dim GUIDByte() As Byte, sGUID As String
    Dim GuidLen As Long
    ReDim GUIDByte(80)
    myGUID = tGUID
    GuidLen = StringFromGUID2(VarPtr(myGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
    sGUID = Left(GUIDByte, GuidLen)
    If Asc(Right$(sGUID, 1)) = 0 Then sGUID = Left$(sGUID, Len(sGUID) - 1)
    GetStringFromGUID = sGUID
End Function

Private Function CheckIDE() As Boolean
    'This function will never be executed in an EXE
    IsIDE = True        'set global flag
    'Set CheckIDE or the Debug.Assert will Break
    CheckIDE = True
End Function

Private Sub Main()
    Dim snWinVer As Single
    
    'AdjustToken
    
    'frmNet.Show
    
'    MsgBox GetOEMCP
'    Stop

    InitLanguage
    
    log.OpenLog
    log.WriteLog "Sub Main: WinMeters is starting..."
    
    GetWindowsVersion snWinVer
    
    log.WriteLog "Sub Main: Determined Windows version: " & snWinVer
    
    If snWinVer < 5 Then
        log.WriteLog "Sub Main: Incompatible Windows version: " & snWinVer & ", exiting."
        log.CloseLog
        MsgBox "WinMeters works only under Microsoft Windows XP and above. Sorry.", vbCritical, "WinMeters critical error"
        End
    End If
    
    IsIDE = False
    'This line is only executed if running in the IDE and then returns True
    Debug.Assert CheckIDE
    'Use the IsIDE flag anywhere
    If Not IsIDE Then
        If App.PrevInstance Or (FindHWND(GethWndTray(True), "WinMeters - running") <> 0) Then
            log.WriteLog "Sub Main: Another instance of program is already running, exiting."
            log.CloseLog
            MsgBox "Another instance of program is already running!", vbExclamation + vbOKOnly, "WinMeters"
            End
        End If
    End If


    App.Title = "WinMeters - running"

    'Call FillBlackCircle
    'MsgBox GetGUID
    log.WriteLog "Sub Main: Start reading program parameters from registry..."
    ShowSplashScreen = GetSetting("WinMeters", "Trays", "Splash", "True")
    SaveSetting "WinMeters", "Trays", "Splash", ShowSplashScreen
    log.WriteLog "Sub Main: ShowSplashScreen - OK, value - " & ShowSplashScreen
     
    If ShowSplashScreen Then
        log.WriteLog "Sub Main: Starting with SplashScreen..."
        frmSplash.Show
        frmSplash.FadeIn
    Else
        log.WriteLog "Sub Main: Starting without SplashScreen..."
    End If
    
    ShowOnlyTotalCPUUsage = GetSetting("WinMeters", "CPU", "ShowTotal", "False")
    SaveSetting "WinMeters", "CPU", "ShowTotal", ShowOnlyTotalCPUUsage
    log.WriteLog "Sub Main: ShowOnlyTotalCPUUsage - OK, value - " & ShowOnlyTotalCPUUsage
    
    log.WriteLog "Sub Main: Getting CPU information..."
    nCPUs = GetCPUsCount(sExtraCPUsInfo)
    log.WriteLog "Sub Main: 'GetCPUsCount(sExtraCPUsInfo)' returned " & nCPUs & " cores and ExtraCPUsInfo '" & Replace$(sExtraCPUsInfo, vbCrLf, " ", 1, 5) & "'"
    'If IsIDE Then nCPUs = 0
    If nCPUs <= 0 Then
        log.WriteLog "Sub Main: Wrong number of CPU/cores! CPU(s) indicator(s) disabled."
    '    MsgBox "CPUs enumerating failed!!! Ensure you have working WMI and try again.", vbCritical, "ERROR"
    '    log.CloseLog
    '    End
    End If
    intCPUs = GetSetting("WinMeters", "CPU", "Interval", "0")
    If intCPUs = 0 Then
        intCPUs = 300
        SaveSetting "WinMeters", "CPU", "Interval", "300"
    End If
    log.WriteLog "Sub Main: Reading CPU refresh interval - OK, value - " & intCPUs
    
    rgbCPUsUser = GetSetting("WinMeters", "CPU", "User", "0")
    If rgbCPUsUser <= 0 Then
        rgbCPUsUser = colDefaultCPUUser 'RGB(0, 64, 200)
        SaveSetting "WinMeters", "CPU", "User", rgbCPUsUser
    End If
    log.WriteLog "Sub Main: Reading CPU indicator color 1 for user usage - OK, value - " & rgbCPUsUser
    'rgbCPUsUser = RGB(0, 64, 200)
    
    rgbCPUsUser2 = GetSetting("WinMeters", "CPU", "User2", "0")
    If rgbCPUsUser2 <= 0 Then
        rgbCPUsUser2 = colDefaultCPUUser2 'RGB(0, 255, 255)
        SaveSetting "WinMeters", "CPU", "User2", rgbCPUsUser2
    End If
    log.WriteLog "Sub Main: Reading CPU indicator color 2 for user usage - OK, value - " & rgbCPUsUser2
    
    'rgbCPUsKernel = vbRed
    rgbCPUsKernel = GetSetting("WinMeters", "CPU", "Kernel", "0")
    If rgbCPUsKernel <= 0 Then
        rgbCPUsKernel = colDefaultCPUKernel 'RGB(220, 0, 0)
        SaveSetting "WinMeters", "CPU", "Kernel", rgbCPUsKernel
    End If
    log.WriteLog "Sub Main: Reading CPU indicator color 1 for kernel usage - OK, value - " & rgbCPUsKernel
    
    rgbCPUsKernel2 = GetSetting("WinMeters", "CPU", "Kernel2", "0")
    If rgbCPUsKernel2 <= 0 Then
        rgbCPUsKernel2 = colDefaultCPUKernel2 'RGB(255, 255, 0)
        SaveSetting "WinMeters", "CPU", "Kernel2", rgbCPUsKernel2
    End If
    log.WriteLog "Sub Main: Reading CPU indicator color 2 for kernel usage - OK, value - " & rgbCPUsKernel2
    
    FontNameCPU = GetSetting("WinMeters", "CPU", "FontName", "0")
    If FontNameCPU = "0" Then
        FontNameCPU = fontDefNameCPU
        SaveSetting "WinMeters", "CPU", "FontName", FontNameCPU
    End If
    log.WriteLog "Sub Main: Reading CPU text indicator font name - OK, value - " & FontNameCPU
    
    FontSizeCPU = GetSetting("WinMeters", "CPU", "FontSize", "0")
    If FontSizeCPU = 0 Then
        FontSizeCPU = fontDefSizeCPU
        SaveSetting "WinMeters", "CPU", "FontSize", FontSizeCPU
    End If
    log.WriteLog "Sub Main: Reading CPU text indicator font size - OK, value - " & FontSizeCPU
    
    FontBoldCPU = GetSetting("WinMeters", "CPU", "FontBold", False)
    SaveSetting "WinMeters", "CPU", "FontBold", FontBoldCPU
    log.WriteLog "Sub Main: Reading CPU text indicator font bold - OK, value - " & FontBoldCPU
    
    FontItalicCPU = GetSetting("WinMeters", "CPU", "FontItalic", False)
    SaveSetting "WinMeters", "CPU", "FontItalic", FontItalicCPU
    log.WriteLog "Sub Main: Reading CPU text indicator font italic - OK, value - " & FontItalicCPU
    
    log.WriteLog "Sub Main: Getting MEMORY information..."
    lMem = GetMemTotal(sExtraMemoryInfo)
    log.WriteLog "Sub Main: 'GetMemTotal(sExtraMemoryInfo)' returned " & lMem & " MB and ExtraMemoryInfo '" & Replace$(sExtraMemoryInfo, vbCrLf, " ", 1, 10) & "'"
    'If IsIDE Then lMem = 0
    If lMem <= 0 Then
        log.WriteLog "Sub Main: Wrong memory data! Memory indicator disabled."
    '    MsgBox "Memory accessing failed!!! Ensure you have working WMI and try again.", vbCritical, "ERROR"
    '    log.CloseLog
    '    End
    End If
    
    intMEM = GetSetting("WinMeters", "MEM", "Interval", "0")
    If intMEM = 0 Then
        intMEM = 1000
        SaveSetting "WinMeters", "MEM", "Interval", "1000"
    End If
    log.WriteLog "Sub Main: Reading MEM refresh interval - OK, value - " & intMEM
    
    AntialiasedMEMIndicator = GetSetting("WinMeters", "MEM", "Smooth", "True")
    SaveSetting "WinMeters", "MEM", "Smooth", AntialiasedMEMIndicator
    log.WriteLog "Sub Main: Reading MEM antialiasing - OK, value - " & AntialiasedMEMIndicator
    
    rgbMEM = GetSetting("WinMeters", "MEM", "Usage", "0")
    If rgbMEM <= 0 Then
        rgbMEM = colDefaultMEMUsage 'RGB(0, 64, 200)
        SaveSetting "WinMeters", "MEM", "Usage", rgbMEM
    End If
    log.WriteLog "Sub Main: Reading MEM indicator color for user usage - OK, value - " & rgbMEM
    
    rgbMEMKernel = GetSetting("WinMeters", "MEM", "Kernel", "0")
    If rgbMEMKernel <= 0 Then
        rgbMEMKernel = colDefaultMEMKernel 'RGB(220, 0, 0)
        SaveSetting "WinMeters", "MEM", "Kernel", rgbMEMKernel
    End If
    log.WriteLog "Sub Main: Reading MEM indicator color for kernel usage - OK, value - " & rgbMEMKernel
    
    
    NetworkPresent = False
    log.WriteLog "Sub Main: Getting NETWORK information..."
    sNetworks = GetNetworkInterfaces()
    log.WriteLog "Sub Main: 'GetNetworkInterfaces' returned " & UBound(sNetworks) & " interfaces"
    
    If UBound(sNetworks) > 0 Then
        NetworkPresent = True
        log.WriteLog "Sub Main: NetworkPresent is  " & NetworkPresent
    End If
    
    'If IsIDE Then NetworkPresent = False
    
    intNet = GetSetting("WinMeters", "NET", "Interval", "0")
    If intNet = 0 Then
        intNet = 200
        SaveSetting "WinMeters", "NET", "Interval", "200"
    End If
    log.WriteLog "Sub Main: Reading NETWORK refresh interval - OK, value - " & intNet
    
    
    intHDD = GetSetting("WinMeters", "HDD", "Interval", "0")
    If intHDD = 0 Then
        intHDD = 200
        SaveSetting "WinMeters", "HDD", "Interval", "200"
    End If
    log.WriteLog "Sub Main: Reading HDD refresh interval - OK, value - " & intHDD
    
    ExtendedHDDInfo = GetSetting("WinMeters", "HDD", "ExtendedInfo", "False")
    SaveSetting "WinMeters", "HDD", "ExtendedInfo", ExtendedHDDInfo
    log.WriteLog "Sub Main: Reading HDD extended info in tooltip - OK, value - " & ExtendedHDDInfo
    
    'rgbdevRead = GetSetting("WinMeters", "HDD", "Read", "0")
    'If rgbdevRead = 0 Then
    '    rgbdevRead = RGB(0, 220, 0)
    '    SaveSetting "WinMeters", "HDD", "Read", rgbdevRead
    'End If
    'rgbdevWrite = GetSetting("WinMeters", "HDD", "Write", "0")
    'If rgbdevWrite = 0 Then
    '    rgbdevWrite = RGB(220, 0, 0)
    '    SaveSetting "WinMeters", "HDD", "Write", rgbdevWrite
    'End If
    
'RGB(0, 64, 200)
    'rgbTemp = GetSetting("WinMeters", "Temperature", "Color", "0")
    'If rgbTemp = 0 Then
    '    rgbTemp = RGB(0, 64, 200)
    '    SaveSetting "WinMeters", "Temperature", "Color", rgbTemp
    'End If

    ReDim vTrays(1 To nCPUs + 4)
    
'    For i = 1 To nCPUs + 2
'        Set vTrays(i) = New WinMetersTray
'        vTrays(i).currNumber = i
'    Next i
    
    log.WriteLog "Sub Main: Loading Settings window..."
    Load wmSettings
    
    wmSettings.checkShowSplash.Value = Abs(ShowSplashScreen)
    
    If NetworkPresent Then
        NetworkInterfaceIsMissing = True
        sActiveNetwork = GetSetting("WinMeters", "NET", "Interface", "0")
        log.WriteLog "Sub Main: Reading active network interface index - OK, value - " & sActiveNetwork
        If sActiveNetwork = "0" Then
            sActiveNetwork = sNetworks(1)
            SaveSetting "WinMeters", "NET", "Interface", sActiveNetwork
        End If
        For i = 1 To UBound(sNetworks)
            wmSettings.comboNet.AddItem sNetworks(i)
            If sActiveNetwork = sNetworks(i) Then NetworkInterfaceIsMissing = False
        Next i
        If NetworkInterfaceIsMissing Then sActiveNetwork = sNetworks(1)
        wmSettings.comboNet.Text = sActiveNetwork
    Else
        log.WriteLog "Sub Main: No working network interfaces found! Network indicator is disabled"
        wmSettings.comboNet.AddItem "No working network interfaces found!"
        wmSettings.comboNet.ListIndex = 0
        wmSettings.comboNet.Enabled = False
        wmSettings.slideNET.Enabled = False
        wmSettings.checkNET.Enabled = False
    End If
    
    
'    On Error Resume Next
'        modMain.KillProcessIcon "Core Temp"
'        Err.Clear
'        modMain.KillTaskByEXEName ("Core Temp.exe")
'        Sleep 1000
'        Err.Clear
'        Shell App.Path & "\Core Temp.exe", vbHide
'        modMain.KillProcessIcon "Core Temp"
'        Err.Clear
'        Sleep 2000
'        DoEvents
'        modMain.KillProcessIcon "Core Temp"
'        'Err.Clear
'        'modMain.SetCurrProcessVisibleInTaskList False, "Core Temp"
'    Err.Clear
'    On Error GoTo 0
    
    wmSettings.tmrCPUs.Interval = intCPUs
    wmSettings.checkCPU.Value = GetSetting("WinMeters", "CPU", "Enabled", "1")
    SaveSetting "WinMeters", "CPU", "Enabled", wmSettings.checkCPU.Value
    log.WriteLog "Sub Main: Reading CPU indicator visibility - OK, value - " & wmSettings.checkCPU.Value
    IndicatorEnabled(1) = -(wmSettings.checkCPU.Value)
    
    ShowDigitsInsteadThermometer = False
    'wmSettings.checkDigits.Value = GetSetting("WinMeters", "CPU", "Digits", "0")
    ShowDigitsInsteadThermometer = GetSetting("WinMeters", "CPU", "Digits", "False")
    wmSettings.checkDigits.Value = -(ShowDigitsInsteadThermometer)
    SaveSetting "WinMeters", "CPU", "Digits", ShowDigitsInsteadThermometer
    wmSettings.cmdSelectCPUFont.Enabled = ShowDigitsInsteadThermometer
    log.WriteLog "Sub Main: Reading CPU indicator textmodded - OK, value - " & ShowDigitsInsteadThermometer
    
    ShowOnlyTotalCPULoad = False
    'wmSettings.checkDigits.Value = GetSetting("WinMeters", "CPU", "Digits", "0")
    If nCPUs = 1 Then
        wmSettings.checkOneTotalCPU.Enabled = False
        wmSettings.checkOneTotalCPU.Visible = False
    Else
        ShowOnlyTotalCPULoad = GetSetting("WinMeters", "CPU", "TotalOnly", "False")
        SaveSetting "WinMeters", "CPU", "TotalOnly", ShowOnlyTotalCPULoad
        wmSettings.checkOneTotalCPU.Value = -(ShowOnlyTotalCPULoad)
        log.WriteLog "Sub Main: Reading CPU indicator totality - OK, value - " & ShowOnlyTotalCPULoad
    End If
    
    'Temperature
    'wmSettings.checkTemp.Value = 0 'GetSetting("WinMeters", "Temperature", "Enabled", "0")
    'SaveSetting "WinMeters", "Temperature", "Enabled", wmSettings.checkTemp.Value
    ShowSolidColors = False
    'wmSettings.chkSolidColor.Value = GetSetting("WinMeters", "CPU", "Solid", "0")
    ShowSolidColors = GetSetting("WinMeters", "CPU", "Solid", "False")
    wmSettings.chkSolidColor.Value = -(ShowSolidColors)
    SaveSetting "WinMeters", "CPU", "Solid", ShowSolidColors
    log.WriteLog "Sub Main: Reading CPU indicator solidity - OK, value - " & ShowSolidColors
    If nCPUs = 0 Then
        wmSettings.checkCPU.Value = 0
        wmSettings.checkCPU.Enabled = False
        wmSettings.frameCPU.Caption = "Error gathering CPU(s) info!"
    End If
    If -(wmSettings.checkCPU.Value) Then
'        If -(wmSettings.checkTemp.Value) Then
'            If (vTrays(nCPUs + 3) Is Nothing) Then
'                Set vTrays(nCPUs + 3) = New WinMetersTray
'                vTrays(nCPUs + 3).currNumber = nCPUs + 3
'                Load vTrays(nCPUs + 3)
'            End If
'        End If
        If ShowOnlyTotalCPULoad Then
            If (vTrays(1) Is Nothing) Then
                log.WriteLog "Sub Main: Loading one indicator for total CPU usage..."
                Set vTrays(1) = New WinMetersTray
                vTrays(1).currNumber = 1
                Load vTrays(1)
            End If
        Else
            For i = nCPUs To 1 Step -1
                If (vTrays(i) Is Nothing) Then
                    log.WriteLog "Sub Main: Loading indicator for CPU usage by core " & i & "..."
                    Set vTrays(i) = New WinMetersTray
                    vTrays(i).currNumber = i
                    Load vTrays(i)
                End If
            Next i
        End If
        'wmSettings.tmrCPUs.Enabled = True
    End If
    
    wmSettings.checkAntiAliasedMem = -(AntialiasedMEMIndicator)
    wmSettings.tmrMem.Interval = intMEM
    wmSettings.checkMEM.Value = GetSetting("WinMeters", "MEM", "Enabled", "1")
    SaveSetting "WinMeters", "MEM", "Enabled", wmSettings.checkMEM.Value
    log.WriteLog "Sub Main: Reading MEM indicator visibility - OK, value - " & wmSettings.checkMEM.Value
    IndicatorEnabled(2) = -(wmSettings.checkMEM.Value)
    
    wmSettings.chkShowKernelMem.Value = GetSetting("WinMeters", "MEM", "ShowKernel", "0")
    ShowKernelMemory = -(wmSettings.chkShowKernelMem.Value)
    log.WriteLog "Sub Main: Reading MEM indicator kernel setting - OK, value - " & ShowKernelMemory
    'SaveSetting "WinMeters", "MEM", "ShowKernel", wmSettings.chkShowKernelMem.Value
    If lMem = 0 Then
        wmSettings.checkMEM.Value = 0
        wmSettings.checkMEM.Enabled = False
        wmSettings.frameMEM.Caption = "Error gathering memory info!"
    End If
    If -(wmSettings.checkMEM.Value) Then
        If (vTrays(nCPUs + 1) Is Nothing) Then
            log.WriteLog "Sub Main: Loading indicator for MEMORY usage..."
            Set vTrays(nCPUs + 1) = New WinMetersTray
            vTrays(nCPUs + 1).currNumber = nCPUs + 1
            Load vTrays(nCPUs + 1)
        End If
        'wmSettings.tmrMem.Enabled = True
    End If
    
    wmSettings.tmrHDD.Interval = intHDD
    wmSettings.checkHDD.Value = GetSetting("WinMeters", "HDD", "Enabled", "1")
    SaveSetting "WinMeters", "HDD", "Enabled", wmSettings.checkHDD.Value
    log.WriteLog "Sub Main: Reading HDD indicator visibility - OK, value - " & wmSettings.checkHDD.Value
    wmSettings.checkExtendedHDDInfo.Value = -(ExtendedHDDInfo)
    IndicatorEnabled(3) = -(wmSettings.checkHDD.Value)
    If -(wmSettings.checkHDD.Value) Then
        If (vTrays(nCPUs + 2) Is Nothing) Then
            log.WriteLog "Sub Main: Loading indicator for HDD usage..."
            Set vTrays(nCPUs + 2) = New WinMetersTray
            vTrays(nCPUs + 2).currNumber = nCPUs + 2
            Load vTrays(nCPUs + 2)
        End If
        'wmSettings.tmrHDD.Enabled = True
    End If
        
    
    If NetworkPresent Then
        wmSettings.tmrNet.Interval = intNet
        wmSettings.checkNET.Value = GetSetting("WinMeters", "NET", "Enabled", "1")
        SaveSetting "WinMeters", "NET", "Enabled", wmSettings.checkNET.Value
        log.WriteLog "Sub Main: Reading NET indicator visibility - OK, value - " & wmSettings.checkNET.Value
        IndicatorEnabled(4) = -(wmSettings.checkNET.Value)
        If -(wmSettings.checkNET.Value) Then
            If (vTrays(nCPUs + 4) Is Nothing) Then
                log.WriteLog "Sub Main: Loading indicator for HDD usage..."
                Set vTrays(nCPUs + 4) = New WinMetersTray
                vTrays(nCPUs + 4).currNumber = nCPUs + 4
                Load vTrays(nCPUs + 4)
            End If
            'wmSettings.tmrNet.Enabled = True
        End If
    Else
        wmSettings.checkNET.Value = 0
        wmSettings.checkNET.Enabled = False
        wmSettings.frameNET.Caption = "Error gathering network info!"
    End If
    
    ShowAdvancedTooltips = True
    'wmSettings.checkShowTooltips.Value = GetSetting("WinMeters", "Trays", "Tooltips", "1")
    ShowAdvancedTooltips = GetSetting("WinMeters", "Trays", "Tooltips", "True")
    wmSettings.checkShowTooltips.Value = -(ShowAdvancedTooltips)
    SaveSetting "WinMeters", "Trays", "Tooltips", ShowAdvancedTooltips
    log.WriteLog "Sub Main: Reading advanced tooltips visibility - OK, value - " & ShowAdvancedTooltips
    
    If ShowSplashScreen Then
        log.WriteLog "Sub Main: Closing SplashScreen..."
        frmSplash.FadeOut
        Unload frmSplash
    End If
    
    IsFirstRun = GetSetting("WinMeters", "Trays", "FirstRun", "1")
    log.WriteLog "Sub Main: Reading info about first application start - OK, value - " & IsFirstRun
    If IsFirstRun = 1 Then
        IsFirstRun = 0
        log.WriteLog "Sub Main: Looks like it's first start on this computer, have to show Setting windows"
        wmSettings.Show
        SaveSetting "WinMeters", "Trays", "FirstRun", "0"
    End If
    
    log.WriteLog "Sub Main: All parameters is initialised, starting refresh timers..."
    
    If IndicatorEnabled(1) Then log.WriteLog "Sub Main: CPU refresh timer started"
    If IndicatorEnabled(2) Then log.WriteLog "Sub Main: MEM refresh timer started"
    If IndicatorEnabled(3) Then log.WriteLog "Sub Main: HDD refresh timer started"
    If IndicatorEnabled(4) Then log.WriteLog "Sub Main: NET refresh timer started"
    
    wmSettings.tmrCPUs.Enabled = IndicatorEnabled(1)
    wmSettings.tmrMem.Enabled = IndicatorEnabled(2)
    wmSettings.tmrHDD.Enabled = IndicatorEnabled(3)
    wmSettings.tmrNet.Enabled = IndicatorEnabled(4)
    
    log.WriteLog "Sub Main: WinMeters IS STARTED. WORKING..."
    log.CloseLog
    'MonitorRefresh
    'MonitorALL 500
    
    'End
End Sub

'Sub MonitorALL(Optional ByVal lTimeout As Long = 1000)
'    Dim strComputer As String, objWMIService As Object, objRefresher As Object, objProcessor As Object, intProcessorUse As Object, iUsage As Integer
'    Dim iCPU As Integer, objMemory As Object, iMemUsage As Long, intAvailableBytes As Object
'    strComputer = "."
'    Set objWMIService = GetObject("winmgmts:" _
'        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
'    Set objRefresher = CreateObject("WbemScripting.Swbemrefresher")
'    Set objProcessor = objRefresher.AddEnum _
'        (objWMIService, "Win32_PerfFormattedData_PerfOS_Processor").objectSet
'    Set objMemory = objRefresher.AddEnum _
'        (objWMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet
'
'    objRefresher.Refresh
'
'    Do
'        iCPU = 1
'        For Each intProcessorUse In objProcessor
'            If IsNull(intProcessorUse.PercentProcessorTime) Then
'                iUsage = 0
'            Else
'                iUsage = intProcessorUse.PercentProcessorTime
'            End If
'            If iCPU <= UBound(vTrays) - 2 Then vTrays(iCPU).DrawPercents iUsage
'            iCPU = iCPU + 1
'            DoEvents
'        Next
'
'        For Each intAvailableBytes In objMemory
'            If IsNull(intAvailableBytes.AvailableMBytes) Then
'                iMemUsage = 0
'            Else
'                iMemUsage = intAvailableBytes.AvailableMBytes
'            End If
'            vTrays(nCPUs + 1).DrawMem 100 - iMemUsage * 100 \ lMem
'            DoEvents
'        Next
'
'        Sleep lTimeout
'        DoEvents
'        objRefresher.Refresh
'    Loop
'End Sub

Private Function GetCPUsCount(Optional ByRef ExtraCPUInfo) As Integer
    Dim strComputer As String, objWMIService As Object, colSettings As Object, objComputer As Object, iRes As Integer, sRes As String, sRes2 As String
    Dim iCPUcounter As Integer, cntr As Integer
    On Error Resume Next
    'strComputer = "."
    'Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\CIMV2")
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    If Err Then
        log.WriteLog "Function GetCPUsCount: Connecting to WMI object... Failed!", Err.Number, Err.Description
        Err.Clear
    Else
        log.WriteLog "Function GetCPUsCount: Connecting to WMI object... Success!"
    End If
    'Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    'Win32_PerfFormattedData_PerfOS_Processor
    Set colSettings = objWMIService.ExecQuery("Select Name from Win32_PerfFormattedData_PerfOS_Processor Where Name<>'_Total'", , 48)
    If Err Then
        log.WriteLog "Function GetCPUsCount: Executing WMI query ''Select Name from Win32_PerfFormattedData_PerfOS_Processor Where Name<>'_Total'''... Failed!", Err.Number, Err.Description
        Err.Clear
    Else
        log.WriteLog "Function GetCPUsCount: Executing WMI query ''Select Name from Win32_PerfFormattedData_PerfOS_Processor Where Name<>'_Total'''... Success!"
    End If
    iCPUcounter = 0
    For Each objComputer In colSettings
        'iRes = objComputer.NumberOfLogicalProcessors
        'Err.Clear
        'If iRes = 0 Then iRes = objComputer.NumberOfProcessors
        If Len(objComputer.Name) > 0 Then iCPUcounter = iCPUcounter + 1: log.WriteLog "Function GetCPUsCount: Found processor/core " & objComputer.Name & "..."
    Next
    iRes = iCPUcounter
    log.WriteLog "Function GetCPUsCount: Total cores found: " & iCPUcounter
    ReDim strCPUsCoresNames(1 To iRes)
    ReDim strCPUsNames(0 To 0)
    ReDim CoresPerCPU(0 To 0)
    If Not IsMissing(ExtraCPUInfo) Then
        log.WriteLog "Function GetCPUsCount: Gettin extra CPUs info..."
        'ReDim ExtraCPUInfo(1 To nCPUs)
        Set colSettings = Nothing
        Set colSettings = CreateObject("WScript.Shell")
        If Err Then
            log.WriteLog "Function GetCPUsCount: CreateObject WScript.Shell for getting CPU info from registry... Failed!", Err.Number, Err.Description
            Err.Clear
        Else
            log.WriteLog "Function GetCPUsCount: CreateObject WScript.Shell for getting CPU info from registry... Success!"
        End If
        For cntr = 1 To iRes
            'sRes = colSettings.RegRead("HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\0\ProcessorNameString")
            strCPUsCoresNames(cntr) = colSettings.RegRead("HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\" & CStr(cntr - 1) & "\ProcessorNameString")
            log.WriteLog "Function GetCPUsCount: CPU" & cntr & " info from registry: " & strCPUsCoresNames(cntr)
        Next cntr
        Set colSettings = Nothing
        Err.Clear
        'strCPUsNames(0) = strCPUsCoresNames(1)
        Set colSettings = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
        If Err Then
            log.WriteLog "Function GetCPUsCount: Executing WMI query 'Select * from Win32_Processor'... Failed!", Err.Number, Err.Description
            Err.Clear
        Else
            log.WriteLog "Function GetCPUsCount: Executing WMI query 'Select * from Win32_Processor'... Success!"
        End If
        For Each objComputer In colSettings
            sRes2 = objComputer.DeviceID
            sRes2 = Replace$(sRes2, "CPU", vbNullString)
            If UBound(strCPUsNames) < Val(sRes2) Then ReDim strCPUsNames(0 To Val(sRes2))
        Next
        ReDim CoresPerCPU(0 To UBound(strCPUsNames))
        strCPUsNames(0) = strCPUsCoresNames(1)
        CoresPerCPU(0) = 0
        iCPUcounter = 1
        
        sRes = "CPU name: " & strCPUsCoresNames(1) & vbCrLf
        log.WriteLog "Function GetCPUsCount: Working CPU name: " & strCPUsCoresNames(1)
        'For cntr = 2 To nCPUs
        '    if
        'Next cntr
        '    sRes2 = objComputer.Name
        '    If Len(sRes) > 0 Or Len(sRes2) > 0 Then
        '        If sRes = sRes2 Then
        '            sRes = "CPU name: " & sRes & vbCrLf
        '        Else
        '            sRes = "CPU name (in registry): " & sRes & vbCrLf & _
        '                    "CPU name (via WMI): " & sRes2 & vbCrLf & _
        '                    "(maybe you have modern CPU and old system?)" & vbCrLf
        '        End If
        '    End If
        '    If Len(objComputer.CurrentClockSpeed) > 0 Then
        '        sRes = sRes & "Current processor speed: " & objComputer.CurrentClockSpeed & " MHz" & vbCrLf
        '        Exit For
        '    End If
        'Next
        ExtraCPUInfo = sRes
    End If
    Err.Clear
    On Error GoTo 0
    GetCPUsCount = iRes
End Function

Private Function GetMemTotal(Optional ByRef sExtraMemInfo As String) As Long ' mem in MB
    Dim strComputer As String, objWMIService As Object, colSettings As Object, objComputer As Object, lRes As Currency, sRes As String
    On Error Resume Next
    'strComputer = "."
'    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\CIMV2")
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    If Err Then
        log.WriteLog "Function GetMemTotal: Connecting to WMI object... Failed!", Err.Number, Err.Description
        Err.Clear
    Else
        log.WriteLog "Function GetMemTotal: Connecting to WMI object... Success!"
    End If
    
'    If Err Or (objWMIService Is Nothing) Then
'        MsgBox "Failed to create WMI object!", vbCritical, "DEBUG ERROR"
'        End
'    End If
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem", , 48)
    If Err Then
        log.WriteLog "Function GetMemTotal: Executing WMI query ''Select * from Win32_ComputerSystem''... Failed!", Err.Number, Err.Description
        Err.Clear
    Else
        log.WriteLog "Function GetMemTotal: Executing WMI query ''Select * from Win32_ComputerSystem''... Success!"
    End If
    
'    If Err Or (colSettings Is Nothing) Then
'        MsgBox "Failed to execute WMI query 'Select TotalPhysicalMemory from Win32_ComputerSystem'!", vbCritical, "DEBUG ERROR"
'        End
'    End If
    For Each objComputer In colSettings
'        MsgBox "Query result (TotalPhysicalMemory): " & objComputer.TotalPhysicalMemory & " bytes"
        log.WriteLog "Function GetMemTotal: Query result of Win32_ComputerSystem.TotalPhysicalMemory: " & objComputer.TotalPhysicalMemory & " bytes"
        lRes = Round(objComputer.TotalPhysicalMemory / 1024 / 1024)
        log.WriteLog "Function GetMemTotal: Working rounded value of available memory: " & lRes & " MB"
'        MsgBox "Working value must be " & objComputer.TotalPhysicalMemory / 1024 / 1024 & " MB" & vbCrLf & _
                    "Calculated value is " & lRes
    Next
    
    If Not IsEmpty(sExtraMemInfo) Then
        log.WriteLog "Function GetMemTotal: Gettin extra MEMORY info..."
        Set colSettings = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory", , 48)
        If Err Then
            log.WriteLog "Function GetMemTotal: Executing WMI query ''Select * from Win32_PhysicalMemory''... Failed!", Err.Number, Err.Description
            Err.Clear
        Else
            log.WriteLog "Function GetMemTotal: Executing WMI query ''Select * from Win32_PhysicalMemory''... Success!"
        End If
        
        For Each objComputer In colSettings
            log.WriteLog "Function GetMemTotal: Query result of Win32_PhysicalMemory.DeviceLocator: " & objComputer.DeviceLocator
            log.WriteLog "Function GetMemTotal: Query result of Win32_PhysicalMemory.Capacity: " & objComputer.Capacity
            log.WriteLog "Function GetMemTotal: Query result of Win32_PhysicalMemory.Speed: " & objComputer.Speed
            sRes = sRes & IIf(Len(objComputer.DeviceLocator) > 0, objComputer.DeviceLocator & ": ", vbNullString) & _
                    IIf(Len(objComputer.Capacity) > 0, Round(objComputer.Capacity / 1024 / 1024) & " MB", vbNullString) & _
                    IIf(Len(objComputer.Speed) > 0, " at speed " & objComputer.Speed & " MHz", vbNullString) & vbCrLf
        Next
        sExtraMemInfo = sRes
    End If
    
    Err.Clear
    On Error GoTo 0
    GetMemTotal = CLng(lRes)
End Function

Private Function GetNetworkInterfaces() As Variant
    Dim strComputer As String, objWMIService As Object, colSettings As Object, objComputer As Object ', lRes As Long, sRes As String
    Dim arrRes() As String
    On Error Resume Next
    'strComputer = "."
'    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\CIMV2")
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    If Err Then
        log.WriteLog "Function GetNetworkInterfaces: Connecting to WMI object... Failed!", Err.Number, Err.Description
        Err.Clear
    Else
        log.WriteLog "Function GetNetworkInterfaces: Connecting to WMI object... Success!"
    End If
    
    Set colSettings = objWMIService.ExecQuery("Select Name from Win32_PerfFormattedData_Tcpip_NetworkInterface where Name<>'MS TCP Loopback interface'", , 48)
    If Err Then
        log.WriteLog "Function GetNetworkInterfaces: Executing WMI query ''Select Name from Win32_PerfFormattedData_Tcpip_NetworkInterface where Name<>'MS TCP Loopback interface'''... Failed!", Err.Number, Err.Description
        Err.Clear
    Else
        log.WriteLog "Function GetNetworkInterfaces: Executing WMI query ''Select Name from Win32_PerfFormattedData_Tcpip_NetworkInterface where Name<>'MS TCP Loopback interface'''... Success!"
    End If
    
    ReDim arrRes(0 To 0)
    For Each objComputer In colSettings
        If objComputer.Name Like "isatap*" Then
            log.WriteLog "Function GetNetworkInterfaces: Found network name '" & objComputer.Name & "'... IGNORING, useless interface"
        Else
            ReDim Preserve arrRes(0 To UBound(arrRes) + 1)
            arrRes(UBound(arrRes)) = objComputer.Name
            log.WriteLog "Function GetNetworkInterfaces: Found network name '" & objComputer.Name & "'..."
        End If
    Next
    
    Err.Clear
    On Error GoTo 0
    GetNetworkInterfaces = arrRes
End Function


'Public Sub MonitorCPUs(Optional ByVal iCurrentCPU As Integer = 0)
'    Dim strComputer As String, objWMIService As Object, objProcessor As Object, intProcessorUse As Object, iUsage() As Integer
'    Dim iCPU As Integer, iUser() As Integer, n As Integer
'    On Error Resume Next
'    ReDim iUsage(1 To nCPUs)
'    ReDim iUser(1 To nCPUs)
'    strComputer = "."
'    Set objWMIService = GetObject("winmgmts:" _
'        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
'    Set objProcessor = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Processor Where Name<>'_Total'")
'    'Set objProcessor = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Processor")
'        iCPU = 1
'        For Each intProcessorUse In objProcessor
'            If IsNull(intProcessorUse.PercentProcessorTime) Then
'                iUsage(iCPU) = 0
'            Else
'                iUsage(iCPU) = intProcessorUse.PercentProcessorTime
'            End If
'
'            If IsNull(intProcessorUse.PercentUserTime) Then
'                iUser(iCPU) = -1
'            Else
'                iUser(iCPU) = intProcessorUse.PercentUserTime
'            End If
'            'If iCPU <= UBound(vTrays) - 2 Then vTrays(iCPU).DrawPercents iUsage, iUser
'            'vTrays(iCPU).DrawPercents iUsage, iUser
'            iCPU = iCPU + 1
'            DoEvents
'        Next
'    If iCurrentCPU > 0 Then
'        vTrays(iCurrentCPU).DrawPercents iUsage(iCurrentCPU), iUser(iCurrentCPU)
'    Else
'        For n = 1 To nCPUs
'            vTrays(n).DrawPercents iUsage(n), iUser(n)
'            DoEvents
'        Next n
'    End If
'    Err.Clear
'    On Error GoTo 0
'        'DoEvents
'End Sub

'Public Sub MonitorTempOLD()
'    Dim strComputer As String, objWMIService As Object, objProcessor As Object, intProcessorUse As Object, iUsage As Double
'    Dim iCPU As Integer
'    On Error Resume Next
'    strComputer = "."
'    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
'    Set objProcessor = objWMIService.ExecQuery("SELECT * FROM MSAcpi_ThermalZoneTemperature", , 48)
'        iCPU = 1: iUsage = 0
'        For Each intProcessorUse In objProcessor
''            If Not IsNull(intProcessorUse.CurrentTemperature) Then
'                iUsage = iUsage + intProcessorUse.CurrentTemperature
''            End If
'            'If iCPU >= nCPUs Then Exit For
'            iCPU = iCPU + 1
'            DoEvents
'        Next
'    iUsage = ((iUsage / (iCPU - 1)) - 2732) / 10#
'    If iUsage < 0 Then iUsage = 0
'    vTrays(nCPUs + 3).DrawTemp iUsage
'    Err.Clear
'    On Error GoTo 0
'        'DoEvents
'End Sub

'-----------------------------------------------
'Public Sub MonitorTemp()
'    Dim ttt As core_temp_shared_data, bRes As Boolean, deltaT As Single, m As Integer
'    On Error Resume Next
'    deltaT = 0
'    modMain.KillProcessIcon "Core Temp"
'    bRes = fnGetCoreTempInfoAlt(ttt)
'    If bRes Then
'        For m = 1 To nCPUs
'            deltaT = deltaT + ttt.fTemp(m)
'        Next m
'        deltaT = deltaT / nCPUs
'    End If
'    vTrays(nCPUs + 3).DrawTemp deltaT
'    DoEvents
'    Err.Clear
'    On Error GoTo 0
'End Sub
'------------------------------------------------

'Public Sub MonitorMEM()
'    Dim strComputer As String, objWMIService As Object
'    Dim objMemory As Object, iMemUsage As Long, intAvailableBytes As Object
'    On Error Resume Next
'    strComputer = "."
'    Set objWMIService = GetObject("winmgmts:" _
'        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
'
'    Set objMemory = objWMIService.ExecQuery("Select AvailableMBytes from Win32_PerfFormattedData_PerfOS_Memory")
'
'        For Each intAvailableBytes In objMemory
'            If IsNull(intAvailableBytes.AvailableMBytes) Then
'                iMemUsage = 0
'            Else
'                iMemUsage = intAvailableBytes.AvailableMBytes
'            End If
'            vTrays(nCPUs + 1).DrawMem 100 - iMemUsage * 100 \ lMem, iMemUsage
'            DoEvents
'        Next
'
'    Err.Clear
'    On Error GoTo 0
'        'DoEvents
'End Sub

'Public Sub MonitorHDD()
'    Dim strComputer As String, objWMIService As Object
'    Dim objHDD As Object, lUsage As Long, intIOTime As Object, bWrite As Boolean, bRead As Boolean
'    On Error Resume Next
'    strComputer = "."
'    bWrite = False: bRead = False
'
'    Set objWMIService = GetObject("winmgmts:" _
'        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
'
'    'Set objHDD = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_System")
'    Set objHDD = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfDisk_LogicalDisk Where Name='_Total'")
'    ''Set objHDD = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfDisk_PhysicalDisk")
'
'    For Each intIOTime In objHDD
'        'PercentDiskWriteTime
'        'FileReadBytesPersec
'        'DiskWriteBytesPersec
'        If IsNull(intIOTime.DiskWriteBytesPersec) Then
'            lUsage = 0
'        Else
'            lUsage = intIOTime.DiskWriteBytesPersec
'        End If
'        If lUsage > 0 Then bWrite = True
'
'        'PercentDiskReadTime
'        If IsNull(intIOTime.DiskReadBytesPersec) Then
'            lUsage = 0
'        Else
'            lUsage = intIOTime.DiskReadBytesPersec
'        End If
'        If lUsage > 0 Then bRead = True
'        DoEvents
'    Next
'
'    If Not bRead And Not bWrite Then
'        vTrays(nCPUs + 2).DrawHDD devIdle
'    ElseIf bRead And Not bWrite Then
'        vTrays(nCPUs + 2).DrawHDD devRead
'    ElseIf Not bRead And bWrite Then
'        vTrays(nCPUs + 2).DrawHDD devWrite
'    Else
'        vTrays(nCPUs + 2).DrawHDD devReadWrite
'    End If
'    DoEvents
'    Err.Clear
'    On Error GoTo 0
'End Sub

Public Sub ChangeCPUsInterval(ByVal NewInterval As Integer)
    If NewInterval < 200 Then NewInterval = 200
    If NewInterval > 1500 Then NewInterval = 1500
    intCPUs = NewInterval
    wmSettings.tmrCPUs.Interval = intCPUs
    SaveSetting "WinMeters", "CPU", "Interval", CStr(intCPUs)
    'wmSettings.tmrCPUs.StopTimer
    'DoEvents
    'wmSettings.tmrCPUs.StartTimer intCPUs
    DoEvents
End Sub

Public Sub ChangeMemInterval(ByVal NewInterval As Integer)
    If NewInterval < 500 Then NewInterval = 500
    If NewInterval > 3000 Then NewInterval = 3000
    intMEM = NewInterval
    wmSettings.tmrMem.Interval = intMEM
    SaveSetting "WinMeters", "MEM", "Interval", CStr(intMEM)
    'wmSettings.tmrMem.StopTimer
    'DoEvents
    'wmSettings.tmrMem.StartTimer intMEM
    DoEvents
End Sub

Public Sub ChangeHDDInterval(ByVal NewInterval As Integer)
    If NewInterval < 200 Then NewInterval = 200
    If NewInterval > 1000 Then NewInterval = 1000
    intHDD = NewInterval
    wmSettings.tmrHDD.Interval = intHDD
    SaveSetting "WinMeters", "HDD", "Interval", CStr(intHDD)
    'wmSettings.tmrHDD.StopTimer
    'DoEvents
    'wmSettings.tmrHDD.StartTimer intHDD
    DoEvents
End Sub

Public Sub ChangeNETInterval(ByVal NewInterval As Integer)
    If NewInterval < 200 Then NewInterval = 200
    If NewInterval > 1000 Then NewInterval = 1000
    intNet = NewInterval
    wmSettings.tmrNet.Interval = intNet
    SaveSetting "WinMeters", "NET", "Interval", CStr(intNet)
    DoEvents
End Sub


Public Function FindHWND(ByVal InitHWND As Long, ByVal PartialName As String) As Long
  Dim CurrWnd As Long, Len1 As Long, ListItem As String
  On Error Resume Next
  'ѕолучаем hWnd, который будет первым в списке
  'через него, мы сможем отыскать другие задачи
  CurrWnd = GetWindow(InitHWND, GW_HWNDFIRST)
  'ѕока возвращаемый hWnd имеет смысл, выполн€ем цикл
  Do While CurrWnd <> 0
    'ѕолучаем длину имени задани€ по CurrWnd
    Len1 = GetWindowTextLength(CurrWnd)
    'ѕолучить им€ задачи из списка
    ListItem = Space$(Len1 + 1)
    Len1 = GetWindowText(CurrWnd, ListItem, Len1 + 1)
    '≈сли получили им€ задачи, провер€ем на SS
    If Len1 > 0 And InStr(UCase$(ListItem), UCase$(PartialName)) > 0 Then
      FindHWND = CurrWnd
      Err.Clear
      On Error GoTo 0
      Exit Function
    End If
    'ѕереходим к следующей задаче из списка
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
    DoEvents
  Loop
  FindHWND = 0
  Err.Clear
  On Error GoTo 0
End Function

Public Sub KillTaskByEXEName(ByVal EXETaskName As String, Optional ByVal hwndIgnore As Long = 0)
    Dim hSnapShot As Long, nProcess As Long
    Dim uProcess As PROCESSENTRY32
    Dim hProcess As Long
  On Error Resume Next
    hSnapShot = CreateToolhelpSnapshot(2, 0)
    uProcess.dwSize = LenB(uProcess)
    nProcess = Process32First(hSnapShot, uProcess)
    Do While nProcess
      If InStr(UCase$(uProcess.szExeFile), UCase$(EXETaskName)) > 0 Then
        hProcess = OpenProcess(&H1F0FFF, 1, uProcess.th32ProcessID)
        If hwndIgnore <> hProcess Then TerminateProcess hProcess, 0
        Exit Do
      End If
      nProcess = Process32Next(hSnapShot, uProcess)
      DoEvents
    Loop
    CloseHandle hSnapShot
  Err.Clear
  On Error GoTo 0
End Sub


Public Sub KillProcessIcon(ByVal sPartialProcessName As String)
    Const NIM_DELETE = &H2
    Dim a As Long, mtypIcon As NOTIFYICONDATA
    On Error Resume Next
    a = FindHWND(wmSettings.hwnd, sPartialProcessName)
    If a <> 0 Then
        mtypIcon.hwnd = a
        Shell_NotifyIcon NIM_DELETE, mtypIcon
    End If
  Err.Clear
  On Error GoTo 0
End Sub

Public Sub SetCurrProcessVisibleInTaskList(Optional ByVal bVisible As Boolean = True, Optional ByVal sPartialExeName As String = vbNullString)
    Dim pid As Long
  On Error Resume Next
  If Len(sPartialExeName) > 0 Then
    pid = FindHWND(wmSettings.hwnd, sPartialExeName)
  Else
    pid = GetCurrentProcessId
  End If
  
  If bVisible Then
    RegisterServiceProcess pid, 0 'Show app
  Else
    RegisterServiceProcess pid, 1 'Hide app
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Public Sub SetOnTopWindow(hwnd As Long, OnTop As Boolean)

  On Error Resume Next
  If OnTop = True Then 'Make the window topmost
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
  Else
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Public Function GethWndTray(Optional ByVal bGetOnlyMainTrayHWND As Boolean = False) As Long
    Dim hwnd As Long, hWnd2 As Long 'handle
    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    If bGetOnlyMainTrayHWND Then
        GethWndTray = hwnd
        Exit Function
    End If
    hwnd = FindWindowEx(hwnd, ByVal 0&, "TrayNotifyWnd", vbNullString)
    hWnd2 = FindWindowEx(hwnd, ByVal 0&, "SysPager", vbNullString) 'uniquement XP
    If (hWnd2 = 0) Then hWnd2 = hwnd ' ME,2000
    hWnd2 = FindWindowEx(hWnd2, ByVal 0&, "ToolbarWindow32", vbNullString) ' ME, 2000, XP...
    If (hWnd2 = 0) Then
        GethWndTray = hwnd ' 95,98
    Else
        GethWndTray = hWnd2 ' ME, 2000, XP...
    End If
End Function

Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
        
    Dim oNewPic As Picture
    Dim tPicConv As PictDesc
    Dim IGuid As GUID
    
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .picType = vbPicTypeIcon
       .hImage = hIcon
    End With
    
    'Call CoCreateGuid(IGuid)
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    
    Set IconToPicture = oNewPic
    DestroyIcon hIcon
End Function

' Returns the version of Windows that the user is running
Public Function GetWindowsVersion(Optional ByRef snVersion As Single) As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"
                If Not IsMissing(snVersion) Then snVersion = CSng(CStr(osv.dwVerMajor) & GetDecimalSeparator() & CStr(osv.dwVerMinor))
                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                            Case 2
                                GetWindowsVersion = "Windows 8/Server 2012"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If
End Function

Public Function GetDecimalSeparator() As String
    Dim iLocale As Integer, sTmpStr As String, lRes As Long, aLen As Long
    On Error Resume Next
    sTmpStr = String$(255, " ") & Chr$(0)
    aLen = 1
    iLocale = GetUserDefaultLCID()
    lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, aLen)
    GetDecimalSeparator = Left$(sTmpStr, aLen)
    Err.Clear
    On Error GoTo 0
End Function

Public Function GetLocaleISOName() As String
    Dim iLocale As Integer, sTmpStr As String, lRes As Long, aLen As Long
    On Error Resume Next
    'sTmpStr = String$(255, " ") & Chr$(0)
    aLen = 1
    iLocale = GetUserDefaultLCID()
    aLen = GetLocaleInfo(iLocale, LOCALE_SISO639LANGNAME, ByVal 0&, 0)
    sTmpStr = String$(aLen - 1, " ") & Chr$(0)
    
    lRes = GetLocaleInfo(iLocale, LOCALE_SISO639LANGNAME, sTmpStr, aLen)
    GetLocaleISOName = Left$(sTmpStr, aLen - 1)
    Err.Clear
    On Error GoTo 0
End Function


Public Sub DisableCloseButton(ByVal FormHWND As Long)
  Dim hMenu As Long, Success As Long
  On Error Resume Next
  hMenu = GetSystemMenu(FormHWND, 0)
  Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
  SendMessage FormHWND, WM_NCACTIVATE, 0&, 0&
  SendMessage FormHWND, WM_NCACTIVATE, 1&, 0
  On Error GoTo 0
End Sub

Private Sub AdjustToken()

      '********************************************************************
      '* This procedure sets the proper privileges to allow a log off or a
      '* shut down to occur under Windows NT.
      '********************************************************************

         Const TOKEN_ADJUST_PRIVILEGES = &H20
         Const TOKEN_QUERY = &H8
         Const SE_PRIVILEGE_ENABLED = &H2

         Dim hdlProcessHandle As Long
         Dim hdlTokenHandle As Long
         Dim tmpLuid As LUID
         Dim tkp As TOKEN_PRIVILEGES
         Dim tkpNewButIgnored As TOKEN_PRIVILEGES
         Dim lBufferNeeded As Long
  
          On Error Resume Next
         'Set the error code of the last thread to zero using the
         'SetLast Error function. Do this so that the GetLastError
         'function does not return a value other than zero for no
         'apparent reason.
         SetLastError 0

         'Use the GetCurrentProcess function to set the hdlProcessHandle
         'variable.
         hdlProcessHandle = GetCurrentProcess()

         OpenProcessToken hdlProcessHandle, _
            (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle

'         'Get the LUID for shutdown privilege
'         LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

        LookupPrivilegeValue "", "SeCreatePermanentPrivilege", tmpLuid

         tkp.PrivilegeCount = 1    ' One privilege to set
         tkp.TheLuid = tmpLuid
         tkp.Attributes = SE_PRIVILEGE_ENABLED

         'Enable the shutdown privilege in the access token of this process
         AdjustTokenPrivileges hdlTokenHandle, _
                               False, _
                               tkp, _
                               Len(tkpNewButIgnored), _
                               tkpNewButIgnored, _
                               lBufferNeeded
      Err.Clear
      On Error Resume Next
End Sub

