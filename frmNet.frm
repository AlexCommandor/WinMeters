VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinMeters network monitor"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "frmNet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   20
      Max             =   100
      SelStart        =   100
      TickFrequency   =   10
      Value           =   100
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Opacity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sent (MB/sec):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   1980
      Width           =   5985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recieved (MB/sec):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sent (MB/sec):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1980
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Recieved (MB/sec):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "frmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_EX_LAYERED    As Long = &H80000
Private Const LWA_ALPHA        As Long = &H2&

Private WithEvents SysMon0 As VBControlExtender
Attribute SysMon0.VB_VarHelpID = -1
Private WithEvents SysMon1 As VBControlExtender
Attribute SysMon1.VB_VarHelpID = -1

Public SysMonIsRunning As Boolean

Public Sub CollectSamples()
    SysMon0.object.CollectSample
    SysMon1.object.CollectSample
End Sub

Private Sub ApplyTransparency(ByVal hwnd As Long, ByVal btPercentTrans As Integer)
' transparency (0 - 255)
    Dim lOldStyle  As Long
    Dim bTrans      As Byte

    bTrans = btPercentTrans * 2.55
    If btPercentTrans >= 99 Then bTrans = 255
    lOldStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, bTrans, LWA_ALPHA
End Sub


Private Sub Form_Load()
    Dim WS As Object, sSysMonString As String, vRes As Variant, i As Long, vResLang As Variant
    Dim oReg As Object, vRegRes As Variant, vRegKeys As Variant, sCounterString As String
    'Dim sResLang() As String
    Dim s1 As String, s2 As String, s3 As String, sTmp As String, lLbound As Long, lUbound As Long
    Const HKEY_LOCAL_MACHINE = &H80000002
    
    On Error Resume Next
    
    SysMonIsRunning = False
    
    Set WS = CreateObject("WScript.Shell")
    sSysMonString = WS.RegRead("HKCR\Sysmon\CurVer\")
    'Set WS = Nothing
    
    If Err Then
        Err.Clear
        SysMonIsRunning = False
        MsgBox "An error occured during creating SystemMonitor control!", vbCritical, "WinMeters network monitor"
        Me.Hide
        On Error GoTo 0
        Exit Sub
    End If
    
    Set SysMon0 = Controls.Add(sSysMonString, "SysMon0", Me)
    With SysMon0.object
        '.ManualUpdate = True
        .Reset
        .BackColor = &H0
        '.BackColorCtl = &HC8D0D4
        .BorderStyle = 0
        .Font.Name = "Tahoma"
        '.Font.Size = "6,75"
        .Font.Size = "6" & modMain.GetDecimalSeparator & "75"
        .Font.Italic = 0          'False
        .Font.UnderLine = 0       'False
        .Font.Strikethrough = 0    'False
        .Font.Weight = 400
        .GraphTitle = vbNullString
        .MaximumScale = 100
        .MinimumScale = 0
        .MonitorDuplicateInstances = False
        .ReadOnly = True
        .ShowHorizontalGrid = True
        .ShowVerticalGrid = True
        .ShowLegend = False
        .ShowToolbar = False
        .ShowValueBar = False
        .TimeBarColor = &HFF0000
        .GridColor = &H808080
        .DataSourceType = 1 'sysmonCurrentActivity
        
        ' Windows 7 specific parameters
        .ShowTimeAxisLabels = False
        .EnableTooltips = False
        .MaximumSamples = 100
        .DataPointCount = 100
        .ChartScroll = True
        
        Err.Clear
        .Counters.Add ("\Network Interface(" & modMain.sActiveNetwork & ")\Bytes Received/Sec")
        If Err Then ' here we can got an error on localized systems, so we must get right name from registry
            Err.Clear
            vRes = WS.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009\Counters")
            If Err Then
                Err.Clear
                vRes = WS.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009\Counter")
                If Not Err Then sCounterString = "Counter"
            Else
                sCounterString = "Counters"
            End If
            'vResLang = WS.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage\Counters")
            'Set WS = Nothing
            If Err Or (Not IsArray(vRes)) Then
                Err.Clear
                SysMonIsRunning = False
                MsgBox "An error occured during adding counters to SystemMonitor control!", vbCritical, "WinMeters network monitor"
                Me.Hide
                On Error GoTo 0
                Exit Sub
            End If
            
            Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
            oReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib", vRegRes
            Set oReg = Nothing
            lLbound = LBound(vRes): lUbound = UBound(vRes)
            
            For Each vRegKeys In vRegRes
                If vRegKeys <> "009" Then
                    vResLang = WS.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\" & vRegKeys & "\" & sCounterString)
                    For i = lLbound To lUbound
                        sTmp = vRes(i)
                        If (UCase$(sTmp) = "NETWORK INTERFACE") And (Len(s1) = 0) Then s1 = vResLang(i) ': Stop
                        If (UCase$(sTmp) = "BYTES RECEIVED/SEC") And (Len(s2) = 0) Then s2 = vResLang(i) ': Stop
                        If (UCase$(sTmp) = "BYTES SENT/SEC") And (Len(s3) = 0) Then s3 = vResLang(i) ': Stop
                    Next i
                    .Counters.Add ("\" & s1 & "(" & modMain.sActiveNetwork & ")\" & s2)
                    Erase vRes
                    Erase vResLang
                    If Not Err Then Exit For
                    Err.Clear
                End If
            Next
        End If
        
        If Err Then
            Err.Clear
            SysMonIsRunning = False
            MsgBox "An error occured during adding counters to SystemMonitor control!", vbCritical, "WinMeters network monitor"
            Me.Hide
            On Error GoTo 0
            Exit Sub
        End If

        .Counters(1).ScaleFactor = -6
        .Counters(1).color = RGB(0, 220, 0)
        .TimeBarColor = RGB(0, 0, 220)
        .UpdateInterval = 3600
        .ReportValueType = 4 'sysmonMaximum
        .ManualUpdate = False
    End With
    SysMon0.Move 0, 360, 8055, 1455
    SysMon0.Visible = True
    
    Set SysMon1 = Controls.Add(sSysMonString, "SysMon1", Me)
    With SysMon1.object
        '.ManualUpdate = True
        .Reset
        .BackColor = &H0
        '.BackColorCtl = &HC8D0D4
        .BorderStyle = 0
        .Font.Name = "Tahoma"
        .Font.Size = "6" & modMain.GetDecimalSeparator & "75"
        .Font.Italic = 0          'False
        .Font.UnderLine = 0       'False
        .Font.Strikethrough = 0    'False
        .Font.Weight = 400
        .GraphTitle = vbNullString
        .MaximumScale = 100
        .MinimumScale = 0
        .MonitorDuplicateInstances = False
        .ReadOnly = True
        .ShowHorizontalGrid = True
        .ShowVerticalGrid = True
        .ShowLegend = False
        .ShowToolbar = False
        .ShowValueBar = False
        .TimeBarColor = &HFF0000
        .GridColor = &H808080
        .DataSourceType = 1 'sysmonCurrentActivity
        
        ' Windows 7 specific parameters
        .ShowTimeAxisLabels = False
        .EnableTooltips = False
        .MaximumSamples = 100
        .DataPointCount = 100
        .ChartScroll = True
        
        Err.Clear
        .Counters.Add ("\Network Interface(" & modMain.sActiveNetwork & ")\Bytes Sent/Sec")
        If Err Then ' here we can got an error on localized systems, so we must get right name from registry
            Err.Clear
'            vRes = WS.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009\Counters")
'            vResLang = WS.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage\Counters")
'            If Err Then
'                Err.Clear
'                SysMonIsRunning = False
'                MsgBox "An error occured during adding counters to SystemMonitor control!", vbCritical, "WinMeters network monitor"
'                Me.Hide
'                On Error GoTo 0
'                Exit Sub
'            End If
'            If IsArray(vRes) Then
'                ReDim sRes(LBound(vRes) To UBound(vRes))
'                For i = LBound(sRes) To UBound(sRes)
'                    sRes(i) = vRes(i)
'                    If (UCase$(sRes(i)) = "NETWORK INTERFACE") And (Len(s1) = 0) Then s1 = vResLang(i): Stop
'                    If (UCase$(sRes(i)) = "BYTES RECEIVED/SEC") And (Len(s2) = 0) Then s2 = vResLang(i): Stop
'                    If (UCase$(sRes(i)) = "BYTES SENT/SEC") And (Len(s3) = 0) Then s3 = vResLang(i): Stop
'                Next i
                .Counters.Add ("\" & s1 & "(" & modMain.sActiveNetwork & ")\" & s3)
'            End If
        End If
        
        .Counters(1).ScaleFactor = -6
        .Counters(1).color = RGB(220, 0, 0)
        .TimeBarColor = RGB(0, 0, 220)
        .UpdateInterval = 3600
        .ReportValueType = 4 'sysmonMaximum
        .ManualUpdate = False
    End With
    SysMon1.Move 0, 2175, 8055, 1455
    SysMon1.Visible = True
    
    If Err Then
        Err.Clear
        SysMonIsRunning = False
        MsgBox "An error occured during adding SystemMonitor control!", vbCritical, "WinMeters network monitor"
        Me.Hide
        On Error GoTo 0
        Exit Sub
    End If
    
    
    SysMonIsRunning = True
    
    'modMain.DisableCloseButton Me.hwnd
    modMain.SetOnTopWindow Me.hwnd, True
    Me.Caption = modMain.sActiveNetwork
    
    On Error GoTo 0
    'Me.SysMon(0).ZOrder 0
    'Me.SysMon(1).ZOrder 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If (UnloadMode = vbFormControlMenu) Or (UnloadMode = vbFormCode) Then
        'Cancel = 1
        'Me.Hide
        For i = UBound(vTrays) To LBound(vTrays) Step -1
            If Not (vTrays(i) Is Nothing) Then
                vTrays(i).mnuPopUp(2).Checked = False
            End If
        Next i
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Controls.Remove SysMon0
    Controls.Remove SysMon1
    Set SysMon0 = Nothing
    Set SysMon1 = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Slider1_Click()
    ApplyTransparency Me.hwnd, Me.Slider1.Value
End Sub

'Private Sub SysMon222_OnSampleCollected(index As Integer)
'    Dim dMax As Double, dMin As Double, dAvg As Double, lStatus As Long, dCurrent As Double, lStat2 As Long
'    On Error Resume Next
'    Me.SysMon(index).Counters(1).GetStatistics dMax, dMin, dAvg, lStatus
'    Me.SysMon(index).Counters(1).GetValue dCurrent, lStat2
'    If lStatus = 0 Then
'        If CLng(dMax * (10 ^ Me.SysMon(index).Counters(1).ScaleFactor)) > 2 Then
'            Me.SysMon(index).MaximumScale = CLng(dMax * (10 ^ Me.SysMon(index).Counters(1).ScaleFactor))
'        Else
'            Me.SysMon(index).MaximumScale = 2
'        End If
'    End If
'    Me.Label1(index).Caption = "Current: " & Round(dCurrent * (10 ^ Me.SysMon(index).Counters(1).ScaleFactor)) & _
'            " MB/s   Min: " & Round(dMin * (10 ^ Me.SysMon(index).Counters(1).ScaleFactor)) & _
'            " MB/s   Max: " & Round(dMax * (10 ^ Me.SysMon(index).Counters(1).ScaleFactor)) & _
'            " MB/s   Average: " & Round(dAvg * (10 ^ Me.SysMon(index).Counters(1).ScaleFactor)) & " MB/s"
'    Me.SysMon(index).UpdateGraph
'    modMain.SetOnTopWindow Me.hwnd, True
'    Err.Clear
'    On Error GoTo 0
'End Sub

Private Sub OnSampleCollected(ByRef objSysMon As VBControlExtender, ByVal Index As Long)
    Dim dMax As Double, dMin As Double, dAvg As Double, lStatus As Long, dCurrent As Double, lStat2 As Long
    Dim lScaleFactor As Long, dScale As Double
    On Error Resume Next
    With objSysMon.object
        .Counters(1).GetStatistics dMax, dMin, dAvg, lStatus
        .Counters(1).GetValue dCurrent, lStat2
        lScaleFactor = .Counters(1).ScaleFactor
        dScale = 10 ^ lScaleFactor
        If lStatus = 0 Then
            If CLng(dMax * dScale) > 1 Then
                .MaximumScale = CLng(dMax * dScale)
            Else
                .MaximumScale = 1
            End If
        End If
        Me.Label1(Index).Caption = "Current: " & Round(dCurrent * dScale) & _
            " MB/s   Min: " & Round(dMin * dScale) & _
            " MB/s   Max: " & Round(dMax * dScale) & _
            " MB/s   Average: " & Round(dAvg * dScale) & " MB/s"
        .UpdateGraph
    End With
    'modMain.SetOnTopWindow Me.hwnd, True
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SysMon0_ObjectEvent(Info As EventInfo)
    If Not SysMonIsRunning Then Exit Sub
   If Info.Name = "OnSampleCollected" Then
        OnSampleCollected SysMon0, 0
   End If
End Sub

Private Sub SysMon1_ObjectEvent(Info As EventInfo)
    If Not SysMonIsRunning Then Exit Sub
   If Info.Name = "OnSampleCollected" Then
        OnSampleCollected SysMon1, 1
   End If
End Sub

