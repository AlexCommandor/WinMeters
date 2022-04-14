VERSION 5.00
Begin VB.Form WinMetersTray 
   Appearance      =   0  'Flat
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6030
   Icon            =   "WinMetersSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp 
         Caption         =   "Settings"
         Index           =   0
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "Show Network monitor"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "Exit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "WinMetersTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

Public currNumber As Integer
Public currGUIDstr As String
Private currGUIDguid As GUID

'Private hwIcon(0 To 3) As Long, netIcon(0 To 3) As Long, lIconIndex As Integer, memIcon As Long
''Icon index is: 0 - 128x128, 1 - 64x64, 2 - 48x48, 3 - 32x32, 4 - 28x28, 5 - 24x24, 6 - 20x20, 7 - 16x16

Private objDraw As New LineGS

Private objWMIServiceForCPU As Object
Private objWMIServiceForMEM As Object
Private objWMIServiceForHDD As Object
'Private objWMIServiceForHDD2 As Object
Private objWMIServiceForNET As Object
Private objRefresherCPU As Object
Private objRefresherMEM As Object
Private objRefresherHDD As Object
Private objFSO As Object
'Private objRefresherHDD2 As Object
Private objRefresherNET As Object
Private objParentCPU As Object
Private objParentMEM As Object
Private objParentHDD As Object
Private objParentNET As Object
Private objHDD As Object ', objChild As Object
'Private objParent2 As Object, objParent3 As Object

Private dat1 As Date, dat2 As Date
Private nDisksPrev As Long, nDisksCurr As Long

Private sCaption As String, sBaloon As String
Public lMeterData1 As Long, lMeterData2 As Long
Public lPictWidth As Long, lPictHeight As Long

'Private Enum eDiskData
'    Partition = 1
'    TotalSize = 2
'    FreeSpace = 3
'    VolumeName = 4
'End Enum




Public Sub DrawPercents(ByVal currPercents As Integer, Optional ByVal iKernelTime As Integer = 0)
    Dim currY As Integer, currU As Integer, sCoreString As String
    Dim lUptime As Long, sUptime As String
    Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long
    On Error Resume Next
    
    lUptime = GetTickCount()
    lUptime = lUptime \ 1000
    Days = lUptime \ (24& * 3600&)
    If Days > 0 Then lUptime = lUptime - (24& * 3600& * Days)
    Hours = lUptime \ 3600&
    If Hours > 0 Then lUptime = lUptime - (3600& * Hours)
    Minutes = lUptime \ 60
    Seconds = lUptime Mod 60
    sUptime = Days & " days, " & Format$(Hours, "00") & ":" & Format$(Minutes, "00") & ":" & Format$(Seconds, "00")
    
  If Not (lMeterData1 = currPercents And lMeterData2 = iKernelTime) Then
       If modMain.ShowDigitsInsteadThermometer Then
            DrawTextPercents currPercents, modMain.rgbCPUsUser ', iKernelTime
       Else
            currY = (Me.lPictWidth - 2) * currPercents / 100
            If modMain.ShowSolidColors Then
                '-------------------------
                Me.pic.Line (Me.lPictWidth / 4, Me.lPictHeight - 2)-((3 * Me.lPictWidth / 4) - 2, Me.lPictHeight - 1 - currY), modMain.rgbCPUsUser, BF
                '-------------------------
            Else
                DoGradient Me.pic, modMain.rgbCPUsUser2, modMain.rgbCPUsUser, _
                    Me.lPictWidth / 4, Me.lPictHeight - 1 - currY, (3 * Me.lPictWidth / 4) - 2, Me.lPictHeight - 2, gradHorizontal
            End If
            If iKernelTime > 0 Then
                currU = (Me.lPictWidth - 2) * iKernelTime / 100
                If modMain.ShowSolidColors Then
                    '-------------------------
                    Me.pic.Line (Me.lPictWidth / 4, Me.lPictHeight - 2)-((3 * Me.lPictWidth / 4) - 2, Me.lPictHeight - 1 - currU), modMain.rgbCPUsKernel, BF
                    '-------------------------
                Else
                    DoGradient Me.pic, modMain.rgbCPUsKernel2, modMain.rgbCPUsKernel, _
                         Me.lPictWidth / 4, Me.lPictHeight - 1 - currU, (3 * Me.lPictWidth / 4) - 2, Me.lPictHeight - 2, gradHorizontal
                End If
            End If
            Me.pic.Line (Me.lPictWidth / 4, Me.lPictHeight - currY - 2)-((3 * Me.lPictWidth / 4) - 2, 1), &HFF00FF, BF
            Me.pic.Line (Me.lPictWidth / 4 - 1, 0)-((3 * Me.lPictWidth / 4) - 1, Me.lPictHeight - 1), vbBlack, B
            SysTray.DrawIcon Me.pic, &HFF00FF
        End If
  End If
  If modMain.ShowOnlyTotalCPULoad Then sCoreString = "Total cores: " Else sCoreString = "Core #" & currNumber & ": "
  
  sCaption = "System uptime: " & sUptime & vbCrLf & sCoreString & vbCrLf & "usage " & currPercents & _
                    "%, kernel time " & iKernelTime & "%"
  sBaloon = IIf(Len(modMain.sExtraCPUsInfo) > 0, modMain.sExtraCPUsInfo & vbCrLf, vbNullString) & _
                 "System uptime: " & sUptime & vbCrLf & vbCrLf & sCoreString & vbCrLf & "usage " & currPercents & _
                    "%, kernel time " & iKernelTime & "%"
    'Me.Caption = sCaption
    'SysTray.TooltipText = sCaption
    'If (frmTooltip.currentTipTray = Me.currNumber) And (frmTooltip.lblTooltipText <> sBaloon) Then frmTooltip.lblTooltipText = sBaloon
    lMeterData1 = currPercents: lMeterData2 = iKernelTime
    If frmTooltip.currentTipTray = Me.currNumber And frmTooltip.Timer1.Enabled Then _
                frmTooltip.lblTooltipText.Caption = sBaloon
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub DrawMem(ByVal currPercents As Integer, Optional ByVal currMemInMbytes As Long = 0, Optional ByVal currKernelMemory As Long = 0)
    On Error Resume Next
    If memIcon <> 0 Then DestroyIcon memIcon
    memIcon = 0
  If Not (lMeterData1 = currPercents And lMeterData2 = CLng(currKernelMemory * 100 / lMem)) Then
    memIcon = MyDrawCircle(currPercents, currKernelMemory * 100 / lMem)
    If memIcon = 0 Or (Not modMain.AntialiasedMEMIndicator) Then
        MyDrawCircle_OLD currPercents, currKernelMemory * 100 / lMem
        Me.pic.Refresh
        SysTray.DrawIcon Me.pic, &HFF00FF
    Else
        SysTray.DrawIcon Me.pic, , memIcon
        DestroyIcon memIcon
    End If
  End If
    If currMemInMbytes > 0 Then
        sCaption = "Memory usage: " & vbCrLf & modMain.lMem - currMemInMbytes & " MB (" & currPercents & "%) used," & vbCrLf & _
                        currMemInMbytes & " MB (" & 100 - currPercents & "%) free," & vbCrLf & _
                        currKernelMemory & " MB kernel memory," & vbCrLf & "total memory " & modMain.lMem & " MB"
        sBaloon = IIf(Len(modMain.sExtraMemoryInfo) > 0, modMain.sExtraMemoryInfo & vbCrLf, vbNullString) & _
                    "Memory usage: " & vbCrLf & modMain.lMem - currMemInMbytes & " MB (" & currPercents & "%) used," & vbCrLf & _
                        currMemInMbytes & " MB (" & 100 - currPercents & "%) free," & vbCrLf & _
                        currKernelMemory & " MB kernel memory," & vbCrLf & "total memory " & modMain.lMem & " MB"
    Else
        sCaption = "Memory usage: " & currPercents & "% used from " & modMain.lMem & " MB"
        sBaloon = IIf(Len(modMain.sExtraMemoryInfo) > 0, modMain.sExtraMemoryInfo & vbCrLf, vbNullString) & _
                    "Memory usage: " & currPercents & "% used from " & modMain.lMem & " MB"
    End If
    'SysTray.DrawIcon Me.pic, &HFF00FF
    'If (frmTooltip.currentTipTray = Me.currNumber) And (frmTooltip.lblTooltipText <> sBaloon) Then frmTooltip.lblTooltipText = sBaloon
    'SysTray.TooltipText = sCaption
    lMeterData1 = currPercents: lMeterData2 = currKernelMemory * 100 / lMem
    If frmTooltip.currentTipTray = Me.currNumber And frmTooltip.Timer1.Enabled Then _
                frmTooltip.lblTooltipText.Caption = sBaloon
    Err.Clear
    On Error GoTo 0
End Sub

'Public Sub DrawTemp(ByVal dTemp As Double)
'    Dim rct As RECT, sTemp As String
'    On Error Resume Next
'    Me.pic.Cls
'
'    With rct
'        .Left = 0
'        .Right = Me.pic.ScaleWidth
'        .Top = 0
'        .Bottom = Me.pic.ScaleHeight
'    End With
'    Me.pic.FontSize = 9
'    Me.pic.FontBold = False
'    Me.pic.ForeColor = modMain.rgbTemp
'    sTemp = Format$(dTemp, "0°")
'    DrawText Me.pic.hDC, sTemp, -1, rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
'    Me.pic.Refresh
'    sCaption = "Average CPUs temp: " & sTemp & "C"
'    sBaloon = "Average CPUs temp: " & sTemp & "C"
'    'If (frmTooltip.currentTipTray = Me.currNumber) And (frmTooltip.lblTooltipText <> sBaloon) Then frmTooltip.lblTooltipText = sBaloon
'    SysTray.DrawIcon Me.pic, &HFF00FF
'    'SysTray.TooltipText = sCaption
'    Err.Clear
'    On Error GoTo 0
'End Sub

Public Sub DrawTextPercents(ByVal iUsage As Integer, cColor As OLE_COLOR)
    Dim rct As RECT, sTemp As String
    On Error Resume Next
    Me.pic.Cls
    With rct
        .Left = 0
        .Right = Me.lPictWidth
        .Top = 0
        .Bottom = Me.lPictHeight
    End With
    Me.pic.FontName = modMain.FontNameCPU
    Me.pic.FontSize = modMain.FontSizeCPU
    Me.pic.FontBold = modMain.FontBoldCPU
    Me.pic.FontItalic = modMain.FontItalicCPU
    Me.pic.ForeColor = cColor
    sTemp = CStr(iUsage) ' & "%" 'Format$(dNumbers, "0%")
    'sTemp = CStr(iUsage) & "%"  'Format$(dNumbers, "0%")
    'DrawText Me.pic.hDC, sTemp, -1, rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    'DrawText Me.pic.hDC, sTemp, Len(sTemp), rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    DrawText Me.pic.hDC, sTemp, Len(sTemp), rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP
    
    'Me.pic.ForeColor = modMain.rgbCPUsKernel
    'sTemp = CStr(iKernel) ' & "%" 'Format$(dNumbers, "0%")
    'DrawText Me.pic.hDC, sTemp, Len(sTemp), rct, DT_CENTER Or DT_BOTTOM Or DT_SINGLELINE
    
    Me.pic.Refresh
    'sCaption = "Average CPUs temp: " & sTemp & "C"
    'sBaloon = "Average CPUs temp: " & sTemp & "C"
   'If (frmTooltip.currentTipTray = Me.currNumber) And (frmTooltip.lblTooltipText <> sBaloon) Then frmTooltip.lblTooltipText = sBaloon
    SysTray.DrawIcon Me.pic, &HFF00FF
    'SysTray.TooltipText = sCaption
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub DrawHDD_OLD(ByVal hddStatus As tDeviceStatus)
    On Error Resume Next
    Me.pic.Cls
    'Me.pic.FillStyle = vbSolid
    'Me.pic.FillColor = vbBlack
    
    'Me.pic.Line (7, 15)-(22, 16), vbBlack, BF
    'If hddStatus = devIdle Then
        Me.pic.Line (Me.lPictWidth / 4 - 1, Me.lPictHeight / 2 - 2)-((3 * Me.lPictWidth / 4) + 1, Me.lPictHeight / 2 + 2), vbBlack, B
    'Else
        'Me.pic.Line (Me.pic.Width \ 4 - 1, Me.pic.Height \ 2 - 2)-((3 * Me.pic.Width \ 4) + 1, Me.pic.Height \ 2 + 2), RGB(0, 64, 200), BF
        'Me.pic.Line (Me.pic.Width \ 4 - 1, Me.pic.Height \ 2 - 2)-((3 * Me.pic.Width \ 4) + 1, Me.pic.Height \ 2 + 2), vbYellow, BF
        'Me.pic.Line (Me.pic.Width \ 4 - 1, Me.pic.Height \ 2 - 2)-((3 * Me.pic.Width \ 4) + 1, Me.pic.Height \ 2 + 2), vbBlack, B
    'End If
    'Me.pic.Line (Me.pic.Width \ 2 + 1, Me.pic.Height \ 2 - 1)-((3 * Me.pic.Width \ 4), Me.pic.Height \ 2 + 1), RGB(200, 200, 0), BF
    
    If hddStatus = devRead Or hddStatus = devReadWrite Then ' drawing read indicator
        'Me.pic.Line (7, 0)-(22, 14), RGB(0, 196, 0), BF
        Me.pic.Line (Me.lPictWidth / 4 - 1, 0)-((3 * Me.lPictWidth / 4) + 1, Me.lPictHeight / 2 - 3), modMain.rgbdevRead, BF
    End If
    If hddStatus = devWrite Or hddStatus = devReadWrite Then ' drawing read indicator
        'Me.pic.Line (7, 17)-(22, 31), vbRed, BF
        Me.pic.Line (Me.lPictWidth / 4 - 1, Me.lPictHeight / 2 + 3)-((3 * Me.lPictWidth / 4) + 1, Me.lPictHeight - 1), modMain.rgbdevWrite, BF
    End If
    Me.pic.Refresh
    sCaption = "HDD"
    'SysTray.DrawIcon Me.picHDD(hddStatus), &HFF00FF
    SysTray.DrawIcon Me.pic, &HFF00FF
    'SysTray.TooltipText = sCaption
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub DrawHDD(ByVal hddStatus As tDeviceStatus)
    On Error Resume Next
    If hwIcon(hddStatus) = 0 Then
        Me.pic.Cls
        Me.pic.Line (Me.lPictWidth / 4 - 1, Me.lPictHeight / 2 - 2)-((3 * Me.lPictWidth / 4) + 1, Me.lPictHeight / 2 + 2), vbBlack, B
        If hddStatus = devRead Or hddStatus = devReadWrite Then ' drawing read indicator
            Me.pic.Line (Me.lPictWidth / 4 - 1, 0)-((3 * Me.lPictWidth / 4) + 1, Me.lPictHeight / 2 - 3), modMain.rgbdevRead, BF
        End If
        If hddStatus = devWrite Or hddStatus = devReadWrite Then ' drawing write indicator
            Me.pic.Line (Me.lPictWidth / 4 - 1, Me.lPictHeight / 2 + 3)-((3 * Me.lPictWidth / 4) + 1, Me.lPictHeight - 1), modMain.rgbdevWrite, BF
        End If
        Me.pic.Refresh
        SysTray.DrawIcon Me.pic, &HFF00FF
    Else
'        'Me.Icon = hwIcon(hddStatus)
        SysTray.DrawIcon Me.pic, , hwIcon(hddStatus)
    End If
    'If (frmTooltip.currentTipTray = Me.currNumber) And (frmTooltip.lblTooltipText <> sBaloon) Then frmTooltip.lblTooltipText = sBaloon
    'SysTray.TooltipText = sCaption
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub DrawNET(ByVal netStatus As tDeviceStatus)
    sCaption = "Monitoring network interface:" & vbCrLf & modMain.sActiveNetwork
    sBaloon = sCaption
    SysTray.DrawIcon Me.pic, , netIcon(netStatus)
    'If (frmTooltip.currentTipTray = Me.currNumber) And (frmTooltip.lblTooltipText <> sBaloon) Then frmTooltip.lblTooltipText = sBaloon
    'SysTray.TooltipText = sCaption
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub Form_Load()
    Dim tmpWidth As Long, tmpHeight As Long
   'iconMinWidth = GetSystemMetrics(SM_CXSMICON)
   'iconMinHeight = GetSystemMetrics(SM_CYSMICON)
    'Me.pic.Width = modMain.iconMinWidth
    'Me.pic.Height = modMain.iconMinHeight
    
    'Call GetWindowsVersion(snglWinVer)
    'If snglWinVer < 6 Then ' pre-Vista windows - use small icons
    log.WriteLog "Loading tray" & Me.currNumber & "..."
    
    Me.mnuPopUp(2).Enabled = modMain.NetworkPresent
    Me.mnuPopUp(2).Checked = False
    
    Me.pic.ScaleMode = vbPixels
    Me.WindowState = vbMinimized
    DoEvents
    Me.Hide
    Set SysTray = New clsSysTray
        
        Me.pic.Width = GetSystemMetrics(SM_CXSMICON)
        log.WriteLog "Loading tray" & Me.currNumber & ": GetSystemMetrics returns icon size " & Me.pic.Width & "x" & Me.pic.Width & " pixels"
        Me.pic.Height = Me.pic.Width 'GetSystemMetrics(SM_CYSMICON)
        'tmpWidth = SysTray.GetSysTrayIconWidth(tmpHeight)
        'If tmpWidth > Me.pic.Width Then Me.pic.Width = tmpWidth: Me.pic.Height = Me.pic.Width
        'If tmpHeight > Me.pic.Height Then Me.pic.Height = tmpHeight
        
        'MsgBox Me.pic.ScaleWidth
        
        If Me.pic.Width >= 16 And Me.pic.Width < 20 Then
            lIconIndex = 7
            Me.pic.Width = 16: Me.pic.Height = 16
        ElseIf Me.pic.Width >= 20 And Me.pic.Width < 24 Then
            lIconIndex = 6
            Me.pic.Width = 20: Me.pic.Height = 20
        ElseIf Me.pic.Width >= 24 And Me.pic.Width < 28 Then
            lIconIndex = 5
            Me.pic.Width = 24: Me.pic.Height = 24
        ElseIf Me.pic.Width >= 28 And Me.pic.Width < 32 Then
            lIconIndex = 4
            Me.pic.Width = 28: Me.pic.Height = 28
        ElseIf Me.pic.Width >= 32 And Me.pic.Width < 48 Then
            lIconIndex = 3
            Me.pic.Width = 32: Me.pic.Height = 32
        ElseIf Me.pic.Width >= 48 And Me.pic.Width < 64 Then
            lIconIndex = 2
            Me.pic.Width = 48: Me.pic.Height = 48
        ElseIf Me.pic.Width >= 64 And Me.pic.Width < 128 Then
            lIconIndex = 1
            Me.pic.Width = 64: Me.pic.Height = 64
        ElseIf Me.pic.Width = 128 Then
            lIconIndex = 0
            Me.pic.Width = 128: Me.pic.Height = 128
        Else
            lIconIndex = 3
            Me.pic.Width = 32: Me.pic.Height = 32
        End If
        
        log.WriteLog "Loading tray" & Me.currNumber & ": Selected standard icon size " & Me.pic.Width & "x" & Me.pic.Width & " pixels"
        'MsgBox Me.pic.Width & " x " & Me.pic.Height
    'Else 'Vista and above - use large icons
    '    Me.pic.Width = GetSystemMetrics(SM_CXICON)
    '    Me.pic.Height = GetSystemMetrics(SM_CYICON)
    'End If
    Me.pic.ScaleMode = vbPixels
   
    Me.lPictWidth = Me.pic.ScaleWidth: Me.lPictHeight = Me.pic.ScaleHeight
    'for win7 compatibility we have to generate unique GUID for each icon in tray
    'and save it into registry to avoid growing Notifycation Tray icon cache
    currGUIDstr = GetSetting("WinMeters", "Trays", App.Path & "\" & App.EXEName & ".exe," & Me.currNumber, "0")
    If Not GetGUIDfromString(currGUIDstr, currGUIDguid) Then ' some error in previously saved GUID or first app start
        Call CoCreateGuid(currGUIDguid)
        currGUIDstr = GetStringFromGUID(currGUIDguid)
        If Len(currGUIDstr) = 0 Then
            MsgBox "Unexpected error occured during getting GUID for app, exiting.", vbCritical, "WinMeters critical error"
            End
        End If
        SaveSetting "WinMeters", "Trays", App.Path & "\" & App.EXEName & ".exe," & Me.currNumber, currGUIDstr
    End If
    log.WriteLog "Loading tray" & Me.currNumber & ": Current GUID for icon tray - " & currGUIDstr
    
    dat1 = Now()
    If Me.currNumber <= modMain.nCPUs Then ' all cpu cores
        Call CreateCPURefresher
        log.WriteLog "Loading tray" & Me.currNumber & ": Creating CPU Refresher... Success!"
        Me.pic.Line (Me.pic.ScaleWidth / 4 - 1, 0)-((3 * Me.pic.ScaleWidth / 4) - 1, Me.pic.ScaleHeight - 1), vbBlack, B
        Me.pic.Refresh
        Me.Caption = ""
        sCaption = "Core #" & currNumber
        sBaloon = "Core #" & currNumber
        SysTray.Init Me, "Core #" & currNumber, currGUIDstr
        SysTray.DrawIcon Me.pic, &HFF00FF
        log.WriteLog "Loading tray" & Me.currNumber & ": Tray loaded!"
    ElseIf Me.currNumber = modMain.nCPUs + 1 Then 'memory
        Set objDraw = New LineGS
        objDraw.DIB Me.pic.hDC, Me.pic.Image.handle, Me.pic.ScaleWidth, Me.pic.ScaleHeight
        Call CreateMEMRefresher
        log.WriteLog "Loading tray" & Me.currNumber & ": Creating MEM Refresher... Success!"
        memIcon = MyDrawCircle(0)
        'Me.pic.Refresh
        Me.Caption = ""
        sCaption = "Memory"
        sBaloon = "Memory"
        SysTray.Init Me, "Memory", currGUIDstr
        If memIcon = 0 Or (Not modMain.AntialiasedMEMIndicator) Then
            SysTray.DrawIcon Me.pic, &HFF00FF
        Else
            SysTray.DrawIcon Me.pic, , memIcon
        End If
        log.WriteLog "Loading tray" & Me.currNumber & ": Tray loaded!"
    ElseIf Me.currNumber = modMain.nCPUs + 2 Then 'HDD
        Call CreateHDDRefresher
        log.WriteLog "Loading tray" & Me.currNumber & ": Creating HDD Refresher... Success!"
        'Set objHDD = objWMIServiceForHDD.ExecQuery("Select * from Win32_DiskDrive", , 16)
        If objHDD Is Nothing Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Executing WMI query 'Select * from Win32_DiskDrive'... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Executing WMI query 'Select * from Win32_DiskDrive'... Success!"
        End If
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        If objFSO Is Nothing Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Creating FileSystemObject... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Creating FileSystemObject... Success!"
        End If
        
        ''Me.pic.Line (7, 15)-(22, 16), vbBlack, BF
        'Me.pic.Line (Me.pic.Width \ 4 - 1, Me.pic.Height \ 2 - 1)-((3 * Me.pic.Width \ 4) + 1, Me.pic.Height \ 2 + 1), vbBlack, BF
        'Me.pic.Refresh
        hwIcon(devIdle) = LoadIconFromMultiRES("HDDICON", 101, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If hwIcon(devIdle) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD Idle status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD Idle status... Success!"
        End If
        
        hwIcon(devRead) = LoadIconFromMultiRES("HDDICON", 102, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If hwIcon(devRead) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD Read status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD Read status... Success!"
        End If
        
        hwIcon(devWrite) = LoadIconFromMultiRES("HDDICON", 103, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If hwIcon(devWrite) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD Write status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD Write status... Success!"
        End If
        
        hwIcon(devReadWrite) = LoadIconFromMultiRES("HDDICON", 104, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If hwIcon(devReadWrite) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD ReadWrite status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for HDD ReadWrite status... Success!"
        End If
        
        Me.Caption = ""
        sCaption = "HDD"
        sBaloon = "HDD"
        SysTray.Init Me, "HDD", currGUIDstr
        If hwIcon(devIdle) = 0 Then
            Me.pic.Line (Me.pic.ScaleWidth / 4 - 1, Me.pic.ScaleHeight / 2 - 1)-((3 * Me.pic.ScaleWidth / 4) + 1, Me.pic.ScaleHeight / 2 + 1), vbBlack, BF
            Me.pic.Refresh
            SysTray.DrawIcon Me.pic, &HFF00FF
        Else
            'Me.Icon = hwIcon(devIdle)
            SysTray.DrawIcon Me.pic, , hwIcon(devIdle)
        End If
        log.WriteLog "Loading tray" & Me.currNumber & ": Tray loaded!"
    ElseIf Me.currNumber = modMain.nCPUs + 3 Then 'temp
        Me.pic.Cls
        Me.Caption = ""
        sCaption = "Average CPUs temp"
        sBaloon = "Average CPUs temp"
        SysTray.Init Me, "Average CPUs temp", currGUIDstr
        SysTray.DrawIcon Me.pic, &HFF00FF
    ElseIf Me.currNumber = modMain.nCPUs + 4 Then  'network
        Call CreateNETRefresher
        log.WriteLog "Loading tray" & Me.currNumber & ": Creating NET Refresher... Success!"
        
        netIcon(devIdle) = LoadIconFromMultiRES("NETICON", 101, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If netIcon(devIdle) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET Idle status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET Idle status... Success!"
        End If
        
        netIcon(devRead) = LoadIconFromMultiRES("NETICON", 102, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If netIcon(devRead) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET Recieve status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET Recieve status... Success!"
        End If
        
        netIcon(devWrite) = LoadIconFromMultiRES("NETICON", 103, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If netIcon(devWrite) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET Send status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET Send status... Success!"
        End If
        
        netIcon(devReadWrite) = LoadIconFromMultiRES("NETICON", 104, lIconIndex, Me.pic.Width, Me.pic.Width, False)
        If netIcon(devReadWrite) = 0 Then
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET SendRecieve status... FAILED!"
        Else
            log.WriteLog "Loading tray" & Me.currNumber & ": Loading icon for NET SendRecieve status... Success!"
        End If
        
        Me.Caption = ""
        sCaption = "Active Network interface:" & vbCrLf & wmSettings.comboNet.Text
        sBaloon = sCaption
        SysTray.Init Me, sCaption, currGUIDstr
        SysTray.DrawIcon Me.pic, , netIcon(devIdle)
        log.WriteLog "Loading tray" & Me.currNumber & ": Tray loaded!"
    End If
   SysTray.TooltipText = ""
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Unload SysTray
'    Set SysTray = Nothing
'End Sub

'Private Sub Form_Terminate()
'    SysTray.Class_Terminate
'    Set SysTray = Nothing
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SysTray.Class_Terminate
    If Not (SysTray Is Nothing) Then Set SysTray = Nothing
    If Me.currNumber <= modMain.nCPUs Then  'CPUs
        objRefresherCPU.DeleteAll
    ElseIf Me.currNumber = modMain.nCPUs + 1 Then 'memory
        objRefresherMEM.DeleteAll
        Set objDraw = Nothing
        DestroyIcon memIcon
    ElseIf Me.currNumber = modMain.nCPUs + 2 Then 'HDD
        objRefresherHDD.DeleteAll
        DestroyIcon hwIcon(0)
        DestroyIcon hwIcon(1)
        DestroyIcon hwIcon(2)
        DestroyIcon hwIcon(3)
    ElseIf Me.currNumber = modMain.nCPUs + 4 Then 'network
        objRefresherNET.DeleteAll
        DestroyIcon netIcon(0)
        DestroyIcon netIcon(1)
        DestroyIcon netIcon(2)
        DestroyIcon netIcon(3)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ProcessTolltip()
    Dim PT As POINTAPI, RT As RECT, NewPT As POINTAPI, PT2 As POINTAPI
    On Error Resume Next
    If (frmTooltip.currentTipTray <> Me.currNumber) Then
        GetCursorPos PT
        SysTray.GetTrayIconRect PT.x, PT.y
        RT.Bottom = SysTray.trayRECT_Bottom
        RT.Left = SysTray.trayRECT_Left
        RT.Right = SysTray.trayRECT_Right
        RT.Top = SysTray.trayRECT_Top
        If Abs(RT.Bottom - PT.y) > 3 And Abs(PT.y - RT.Top) > 3 And Abs(PT.x - RT.Left) > 3 And Abs(RT.Right - PT.x) > 3 Then
            frmTooltip.Timer1.Enabled = False
            frmTooltip.Hide
            If frmTooltip.lblTooltipText <> sBaloon Then frmTooltip.lblTooltipText = sBaloon
            NewPT.x = (RT.Left + RT.Right) / 2
            NewPT.y = RT.Top - (frmTooltip.Height / Screen.TwipsPerPixelY)
            If NewPT.x < 1 Then NewPT.x = RT.Left
            If NewPT.y < 1 Then NewPT.y = RT.Bottom
            NewPT.x = NewPT.x * Screen.TwipsPerPixelX
            NewPT.y = NewPT.y * Screen.TwipsPerPixelY
            If NewPT.x >= (Screen.Width - frmTooltip.Width) Then NewPT.x = Screen.Width - frmTooltip.Width - Screen.TwipsPerPixelX
    
            frmTooltip.currentTipTray = Me.currNumber
            frmTooltip.Move NewPT.x, NewPT.y
            Sleep 50
            frmTooltip.Timer1.Enabled = True
            frmTooltip.Show
            modMain.SetOnTopWindow frmTooltip.hwnd, True
            
            'frmTooltip.Timer1.Enabled = False
            'frmTooltip.ShowTooltip Me.currNumber, NewPT.X, NewPT.Y
        End If
        'Else
            'frmTooltip.Timer1.Enabled = False
            'frmTooltip.currentTipTray = 0
            'frmTooltip.Hide
            ''frmTooltip.HideTooltip
        'End If
    Else
        If Not frmTooltip.Timer1.Enabled Then 'tooltips showing current indicator but timeout is ended. Restart tooltip
            'Sleep 50
            frmTooltip.Timer1.Enabled = True
            frmTooltip.Show
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    SysTray.MouseMove Button, x, Me
    'If -(wmSettings.checkShowTooltips.Value) Then Call ProcessTolltip
    If modMain.ShowAdvancedTooltips Then ProcessTolltip
    'SysTray.TooltipText = Me.Caption
    'SysTray.ShowBalloonTip SysTray.TooltipText, beInformation, , 100
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    SysTray.MouseMove Button, x, Me
    If modMain.ShowAdvancedTooltips Then ProcessTolltip
    'SysTray.TooltipText = Me.Caption
    'SysTray.ShowBalloonTip SysTray.TooltipText, beInformation, , 100
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Select Case Me.mnuPopUp(Index).Caption
        Case "Settings"
            wmSettings.Show
        
        Case "Show Network monitor"
            On Error Resume Next
            For i = UBound(vTrays) To LBound(vTrays) Step -1
                If Not (vTrays(i) Is Nothing) Then
                    vTrays(i).mnuPopUp(2).Checked = Not vTrays(i).mnuPopUp(2).Checked
                End If
            Next i
            
            If Me.mnuPopUp(2).Checked Then
                Load frmNet
                frmNet.Show
            Else
                frmNet.Hide
                Unload frmNet
            End If
            Err.Clear
            On Error GoTo 0
            
        Case "Exit"
            On Error Resume Next
            If modMain.ShowSplashScreen Then
                Load frmSplash
                frmSplash.lblWarning.Caption = "WinMeters is stopping, please wait..."
                frmSplash.Show
                frmSplash.FadeIn
            End If

            frmTooltip.Timer1.Enabled = False
            wmSettings.tmrCPUs.Enabled = False
            wmSettings.tmrHDD.Enabled = False
            wmSettings.tmrMem.Enabled = False
            wmSettings.tmrNet.Enabled = False
            For i = UBound(vTrays) To LBound(vTrays) Step -1
                If Not (vTrays(i) Is Nothing) Then
                    If i <> Me.currNumber Then
                        'Set vTrays(i).SysTray = Nothing
                        Unload vTrays(i)
                        If Not (vTrays(i) Is Nothing) Then Set vTrays(i) = Nothing
                    End If
                End If
            Next i
            Unload frmTooltip
            Unload wmSettings
            Unload frmNet
            SysTray.Class_Terminate
            If Not (SysTray Is Nothing) Then Set SysTray = Nothing
            'Set objChild = Nothing
            Set objDraw = Nothing
            Set objHDD = Nothing
            'Set objParent2 = Nothing
            'Set objParent3 = Nothing
            Set objParentCPU = Nothing
            Set objParentHDD = Nothing
            Set objParentMEM = Nothing
            Set objParentNET = Nothing
            Set objRefresherCPU = Nothing
            Set objRefresherHDD = Nothing
            'Set objRefresherHDD2 = Nothing
            Set objRefresherMEM = Nothing
            Set objRefresherNET = Nothing
            Set objWMIServiceForCPU = Nothing
            Set objWMIServiceForHDD = Nothing
            Set objWMIServiceForMEM = Nothing
            Set objWMIServiceForNET = Nothing
            'Unload Me
            If modMain.ShowSplashScreen Then
                Sleep 200
                frmSplash.FadeOut
                Unload frmSplash
            End If
            Err.Clear
            On Error GoTo 0
            End
    End Select
End Sub

'Private Sub SysTray_BalloonClicked()
'    MsgBox "Balloon tip was clicked", vbInformation, "Notice"
'End Sub

Private Sub SysTray_DoubleClick()
    Dim sBuff As String, lLen As Long, lRes As Long
    'SysTray.ShowBalloonTip "Double click tray icon", beInformation, "Balloon Tip"
    
    'wmSettings.Show
    sBuff = String(255, " ")
    lRes = GetWindowsDirectory(sBuff, 255)
    If lRes <> 0 Then
        lLen = InStr(1, sBuff, Chr$(0))
        sBuff = Left$(sBuff, lLen - 1)
        Shell sBuff & "\system32\taskmgr.exe", vbNormalFocus
    End If
End Sub

'Private Sub SysTray_LeftClick()
'    SysTray.ShowBalloonTip sBaloon, beInformation, "WinMeters info", 1000
'End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mnuSysTray
End Sub

Private Sub MyDrawCircle_OLD(ByVal iPercent As Integer, Optional ByVal iKernel As Integer = 0)
    Dim dPerc As Double, j As Integer, k As Integer
    
    Me.pic.FillStyle = vbSolid

    If iPercent <= 75 Then
        Me.pic.FillColor = &HFF00FF
        Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, &HFF00FF ', -Pi / 2 - dPerc, -Pi / 2
'        If iKernel < iPercent Then
            dPerc = 2 * Pi * iPercent / 100
            Me.pic.FillColor = modMain.rgbMEM 'RGB(0, 64, 200)
            Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, modMain.rgbMEM, -Pi / 2, -Pi / 2 - dPerc
            If wmSettings.chkShowKernelMem.Value = 1 Then
                dPerc = 2 * Pi * iKernel / 100
                Me.pic.FillColor = modMain.rgbMEMKernel
                Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, modMain.rgbMEMKernel, -Pi / 2, -Pi / 2 - dPerc
            End If
'        Else
'            dPerc = 2 * Pi * iKernel / 100
'            Me.pic.FillColor = modMain.rgbMEMKernel
'            Me.pic.Circle (me.lpictwidth \ 2 - 1, me.lpictheight \ 2 - 1), me.lpictwidth \ 2 - 1, modMain.rgbMEMKernel, -Pi / 2, -Pi / 2 - dPerc
'        End If
    Else
'        If iKernel < iPercent Then
            dPerc = 2 * Pi * iPercent / 100
            Me.pic.FillColor = modMain.rgbMEM 'RGB(0, 64, 200)
            Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, modMain.rgbMEM ', -Pi / 2, -2 * Pi
            Me.pic.FillColor = &HFF00FF
            Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, &HFF00FF, -Pi / 2 + (2 * Pi - dPerc), -Pi / 2
            If modMain.ShowKernelMemory Then
                dPerc = 2 * Pi * iKernel / 100
                Me.pic.FillColor = modMain.rgbMEMKernel
                Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, modMain.rgbMEMKernel, -Pi / 2, -Pi / 2 - dPerc
            End If
'        Else
'            dPerc = 2 * Pi * iKernel / 100
'            Me.pic.FillColor = modMain.rgbMEMKernel 'RGB(0, 64, 200)
'            Me.pic.Circle (me.lpictwidth \ 2 - 1, me.lpictheight \ 2 - 1), me.lpictwidth \ 2 - 1, modMain.rgbMEMKernel ', -Pi / 2, -2 * Pi
'            Me.pic.FillColor = &HFF00FF
'            Me.pic.Circle (me.lpictwidth \ 2 - 1, me.lpictheight \ 2 - 1), me.lpictwidth \ 2 - 1, &HFF00FF, -Pi / 2 + (2 * Pi - dPerc), -Pi / 2
'        End If
    End If
    Me.pic.FillStyle = vbSolid
    Me.pic.FillColor = vbBlack
    
    Me.pic.Circle (Me.lPictWidth / 2 - 1, Me.lPictHeight / 2 - 1), Me.lPictWidth / 2 - 1, 0, 2 * Pi, vbBlack
    Me.pic.Refresh
End Sub

Private Function MyDrawCircle(ByVal iPercent As Integer, Optional ByVal iKernel As Integer = 0) As Long
    Dim dPerc As Double, j As Integer, k As Integer, dKernPerc As Double
    If objDraw Is Nothing Then
        Set objDraw = New LineGS
        'objDraw.DIB Me.pic.hDC, Me.pic.Image.handle, me.lpictwidth, me.lpictheight
    End If
    With objDraw
        .CreateEmptyDIB Me.lPictWidth, Me.lPictHeight
        '.CircleDIB me.lpictwidth / 2, me.lpictheight / 2, me.lpictwidth / 2 - 1, me.lpictwidth / 2 - 1, vbBlack, Thin
        If iPercent > 0 Then
            dPerc = iPercent * 360 / 100
            dKernPerc = iKernel * 360 / 100
            .PieDIB Me.lPictWidth / 2, Me.lPictHeight / 2, Me.lPictWidth / 2 - 2, 0, dPerc, modMain.rgbMEM
            If modMain.ShowKernelMemory And (dKernPerc > 2) Then
                .PieDIB Me.lPictWidth / 2, Me.lPictHeight / 2, Me.lPictWidth / 2 - 2, 0, dKernPerc, modMain.rgbMEMKernel
            End If
        End If
        .CircleDIB Me.lPictWidth \ 2, Me.lPictHeight \ 2, Me.lPictWidth \ 2 - 1, Me.lPictWidth / 2 - 1, vbBlack, Thin
        MyDrawCircle = .GetIconFromDIB
    End With
    'Set objDraw = Nothing
End Function

'Private Sub MyDraw_CircleBAD(ByVal iPercent As Integer)
'    Dim dPerc As Double, j As Integer, k As Integer
'
'    'Me.pic.FillStyle = vbTransparent
'    'Me.pic.Circle (15, 15), 15, vbBlack
'    'Me.pic.Circle (Me.pic.Width \ 2 - 1, Me.pic.Height \ 2 - 1), Me.pic.Width \ 2 - 1, 0, 2 * Pi, vbBlack
'
'    Me.pic.FillStyle = vbSolid
'    Me.pic.Cls
'    Me.pic.BackColor = &HFF00FF
'    Me.pic.FillColor = &HFF00FF
'    dPerc = 360 * iPercent / 100
'    ExtFloodFill Me.pic.hDC, 1, 1, &HFF00FF, FLOODFILLSURFACE
'    Me.pic.Circle (Me.pic.ScaleWidth \ 2 - 1, Me.pic.ScaleHeight \ 2 - 1), Me.pic.ScaleWidth \ 2 - 1, modMain.rgbMEM, 0, 2 * Pi
'    'If iPercent < 99 Then
'        Me.pic.FillColor = modMain.rgbMEM
'        MoveToEx Me.pic.hDC, Me.pic.ScaleWidth \ 2 - 1, Me.pic.ScaleHeight \ 2 - 1, ByVal 0&
'        AngleArc Me.pic.hDC, Me.pic.ScaleWidth \ 2 - 1, Me.pic.ScaleHeight \ 2 - 1, Me.pic.ScaleWidth \ 2 - 1, 90, dPerc
'        LineTo Me.pic.hDC, Me.pic.ScaleWidth \ 2 - 1, Me.pic.ScaleHeight \ 2 - 1
'        If iPercent > 2 Then
'            If GetPixel(Me.pic.hDC, Me.pic.ScaleWidth \ 2 - 2, 4) = &HFF00FF Then _
'                        ExtFloodFill Me.pic.hDC, Me.pic.ScaleWidth \ 2 - 2, 4, modMain.rgbMEM, FLOODFILLBORDER
'        End If
'        'Me.pic.Circle (Me.pic.Width \ 2 - 1, Me.pic.Height \ 2 - 1), Me.pic.Width \ 2 - 1, modMain.rgbMEM, -Pi / 2, -Pi / 2 - dPerc
'    'End If
'
'    Me.pic.FillStyle = vbSolid
'    Me.pic.FillColor = vbBlack
'
'    Me.pic.Circle (Me.pic.Width \ 2 - 1, Me.pic.Height \ 2 - 1), Me.pic.Width \ 2 - 1, vbBlack, 0, 2 * Pi
'    Me.pic.Refresh
'End Sub

'Private Function GetGUID() As String
'    Dim MyGUID As GUID, NewGUID As GUID
'    Dim GUIDByte() As Byte, sGUID As String
'    Dim GuidLen As Long
'    CoCreateGuid MyGUID
'    ReDim GUIDByte(80)
'    GuidLen = StringFromGUID2(VarPtr(MyGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
'    sGUID = Left(GUIDByte, GuidLen)
'    GuidLen = CLSIDFromString(StrPtr(sGUID), VarPtr(NewGUID.Data1))
'    If Asc(Right$(sGUID, 1)) = 0 Then sGUID = Left$(sGUID, Len(sGUID) - 1)
'    GetGUID = sGUID
'End Function

Private Function LoadIconFromMultiRES(ResID, ResName, IconIndex, Optional PixelsX = 16, Optional PixelsY = 16, Optional bDefaultSize As Boolean = False) As Long
    Const ICRESVER As Long = &H30000
    Dim IconFile As Long
    Dim IconRes() As Byte
    Dim hIcon As Long
    Dim lDirPos As Long, lSize As Long, lBMPpos As Long
    
    'Dim tmpIcon() As Byte
    
    'Load the icon from desired resource
    IconRes = LoadResData(ResName, ResID)
    
    'Grab the chosen icon from the file Index; 0 = 1st Icon
    lDirPos = 6 + (IconIndex) * 16

    CopyMemory lSize, IconRes(lDirPos + 8), 4&

    CopyMemory lBMPpos, IconRes(lDirPos + 12), 4&
    
    
    'ReDim tmpIcon(lSize - 1)
    'CopyMemory tmpIcon(0), IconRes(lBMPpos), lSize
    'Open App.Path & "\test.ico" For Binary Access Write As #3
    '    Put #3, , tmpIcon
    'Close #3
    
    'Create the Icon File
    If Not bDefaultSize Then
        'hIcon = CreateIconFromResourceEx(IconRes(lBMPpos), UBound(IconRes) - 21&, True, ICRESVER, PixelsX, PixelsY, 0&)
        hIcon = CreateIconFromResourceEx(IconRes(lBMPpos), lSize, True, ICRESVER, PixelsX, PixelsY, LR_DEFAULTCOLOR)
    Else
        'hIcon = CreateIconFromResource(IconRes(lBMPpos), UBound(IconRes) - 21&, True, ICRESVER)
        hIcon = CreateIconFromResource(IconRes(lBMPpos), lSize, True, ICRESVER)
    End If
    
    'If there is data, set the icon to the desired source
    If hIcon > 0 Then LoadIconFromMultiRES = hIcon
End Function

Private Function CreateNETRefresher() As Boolean
    On Error Resume Next
    Set objWMIServiceForNET = GetObject("winmgmts:\\.\root\CIMV2")
    Set objRefresherNET = CreateObject("WbemScripting.Swbemrefresher")
    Set objParentNET = objRefresherNET.AddEnum _
        (objWMIServiceForNET, "Win32_PerfFormattedData_Tcpip_NetworkInterface").ObjectSet
    objRefresherNET.Refresh
    If Err Then CreateNETRefresher = False Else CreateNETRefresher = True
    Err.Clear
    On Error GoTo 0
End Function

'Private Function DropNETRefresher() As Boolean
'    On Error Resume Next
'    objRefresherNET.DeleteAll
'    Set objParentNET = Nothing
'    Set objRefresherNET = Nothing
'    Set objWMIServiceForNET = Nothing
'    If Err Then DropNETRefresher = False Else DropNETRefresher = True
'    Err.Clear
'    On Error GoTo 0
'End Function

Public Function RefreshNET() As Boolean
    Dim bWrite As Boolean, bRead As Boolean, lUsage As Long, objChild As Object, sBandwidth As String
    On Error Resume Next
    If objRefresherNET Is Nothing Then
        If Not CreateNETRefresher Then Exit Function
    End If
    objRefresherNET.Refresh
    bWrite = False: bRead = False

'"CurrentBandwidth:

    For Each objChild In objParentNET
        If objChild.Name = modMain.sActiveNetwork Then
'            If IsNull(objChild.BytesSentPersec) Then
'                lUsage = 0
'            Else
                lUsage = objChild.BytesSentPersec
'            End If
            If lUsage > 0 Then bWrite = True
            
'            If IsNull(objChild.BytesReceivedPersec) Then
'                lUsage = 0
'            Else
                lUsage = objChild.BytesReceivedPersec
'            End If
            If lUsage > 0 Then bRead = True
            sBandwidth = objChild.CurrentBandwidth
            
            Exit For
        End If
    Next
    Set objChild = Nothing
    If netIcon(0) = 0 Or netIcon(1) = 0 Or netIcon(2) = 0 Or netIcon(3) = 0 Then
        If netIcon(0) <> 0 Then DestroyIcon netIcon(0)
        If netIcon(1) <> 0 Then DestroyIcon netIcon(1)
        If netIcon(2) <> 0 Then DestroyIcon netIcon(2)
        If netIcon(3) <> 0 Then DestroyIcon netIcon(3)
        netIcon(0) = 0: netIcon(1) = 0: netIcon(2) = 0: netIcon(3) = 0
        netIcon(devIdle) = LoadIconFromMultiRES("NETICON", 101, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
        netIcon(devRead) = LoadIconFromMultiRES("NETICON", 102, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
        netIcon(devWrite) = LoadIconFromMultiRES("NETICON", 103, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
        netIcon(devReadWrite) = LoadIconFromMultiRES("NETICON", 104, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
    End If
    '"————————————————————————————" & vbCrLf &
    sCaption = "Monitoring network interface:" & vbCrLf & vbCrLf & _
            wmSettings.comboNet.Text & vbCrLf & "Currend bandwidth " & Round(sBandwidth / 1000 / 1000) & " Mbps"
    sBaloon = sCaption
    If Not bRead And Not bWrite And SysTray.GetIcon <> netIcon(devIdle) Then
        'Me.DrawNET devIdle
        SysTray.DrawIcon Me.pic, , netIcon(devIdle)
    ElseIf bRead And Not bWrite And SysTray.GetIcon <> netIcon(devRead) Then
        'Me.DrawNET devRead
        SysTray.DrawIcon Me.pic, , netIcon(devRead)
    ElseIf Not bRead And bWrite And SysTray.GetIcon <> netIcon(devWrite) Then
        'Me.DrawNET devWrite
        SysTray.DrawIcon Me.pic, , netIcon(devWrite)
    ElseIf bRead And bWrite And SysTray.GetIcon <> netIcon(devReadWrite) Then
        'Me.DrawNET devReadWrite
        SysTray.DrawIcon Me.pic, , netIcon(devReadWrite)
    End If
    If frmTooltip.currentTipTray = Me.currNumber And frmTooltip.Timer1.Enabled Then _
                frmTooltip.lblTooltipText.Caption = sBaloon
                
    If Me.mnuPopUp(2).Checked Then
        If frmNet.SysMonIsRunning Then frmNet.CollectSamples
    End If
    If Err Then RefreshNET = False Else RefreshNET = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function CreateHDDRefresher() As Boolean
    On Error Resume Next
    objRefresherHDD.DeleteAll
    Err.Clear
    Set objRefresherHDD = Nothing
    Set objWMIServiceForHDD = Nothing
    Set objParentHDD = Nothing
    'Set objParent2 = Nothing
    Set objWMIServiceForHDD = GetObject("winmgmts:\\.\root\CIMV2")
    'Set objWMIServiceForHDD2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
    Set objRefresherHDD = CreateObject("WbemScripting.Swbemrefresher")
    'Set objRefresherHDD2 = CreateObject("WbemScripting.Swbemrefresher")
    Set objParentHDD = objRefresherHDD.AddEnum _
        (objWMIServiceForHDD, "Win32_PerfFormattedData_PerfDisk_LogicalDisk").ObjectSet
    'Set objParent2 = objRefresherHDD.AddEnum _
        (objWMIServiceForHDD, "Win32_PerfFormattedData_PerfDisk_PhysicalDisk").ObjectSet
    'Set objParent3 = objRefresher.AddEnum _
        (objWMIService, "Win32_DiskDrive")
    objRefresherHDD.Refresh
    'objRefresherHDD2.Refresh
      Set objHDD = Nothing
      Set objHDD = objWMIServiceForHDD.ExecQuery("Select * from Win32_DiskDrive", , 16)
    If Err Then CreateHDDRefresher = False Else CreateHDDRefresher = True
    Err.Clear
    On Error GoTo 0
End Function

'Private Function DropHDDRefresher() As Boolean
'    On Error Resume Next
'    objRefresherHDD.DeleteAll
'    objRefresherHDD2.DeleteAll
'    Set objParent2 = Nothing
'    Set objParentHDD = Nothing
'    Set objRefresherHDD2 = Nothing
'    Set objRefresherHDD = Nothing
'    Set objWMIServiceForHDD2 = Nothing
'    Set objWMIServiceForHDD = Nothing
'    If Err Then DropHDDRefresher = False Else DropHDDRefresher = True
'    Err.Clear
'    On Error GoTo 0
'End Function

Public Function RefreshHDD() As Boolean
    Dim bWrite As Boolean, bRead As Boolean, lUsage As Long
    Dim sDisks() As String
    'Dim sPartitions() As String, sVolName() As String, sTotalSize() As String, sFreeMegabytes() As String, sFileSystem() As String
    Dim DiskInfo() As String, varInfo As Variant
    Dim i As Long, j As Long, objChild As Object, objTmp As Object
    'Dim ddd1 As Long, ddd2 As Long
    
    On Error Resume Next
    
    'log.WriteLog "Refreshing HDD..."
    
    dat2 = Now() - dat1
    'ddd1 = GetTickCount()
    
    If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
    nDisksCurr = objFSO.Drives.Count
    
    'log.WriteLog "Previouse disks count: " & nDisksPrev
    'log.WriteLog "Current disks count: " & nDisksCurr
    
    If (objRefresherHDD Is Nothing) Then  'Or (objRefresherHDD2 Is Nothing) Then
        If Not CreateHDDRefresher Then Exit Function
    End If
    
    'If (objHDD Is Nothing) Then
    '        Set objHDD = objWMIServiceForHDD.ExecQuery("Select * from Win32_DiskDrive", , 16)
    'End If
    
    If (nDisksCurr <> nDisksPrev) And (nDisksPrev <> 0) Then
        'log.WriteLog "Reinitializing HDD Refresher because of disks count changed..."
      For i = 1 To Abs(nDisksCurr - nDisksPrev)
        CreateHDDRefresher
      Next i
    End If
    nDisksPrev = nDisksCurr
    
    objRefresherHDD.Refresh
    bWrite = False: bRead = False
    ReDim sDisks(1 To 6, 0 To 0)
    Set objChild = Nothing
    
    For Each objChild In objParentHDD
        'If objChild.Name <> "_Total" Then
            If Not bWrite Then
                lUsage = 0
                lUsage = objChild.DiskWriteBytesPersec
                If lUsage > 0 Then bWrite = True
            End If
            
            If Not bRead Then
                lUsage = 0
                lUsage = objChild.DiskReadBytesPersec
                If lUsage > 0 Then bRead = True
                lUsage = 0
            End If
        'End If
    Next
    
  If (Second(dat2) Mod 5) = 0 And frmTooltip.Timer1.Enabled And modMain.ShowAdvancedTooltips Then
  'If (Second(dat2) Mod 5) = 0 Then
    
    log.WriteLog "Getting info about harddisks..."
    varInfo = modDiskIO.GetDrivesInfo
    'log.WriteLog "GetDriveInfo success. Retrieving hard disks parameters via WMI..."
    ReDim DiskInfo(0 To UBound(varInfo, 1), 0 To UBound(varInfo, 2))
    DiskInfo = varInfo
    Erase varInfo
    
    Set objChild = Nothing
    For Each objChild In objHDD
        i = objChild.Index
        If i <= UBound(sDisks, 2) Then
            sDisks(1, i) = i
        Else
            ReDim Preserve sDisks(1 To 6, 0 To i)
            sDisks(1, i) = i
        End If
        sDisks(2, i) = objChild.Caption
        sDisks(3, i) = objChild.Size
        sDisks(4, i) = objChild.Status
        If IsNull(objChild.Signature) Or objChild.Signature = 0 Then
            sDisks(5, i) = "0"
        Else
            sDisks(5, i) = Val(objChild.Signature)
        End If
        sDisks(6, i) = objChild.InterfaceType
    Next
    Set objChild = Nothing
    
    'log.WriteLog "Retrieving hard disks parameters via WMI - SUCCESS. Generating tooltip text..."
    
    sBaloon = vbNullString
    For i = 0 To 99 'UBound(sDisks, 2)
        'If i > 0 Then
                sBaloon = sBaloon & "—> " ' & vbCrLf
        If modMain.ExtendedHDDInfo Then
            If i <> 99 Then
                sBaloon = sBaloon & "Drive " & CStr(i) & ": " & sDisks(2, i) & " <" & _
                        Round(sDisks(3, i) / 1000 / 1000 / 1000) & "(" & _
                        Round(sDisks(3, i) / 1024 / 1024 / 1024) & ") GB>, " & IIf(Len(sDisks(6, i)) > 0, sDisks(6, i) & ", ", vbNullString) & _
                        "status: " & sDisks(4, i) '& ", signature <"  & hex$(sDisks(5, i)) & ">" & vbCrLf
                If Val(sDisks(5, i)) <> 0 Then ' here is normal MBR disk with numeric signature
                    sBaloon = sBaloon & ", MBR, signature <" & Hex$(Val(sDisks(5, i))) & ">" & vbCrLf
                Else ' without signature we have GPT disk with disk GUID or unrecognized RAW disk
                    sBaloon = sBaloon & ", GPT or RAW" & vbCrLf
                End If
                If Right$(sBaloon, 2) <> vbCrLf Then sBaloon = sBaloon & vbCrLf
            Else
                sBaloon = sBaloon & "Unmatched volume(s) (RAID, spanned etc):" & vbCrLf
            End If
            
            
            For j = 0 To UBound(DiskInfo, 2)
                If i <> 99 Then
                    If (sDisks(1, i) = DiskInfo(DISK_INFO.DriveIndex, j)) And (DiskInfo(DISK_INFO.TotalMegaBytes, j) > 0) Then
                        sBaloon = sBaloon & _
                                    "   —> partition " & DiskInfo(DISK_INFO.PartitionNumber, j) & _
                                    " <" & DiskInfo(DISK_INFO.DriveBusType, j) & _
                                    ", " & DiskInfo(DISK_INFO.PartitionStyle, j) & ">: " & DiskInfo(DISK_INFO.DrivePath, j) & vbCrLf & _
                                    "      device path " & DiskInfo(DISK_INFO.PartitionPath, j) & vbCrLf '&
                                    
                        If DiskInfo(DISK_INFO.PartitionStyle, j) = "GPT" Then
                            sBaloon = sBaloon & _
                                    "      GPT GUID " & DiskInfo(DISK_INFO.PartitionGPT_GUID, j) & vbCrLf & _
                                    "      GPT Name <" & DiskInfo(DISK_INFO.PartitionGPT_Name, j) & ">" & vbCrLf
                        End If
                        
                        sBaloon = sBaloon & _
                                    "      volume path " & DiskInfo(DISK_INFO.MatchedVolume, j) & vbCrLf '&
                        
                        If DiskInfo(DISK_INFO.PartitionLetter, j) <> modDiskIO.NO_DOS_LETTERS Then
                            sBaloon = sBaloon & _
                                    "      DOS disk letter <" & DiskInfo(DISK_INFO.PartitionLetter, j) & ":>, " '&
                        Else
                            sBaloon = sBaloon & _
                                    "      " & DiskInfo(DISK_INFO.PartitionLetter, j) & ", " '&
                        End If
                        
                        sBaloon = sBaloon & _
                                    "mount point(s): " & DiskInfo(DISK_INFO.VolumeLettersAndFolders, j) & vbCrLf & _
                                    "      label <" & DiskInfo(DISK_INFO.VolumeName, j) & _
                                    ">, f/s <" & DiskInfo(DISK_INFO.VolumeFileSystem, j) & ">"
                        
                        If DiskInfo(DISK_INFO.PartitionStyle, j) = "MBR" Then
                            sBaloon = sBaloon & _
                                    ", s/n <" & DiskInfo(DISK_INFO.VolumeSerial, j) & ">" & vbCrLf '&
                        End If
                                    
                        If Right$(sBaloon, 2) <> vbCrLf Then sBaloon = sBaloon & vbCrLf
                        
                        sBaloon = sBaloon & _
                                    "      " & DiskInfo(DISK_INFO.TotalMegaBytes, j) & " MB total, " & _
                                    DiskInfo(DISK_INFO.FreeMegaBytes, j) & " MB free, " & _
                                    DiskInfo(DISK_INFO.UsedMegaBytes, j) & " MB used" & vbCrLf
                    End If
                Else
                    If (DiskInfo(DISK_INFO.DriveIndex, j) = 99) And (DiskInfo(DISK_INFO.TotalMegaBytes, j) > 0) Then
                        sBaloon = sBaloon & _
                                    "   —> volume " & DiskInfo(DISK_INFO.DrivePath, j) & _
                                    " <" & DiskInfo(DISK_INFO.PartitionStyle, j) & ">: " & vbCrLf & _
                                    "      device path " & DiskInfo(DISK_INFO.PartitionPath, j) & vbCrLf '& _
                                    IIf(Len(DiskInfo(DISK_INFO.DriveName, j)) > 0, _
                                        "      belongs to drive " & DiskInfo(DISK_INFO.DriveName, j) & vbCrLf, vbNullString)
                                    
                        If DiskInfo(DISK_INFO.PartitionStyle, j) = "GPT" Then
                            sBaloon = sBaloon & _
                                    "      GPT GUID " & DiskInfo(DISK_INFO.PartitionGPT_GUID, j) & vbCrLf & _
                                    "      GPT Name <" & DiskInfo(DISK_INFO.PartitionGPT_Name, j) & ">" & vbCrLf
                        End If
                        
                        If DiskInfo(DISK_INFO.PartitionLetter, j) <> modDiskIO.NO_DOS_LETTERS Then
                            sBaloon = sBaloon & _
                                    "      DOS disk letter <" & DiskInfo(DISK_INFO.PartitionLetter, j) & ":>, " '&
                        Else
                            sBaloon = sBaloon & _
                                    "      " & DiskInfo(DISK_INFO.PartitionLetter, j) & ", " '&
                        End If
                        
                        sBaloon = sBaloon & _
                                    "mount point(s): " & DiskInfo(DISK_INFO.VolumeLettersAndFolders, j) & vbCrLf & _
                                    "      label <" & DiskInfo(DISK_INFO.VolumeName, j) & _
                                    ">, f/s <" & DiskInfo(DISK_INFO.VolumeFileSystem, j) & ">"
                        
                        If DiskInfo(DISK_INFO.PartitionStyle, j) = "MBR" Then
                            sBaloon = sBaloon & _
                                    ", s/n <" & DiskInfo(DISK_INFO.VolumeSerial, j) & ">" & vbCrLf '&
                        End If
                                    
                        If Right$(sBaloon, 2) <> vbCrLf Then sBaloon = sBaloon & vbCrLf
                        
                        sBaloon = sBaloon & _
                                    "      " & DiskInfo(DISK_INFO.TotalMegaBytes, j) & " MB total, " & _
                                    DiskInfo(DISK_INFO.FreeMegaBytes, j) & " MB free, " & _
                                    DiskInfo(DISK_INFO.UsedMegaBytes, j) & " MB used" & vbCrLf
                    End If
                End If
            Next j
        Else
            If i <> 99 Then
                sBaloon = sBaloon & "Drive " & CStr(i) & ": " & sDisks(2, i) & " - " & _
                        Round(sDisks(3, i) / 1000 / 1000 / 1000) & "(" & _
                        Round(sDisks(3, i) / 1024 / 1024 / 1024) & ") GB, status: " & sDisks(4, i) & vbCrLf
            Else
                sBaloon = sBaloon & "Unmatched volume(s) (RAID, spanned etc):" & vbCrLf
            End If
            For j = 0 To UBound(DiskInfo, 2)
                If i <> 99 Then
                    If (sDisks(1, i) = DiskInfo(DISK_INFO.DriveIndex, j)) And (DiskInfo(DISK_INFO.TotalMegaBytes, j) > 0) Then
                        sBaloon = sBaloon & _
                                    "   —> partition " & DiskInfo(DISK_INFO.PartitionNumber, j) & _
                                    ": mount(s): " & DiskInfo(DISK_INFO.VolumeLettersAndFolders, j) & _
                                    ", label <" & DiskInfo(DISK_INFO.VolumeName, j) & _
                                    ">, f/s <" & DiskInfo(DISK_INFO.VolumeFileSystem, j) & ">" & vbCrLf & _
                                    "      " & DiskInfo(DISK_INFO.TotalMegaBytes, j) & " MB total, " & _
                                    DiskInfo(DISK_INFO.FreeMegaBytes, j) & " MB free, " & _
                                    DiskInfo(DISK_INFO.UsedMegaBytes, j) & " MB used" & vbCrLf
                    End If
                Else
                    If (DiskInfo(DISK_INFO.DriveIndex, j) = 99) And (DiskInfo(DISK_INFO.TotalMegaBytes, j) > 0) Then
                        sBaloon = sBaloon & _
                                    "   —> volume " & DiskInfo(DISK_INFO.DrivePath, j) & _
                                    ": mount(s): " & DiskInfo(DISK_INFO.VolumeLettersAndFolders, j) & _
                                    ", label <" & DiskInfo(DISK_INFO.VolumeName, j) & _
                                    ">, f/s <" & DiskInfo(DISK_INFO.VolumeFileSystem, j) & ">" & vbCrLf & _
                                    "      " & DiskInfo(DISK_INFO.TotalMegaBytes, j) & " MB total, " & _
                                    DiskInfo(DISK_INFO.FreeMegaBytes, j) & " MB free, " & _
                                    DiskInfo(DISK_INFO.UsedMegaBytes, j) & " MB used" & vbCrLf
                    End If
                End If
            Next j
        End If
        If i = UBound(sDisks, 2) Then
                    i = 98
        End If
    Next i
    Do While Right$(sBaloon, 2) = vbCrLf
        sBaloon = Left$(sBaloon, Len(sBaloon) - 2)
    Loop
    
    sCaption = vbCrLf & "—> Unmatched volume(s) (RAID, spanned etc):"
    If InStrRev(sBaloon, sCaption, , vbTextCompare) = (Len(sBaloon) - Len(sCaption) + 1) Then _
            sBaloon = Left$(sBaloon, Len(sBaloon) - Len(sCaption))
    
    sCaption = sBaloon
    
    If Second(dat2) > 30 Then dat1 = Now()
  End If
  
  'log.WriteLog "Tooltip text was generated successful."
    'Err.Clear
    'On Error GoTo 0
    If hwIcon(0) = 0 Or hwIcon(1) = 0 Or hwIcon(2) = 0 Or hwIcon(3) = 0 Then
        If hwIcon(0) <> 0 Then DestroyIcon hwIcon(0)
        If hwIcon(1) <> 0 Then DestroyIcon hwIcon(1)
        If hwIcon(2) <> 0 Then DestroyIcon hwIcon(2)
        If hwIcon(3) <> 0 Then DestroyIcon hwIcon(3)
        hwIcon(0) = 0: hwIcon(1) = 0: hwIcon(2) = 0: hwIcon(3) = 0
        hwIcon(devIdle) = LoadIconFromMultiRES("HDDICON", 101, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
        hwIcon(devRead) = LoadIconFromMultiRES("HDDICON", 102, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
        hwIcon(devWrite) = LoadIconFromMultiRES("HDDICON", 103, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
        hwIcon(devReadWrite) = LoadIconFromMultiRES("HDDICON", 104, lIconIndex, Me.lPictWidth, Me.lPictWidth, False)
    End If
    If Not bRead And Not bWrite And SysTray.GetIcon <> hwIcon(devIdle) Then
        'Me.DrawHDD devIdle
        SysTray.DrawIcon Me.pic, , hwIcon(devIdle)
    ElseIf bRead And Not bWrite And SysTray.GetIcon <> hwIcon(devRead) Then
        'Me.DrawHDD devRead
        SysTray.DrawIcon Me.pic, , hwIcon(devRead)
    ElseIf Not bRead And bWrite And SysTray.GetIcon <> hwIcon(devWrite) Then
        'Me.DrawHDD devWrite
        SysTray.DrawIcon Me.pic, , hwIcon(devWrite)
    ElseIf bRead And bWrite And SysTray.GetIcon <> hwIcon(devReadWrite) Then
        'Me.DrawHDD devReadWrite
        SysTray.DrawIcon Me.pic, , hwIcon(devReadWrite)
    End If
    If (frmTooltip.currentTipTray = Me.currNumber) And frmTooltip.Timer1.Enabled And (frmTooltip.lblTooltipText.Caption <> sBaloon) Then _
                frmTooltip.lblTooltipText.Caption = sBaloon
    If Err Then RefreshHDD = False Else RefreshHDD = True
    
    'ddd2 = GetTickCount()
    'log.WriteLog "RefreshHDD: cycle period is " & (ddd2 - ddd1) & " milliseconds"
    
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetDriveLetterFromDosDeviceName(ByRef DosDeviceName As String) As String
    Dim sPath As String, sDev As String, lBuff As String, i As Long
    On Error Resume Next
    For i = Asc("A") To Asc("Z")
        sPath = Space$(250)
        sDev = Chr$(i) & ":"
        lBuff = 0
        lBuff = QueryDosDevice(sDev, sPath, 250)
        If lBuff > 0 Then
            sPath = Left$(sPath, lBuff - 2)
            sPath = Replace$(sPath, "\Device\", vbNullString, 1, 1, vbBinaryCompare)
            If sPath = DosDeviceName Then
                GetDriveLetterFromDosDeviceName = sDev
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next i
    GetDriveLetterFromDosDeviceName = vbNullString
    Err.Clear
    On Error GoTo 0
End Function


Private Function CreateMEMRefresher() As Boolean
    On Error Resume Next
    Set objWMIServiceForMEM = GetObject("winmgmts:\\.\root\CIMV2")
    Set objRefresherMEM = CreateObject("WbemScripting.Swbemrefresher")
    'Set objProcessor = objRefresher.AddEnum _
        (objWMIService, "Win32_PerfFormattedData_PerfOS_Processor").objectSet
    Set objParentMEM = objRefresherMEM.AddEnum _
        (objWMIServiceForMEM, "Win32_PerfFormattedData_PerfOS_Memory").ObjectSet
    'Set objParent = objRefresher.AddEnum _
        (objWMIService, "Win32_PerfFormattedData_PerfDisk_LogicalDisk").objectSet
    objRefresherMEM.Refresh
    If Err Then CreateMEMRefresher = False Else CreateMEMRefresher = True
    Err.Clear
    On Error GoTo 0
End Function

Public Function RefreshMEM() As Boolean
    Dim iMemUsage As Long, lPoolPaged As Currency, lPoolNonpaged As Currency, lKernelMem As Long, objChild As Object
    On Error Resume Next
    If objRefresherMEM Is Nothing Then
        If Not CreateMEMRefresher Then Exit Function
    End If
    objRefresherMEM.Refresh
    For Each objChild In objParentMEM
'        If IsNull(objChild.AvailableMBytes) Then
'            iMemUsage = 0
'        Else
            iMemUsage = objChild.AvailableMBytes
'        End If
'PoolNonpagedBytes
'        If IsNull(objChild.PoolNonpagedBytes) Then
'            lPoolNonpaged = 0
'        Else
            lPoolNonpaged = objChild.PoolNonpagedBytes / 1024 / 1024
'        End If

'        If IsNull(objChild.PoolPagedBytes) Then
'            lPoolPaged = 0
'        Else
            lPoolPaged = objChild.PoolPagedBytes / 1024 / 1024
'        End If
        
        lKernelMem = lPoolNonpaged + lPoolPaged
        
        Me.DrawMem 100 - iMemUsage * 100 / lMem, iMemUsage, lKernelMem ' * 100 / lMem
        'DoEvents
    Next
    Set objChild = Nothing
    If Err Then RefreshMEM = False Else RefreshMEM = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function CreateCPURefresher() As Boolean
    On Error Resume Next
    Set objWMIServiceForCPU = GetObject("winmgmts:" _
        & "\\.\root\CIMV2")
    Set objRefresherCPU = CreateObject("WbemScripting.Swbemrefresher")
    Set objParentCPU = objRefresherCPU.AddEnum _
        (objWMIServiceForCPU, "Win32_PerfFormattedData_PerfOS_Processor").ObjectSet
    'Set objParent = objRefresher.AddEnum _
        (objWMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet
    'Set objParent = objRefresher.AddEnum _
        (objWMIService, "Win32_PerfFormattedData_PerfDisk_LogicalDisk").objectSet
    objRefresherCPU.Refresh
    If Err Then CreateCPURefresher = False Else CreateCPURefresher = True
    Err.Clear
    On Error GoTo 0
End Function

Public Function RefreshCPU() As Boolean
    Dim iUsage As Integer, iCPU As Integer, iKernel As Integer, objChild As Object ', bConditional As Boolean
    On Error Resume Next
    If objRefresherCPU Is Nothing Then
        If Not CreateCPURefresher Then Exit Function
    End If
    objRefresherCPU.Refresh
    For Each objChild In objParentCPU
        If modMain.ShowOnlyTotalCPULoad Then
            If objChild.Name = "_Total" Then
                iUsage = objChild.PercentProcessorTime
                iKernel = objChild.PercentPrivilegedTime
                Exit For
            End If
        Else
            If Val(objChild.Name) = Val((Me.currNumber - 1)) Then
                iUsage = objChild.PercentProcessorTime
'                If IsNull(objChild.PercentIdleTime) Then
'                    iUsage = 0
'                Else
'                    iUsage = 100 - objChild.PercentIdleTime
'                End If
                    
                iKernel = objChild.PercentPrivilegedTime
                Exit For
            End If
        End If
    Next
    Set objChild = Nothing
    Me.DrawPercents iUsage, iKernel
    If Err Then RefreshCPU = False Else RefreshCPU = True
    Err.Clear
    On Error GoTo 0
End Function

