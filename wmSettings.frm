VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form wmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinMeters Settings"
   ClientHeight    =   8505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9765
   Icon            =   "wmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   651
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCloseMe 
      Cancel          =   -1  'True
      Caption         =   "Close this window"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   53
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit  WinMeters"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   46
      Top             =   7920
      Width           =   975
   End
   Begin VB.CheckBox checkShowSplash 
      Caption         =   "Show Splash Screen during start/stop WinMeters"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   7740
      Width           =   5175
   End
   Begin VB.CheckBox checkAutostart 
      Caption         =   "Autostart WinMeters with Windows"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   8160
      Width           =   5055
   End
   Begin VB.CheckBox checkShowTooltips 
      Caption         =   "Show advanced Tooltips with extra system information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   7320
      Value           =   1  'Checked
      Width           =   5175
   End
   Begin VB.Frame frameAbout 
      Caption         =   "About WinMeters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   6600
      TabIndex        =   40
      Top             =   3480
      Width           =   3015
      Begin VB.TextBox txtAbout 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4335
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   41
         Text            =   "wmSettings.frx":0442
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frameNET 
      Caption         =   "Network indicator parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   6255
      Begin VB.CheckBox checkNET 
         Caption         =   "Show indicator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox comboNet 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1560
         Width           =   5535
      End
      Begin MSComctlLib.Slider slideNET 
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   100
         Min             =   200
         Max             =   1000
         SelStart        =   200
         TickFrequency   =   100
         Value           =   200
      End
      Begin VB.Label lblRefreshInterval 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh interval (milliseconds):"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   51
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         Height          =   195
         Left            =   285
         TabIndex        =   33
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         Height          =   195
         Left            =   5595
         TabIndex        =   32
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblNetInterface 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Network interface to monitoring:"
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   1320
         Width           =   4695
      End
   End
   Begin VB.Frame frameHDD 
      Caption         =   "HDD indicator parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   6255
      Begin VB.CheckBox checkExtendedHDDInfo 
         Caption         =   "Show extended HDDs info in Tooltip"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CheckBox checkHDD 
         Caption         =   "Show indicator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin MSComctlLib.Slider slideHDD 
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   100
         Min             =   200
         Max             =   1000
         SelStart        =   200
         TickFrequency   =   100
         Value           =   200
      End
      Begin VB.Label lblRefreshInterval 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh interval (milliseconds):"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   50
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         Height          =   195
         Left            =   5595
         TabIndex        =   26
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         Height          =   195
         Left            =   285
         TabIndex        =   25
         Top             =   1200
         Width           =   285
      End
   End
   Begin VB.Frame frameMEM 
      Caption         =   "Memory indicator parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   9495
      Begin VB.CheckBox checkAntiAliasedMem 
         Caption         =   "Show antialiased (smooth) memory indicator"
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         Top             =   1200
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CommandButton cmdResetColors 
         Caption         =   "Reset colors"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   8160
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkShowKernelMem 
         Caption         =   "Show kernel memory"
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdSelectMemColor 
         Height          =   255
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox checkMEM 
         Caption         =   "Show indicator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdSelectMemColor 
         Height          =   255
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox picMEM 
         AutoRedraw      =   -1  'True
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
         Left            =   6120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Top             =   600
         Width           =   480
      End
      Begin MSComctlLib.Slider slideMEM 
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   500
         SmallChange     =   500
         Min             =   500
         Max             =   3000
         SelStart        =   500
         TickFrequency   =   500
         Value           =   500
      End
      Begin VB.Label lblRefreshInterval 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh interval (milliseconds):"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   49
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label lblMainUsage 
         AutoSize        =   -1  'True
         Caption         =   "<- Main usage"
         Height          =   195
         Index           =   1
         Left            =   8040
         TabIndex        =   38
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label lblKernelUsage 
         AutoSize        =   -1  'True
         Caption         =   "<- Kernel usage"
         Height          =   195
         Index           =   1
         Left            =   8040
         TabIndex        =   37
         Top             =   855
         Width           =   1110
      End
      Begin VB.Label lblIndicatorColors 
         AutoSize        =   -1  'True
         Caption         =   "Indicator colors"
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   20
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         Height          =   195
         Left            =   285
         TabIndex        =   19
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3000"
         Height          =   195
         Left            =   5595
         TabIndex        =   18
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Frame frameCPU 
      Caption         =   "CPU indicator parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CheckBox checkOneTotalCPU 
         Alignment       =   1  'Right Justify
         Caption         =   "One total indicator"
         Height          =   255
         Left            =   4320
         TabIndex        =   47
         Top             =   240
         Width           =   1600
      End
      Begin VB.CommandButton cmdSelectCPUFont 
         Caption         =   "Select font"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4080
         TabIndex        =   45
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox checkDigits 
         Caption         =   "Show digits instead of thermometer"
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkSolidColor 
         Caption         =   "Use solid colors"
         Height          =   255
         Left            =   6240
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox checkCPU 
         Caption         =   "Show indicator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelectUsageColor 
         Height          =   255
         Index           =   0
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdSelectKernelColor 
         Height          =   255
         Index           =   0
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   825
         Width           =   495
      End
      Begin VB.PictureBox picCPU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   6120
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   480
         Width           =   480
      End
      Begin VB.CommandButton cmdSelectKernelColor 
         Height          =   255
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   825
         Width           =   495
      End
      Begin VB.CommandButton cmdSelectUsageColor 
         Height          =   255
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdResetColors 
         Caption         =   "Reset colors"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   8160
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin MSComctlLib.Slider slideCPU 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   100
         Min             =   200
         Max             =   1500
         SelStart        =   200
         TickFrequency   =   100
         Value           =   200
      End
      Begin VB.Label lblRefreshInterval 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh interval (milliseconds):"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label lblKernelUsage 
         AutoSize        =   -1  'True
         Caption         =   "<- Kernel usage"
         Height          =   195
         Index           =   0
         Left            =   8040
         TabIndex        =   35
         Top             =   855
         Width           =   1110
      End
      Begin VB.Label lblMainUsage 
         AutoSize        =   -1  'True
         Caption         =   "<- Main usage"
         Height          =   195
         Index           =   0
         Left            =   8040
         TabIndex        =   34
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1500"
         Height          =   195
         Left            =   5595
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblIndicatorColors 
         AutoSize        =   -1  'True
         Caption         =   "Indicator colors"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Timer tmrHDD 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1920
      Top             =   -120
   End
   Begin VB.Timer tmrMem 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   -120
   End
   Begin VB.Timer tmrCPUs 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   960
      Top             =   -120
   End
   Begin VB.Timer tmrNet 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   -120
   End
End
Attribute VB_Name = "wmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF


'Private wmsetIcon As Long, wmsetIcoBuf() As Byte

'Public WithEvents tmrCPUs As APITimer
'Public WithEvents tmrMem As APITimer
'Public WithEvents tmrHDD As APITimer


' Show the common dialog for choosing a color.
' Return the chosen color, or -1 if the dialog is canceled
'
' hParent is the handle of the parent form
' bFullOpen specifies whether the dialog will be open with the Full style
' (allows to choose many more colors)
' InitColor is the color initially selected when the dialog is open

' Example:
'    Dim oleNewColor As OLE_COLOR
'    oleNewColor = ShowColorsDialog(Me.hwnd, True, vbRed)
'    If oleNewColor <> -1 Then Me.BackColor = oleNewColor

Private Function ShowColorDialog(Optional ByVal hParent As Long, _
    Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) _
    As Long
    Dim CC As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long

    ' translate the initial OLE color to a long value
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    'fill the ChooseColorStruct struct
    With CC
        .lStructSize = Len(CC)
        .hwndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = lInitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
            CC_FULLOPEN, 0)
    End With

    ' Show the dialog
    If ChooseColor(CC) Then
        'if not canceled, return the color
        ShowColorDialog = CC.rgbResult
    Else
        'else return -1
        ShowColorDialog = -1
    End If
End Function

Private Sub checkAntiAliasedMem_Click()
    modMain.AntialiasedMEMIndicator = -(Me.checkAntiAliasedMem.Value)
    SaveSetting "WinMeters", "MEM", "Smooth", AntialiasedMEMIndicator
    If Not (modMain.vTrays(modMain.nCPUs + 1) Is Nothing) Then
        modMain.vTrays(modMain.nCPUs + 1).DrawMem modMain.vTrays(modMain.nCPUs + 1).lMeterData1 + 1, modMain.vTrays(modMain.nCPUs + 1).lMeterData2 + 1
    End If
End Sub

Private Sub checkAutostart_Click()
    Dim wsShell As Object, sRes As String, sPath As String
    On Error Resume Next
    sPath = Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34)
    Set wsShell = CreateObject("WScript.Shell")
    wsShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters"
    If -(checkAutostart.Value) Then
        wsShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters", sPath, "REG_SZ"
        sRes = wsShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters")
        If sRes <> sPath Then
            MsgBox "Unable to access Windows registry! Try to add WinMeters in Autostart manually.", vbCritical + vbOKOnly, "WinMeters error"
            Set wsShell = Nothing
            Exit Sub
        End If
    End If
    Set wsShell = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

'Private Sub CancelButton_Click()
'    'Unload Me
'    Me.Hide
'End Sub

Private Sub checkCPU_Click()
    Dim n As Integer, objControl As Object, nCounter As Integer
    'Me.checkTemp.Enabled = -(Me.checkCPU.Value)
    If modMain.ShowOnlyTotalCPULoad Then nCounter = 1 Else nCounter = modMain.nCPUs
    If -(Me.checkCPU.Value) Then
        For n = nCounter To 1 Step -1
            Set modMain.vTrays(n) = New WinMetersTray
            modMain.vTrays(n).currNumber = n
            Load modMain.vTrays(n)
        Next n
        Me.tmrCPUs.Enabled = True
        'tmrCPUs.StartTimer modMain.intCPUs
    Else
        Me.tmrCPUs.Enabled = False
        'tmrCPUs.StopTimer
        For n = modMain.nCPUs To 1 Step -1
            If Not (modMain.vTrays(n) Is Nothing) Then
                Unload modMain.vTrays(n)
                Set modMain.vTrays(n) = Nothing
            End If
        Next n
        'If Not (modMain.vTrays(modMain.nCPUs + 3) Is Nothing) Then
        '    Unload modMain.vTrays(modMain.nCPUs + 3)
        '    Set modMain.vTrays(modMain.nCPUs + 3) = Nothing
        'End If
    End If
    
    SaveSetting "WinMeters", "CPU", "Enabled", wmSettings.checkCPU.Value
    modMain.IndicatorEnabled(1) = -(wmSettings.checkCPU.Value)
    'SaveSetting "WinMeters", "CPU", "Temperature", wmSettings.checkTemp.Value
    ' 633x105
    On Error Resume Next
    For Each objControl In Me.Controls
        If objControl.Container.Name = "frameCPU" And objControl.Name <> "checkCPU" Then objControl.Enabled = -(wmSettings.checkCPU.Value)
    Next
    Err.Clear
    On Error GoTo 0
    If -(wmSettings.checkCPU.Value) Then
        Me.cmdSelectUsageColor(1).Enabled = Me.chkSolidColor.Value - 1
        Me.cmdSelectKernelColor(1).Enabled = Me.chkSolidColor.Value - 1
        Me.cmdSelectCPUFont.Enabled = -(Me.checkDigits.Value)
    End If
End Sub

Private Sub checkDigits_Click()
    Dim i As Integer
    modMain.ShowDigitsInsteadThermometer = -(Me.checkDigits.Value)
    SaveSetting "WinMeters", "CPU", "Digits", modMain.ShowDigitsInsteadThermometer
    Me.cmdSelectCPUFont.Enabled = modMain.ShowDigitsInsteadThermometer
    For i = 1 To modMain.nCPUs
        If Not (modMain.vTrays(i) Is Nothing) Then _
                modMain.vTrays(i).DrawPercents modMain.vTrays(i).lMeterData1 + 1, modMain.vTrays(i).lMeterData2 + 1
    Next i
End Sub

Private Sub checkExtendedHDDInfo_Click()
    modMain.ExtendedHDDInfo = -(Me.checkExtendedHDDInfo.Value)
    SaveSetting "WinMeters", "HDD", "ExtendedInfo", modMain.ExtendedHDDInfo
End Sub

Private Sub checkHDD_Click()
    Dim objControl As Object
    If -(Me.checkHDD.Value) Then
        Set modMain.vTrays(modMain.nCPUs + 2) = New WinMetersTray
        modMain.vTrays(modMain.nCPUs + 2).currNumber = modMain.nCPUs + 2
        Load modMain.vTrays(modMain.nCPUs + 2)
        Me.tmrHDD.Enabled = True
        'tmrHDD.StartTimer modMain.intHDD
    Else
        Me.tmrHDD.Enabled = False
        'tmrHDD.StopTimer
        If Not (modMain.vTrays(modMain.nCPUs + 2) Is Nothing) Then Unload modMain.vTrays(modMain.nCPUs + 2)
        Set modMain.vTrays(modMain.nCPUs + 2) = Nothing
    End If
    SaveSetting "WinMeters", "HDD", "Enabled", wmSettings.checkHDD.Value
    modMain.IndicatorEnabled(3) = -(wmSettings.checkHDD.Value)
    On Error Resume Next
    For Each objControl In Me.Controls
        If objControl.Container.Name = "frameHDD" And objControl.Name <> "checkHDD" Then objControl.Enabled = -(wmSettings.checkHDD.Value)
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub checkMEM_Click()
    Dim objControl As Object
    If -(Me.checkMEM.Value) Then
        Set modMain.vTrays(modMain.nCPUs + 1) = New WinMetersTray
        modMain.vTrays(modMain.nCPUs + 1).currNumber = modMain.nCPUs + 1
        Load modMain.vTrays(modMain.nCPUs + 1)
        Me.tmrMem.Enabled = True
        'tmrMem.StartTimer modMain.intMEM
    Else
        Me.tmrMem.Enabled = False
        'tmrMem.StopTimer
        If Not (modMain.vTrays(modMain.nCPUs + 1) Is Nothing) Then Unload modMain.vTrays(modMain.nCPUs + 1)
        Set modMain.vTrays(modMain.nCPUs + 1) = Nothing
    End If
    SaveSetting "WinMeters", "MEM", "Enabled", wmSettings.checkMEM.Value
    modMain.IndicatorEnabled(2) = -(wmSettings.checkMEM.Value)
    On Error Resume Next
    For Each objControl In Me.Controls
        If objControl.Container.Name = "frameMEM" And objControl.Name <> "checkMEM" Then objControl.Enabled = -(wmSettings.checkMEM.Value)
    Next
    Err.Clear
    On Error GoTo 0
    If -(wmSettings.checkMEM.Value) Then
        Me.cmdSelectMemColor(1).Enabled = -(Me.chkShowKernelMem.Value)
    End If
End Sub

Private Sub checkNET_Click()
    Dim objControl As Object
    If -(Me.checkNET.Value) Then
        Set modMain.vTrays(modMain.nCPUs + 4) = New WinMetersTray
        modMain.vTrays(modMain.nCPUs + 4).currNumber = modMain.nCPUs + 4
        Load modMain.vTrays(modMain.nCPUs + 4)
        Me.tmrNet.Enabled = True
    Else
        Me.tmrNet.Enabled = False
        If Not (modMain.vTrays(modMain.nCPUs + 4) Is Nothing) Then Unload modMain.vTrays(modMain.nCPUs + 4)
        Set modMain.vTrays(modMain.nCPUs + 4) = Nothing
    End If
    SaveSetting "WinMeters", "NET", "Enabled", wmSettings.checkNET.Value
    modMain.IndicatorEnabled(4) = -(wmSettings.checkNET.Value)
    On Error Resume Next
    For Each objControl In Me.Controls
        If objControl.Container.Name = "frameNET" And objControl.Name <> "checkNET" Then objControl.Enabled = -(wmSettings.checkNET.Value)
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub checkOneTotalCPU_Click()
    modMain.ShowOnlyTotalCPULoad = -(Me.checkOneTotalCPU.Value)
    SaveSetting "WinMeters", "CPU", "TotalOnly", modMain.ShowOnlyTotalCPULoad
    If modMain.ShowOnlyTotalCPULoad Then
        For i = 2 To modMain.nCPUs
            If Not (modMain.vTrays(i) Is Nothing) Then Unload modMain.vTrays(i)
            Set modMain.vTrays(i) = Nothing
        Next i
    Else
        If Not (modMain.vTrays(1) Is Nothing) Then Unload modMain.vTrays(1)
        Set modMain.vTrays(1) = Nothing
        For i = nCPUs To 1 Step -1
            If (modMain.vTrays(i) Is Nothing) Then
                Set modMain.vTrays(i) = New WinMetersTray
                modMain.vTrays(i).currNumber = i
                Load modMain.vTrays(i)
            End If
        Next i
    End If
End Sub

Private Sub checkShowSplash_Click()
    modMain.ShowSplashScreen = -(wmSettings.checkShowSplash.Value)
    SaveSetting "WinMeters", "Trays", "Splash", modMain.ShowSplashScreen
End Sub

Private Sub checkShowTooltips_Click()
    modMain.ShowAdvancedTooltips = -(wmSettings.checkShowTooltips.Value)
    SaveSetting "WinMeters", "Trays", "Tooltips", modMain.ShowAdvancedTooltips
    If Not modMain.ShowAdvancedTooltips Then
        frmTooltip.Timer1.Enabled = False
        frmTooltip.Hide
    End If
End Sub

'Private Sub checkTemp_Click()
'    If -(Me.checkTemp.Value) Then
'        Set modMain.vTrays(modMain.nCPUs + 3) = New WinMetersTray
'        modMain.vTrays(modMain.nCPUs + 3).currNumber = modMain.nCPUs + 3
'        Load modMain.vTrays(modMain.nCPUs + 3)
'    Else
'        If Not (modMain.vTrays(modMain.nCPUs + 3) Is Nothing) Then Unload modMain.vTrays(modMain.nCPUs + 3)
'        Set modMain.vTrays(modMain.nCPUs + 3) = Nothing
'    End If
'    SaveSetting "WinMeters", "Temperature", "Enabled", wmSettings.checkTemp.Value
'End Sub

Private Sub chkShowKernelMem_Click()
    If Me.chkShowKernelMem.Value = 0 Then
        Me.cmdSelectMemColor(1).Enabled = False
    Else
        Me.cmdSelectMemColor(1).Enabled = True
    End If
    SaveSetting "WinMeters", "MEM", "ShowKernel", Me.chkShowKernelMem.Value
    modMain.ShowKernelMemory = -(Me.chkShowKernelMem.Value)
    TuneMEMIndicator modMain.rgbMEM, modMain.rgbMEMKernel
    'trick! - here we must "shake" real mem indicator because it will not be refreshed immediately but only after changing of mem usage
    If Not (modMain.vTrays(modMain.nCPUs + 1) Is Nothing) Then
        modMain.vTrays(modMain.nCPUs + 1).DrawMem modMain.vTrays(modMain.nCPUs + 1).lMeterData1 + 1, modMain.vTrays(modMain.nCPUs + 1).lMeterData2 + 1
    End If
End Sub

Private Sub chkSolidColor_Click()
    If Me.chkSolidColor.Value = 1 Then
        Me.cmdSelectKernelColor(1).Enabled = False
        Me.cmdSelectUsageColor(1).Enabled = False
    Else
        Me.cmdSelectKernelColor(1).Enabled = True
        Me.cmdSelectUsageColor(1).Enabled = True
    End If
    modMain.ShowSolidColors = -(Me.chkSolidColor.Value)
    SaveSetting "WinMeters", "CPU", "Solid", modMain.ShowSolidColors
    TuneCPUIndicator modMain.rgbCPUsUser, modMain.rgbCPUsUser2, modMain.rgbCPUsKernel, modMain.rgbCPUsKernel2
End Sub

Private Sub cmdCloseMe_Click()
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    On Error Resume Next

        If modMain.ShowSplashScreen Then
            Load frmSplash
            frmSplash.lblWarning.Caption = "WinMeters is stopping, please wait..."
            frmSplash.Show
            frmSplash.FadeIn
        End If

    frmTooltip.Timer1.Enabled = False
    Unload frmTooltip
    wmSettings.tmrCPUs.Enabled = False
    wmSettings.tmrHDD.Enabled = False
    wmSettings.tmrMem.Enabled = False
    wmSettings.tmrNet.Enabled = False
    Unload wmSettings
    For i = UBound(vTrays) To LBound(vTrays) Step -1
        If Not (vTrays(i) Is Nothing) Then
            Unload vTrays(i)
            If Not (vTrays(i) Is Nothing) Then Set vTrays(i) = Nothing
        End If
    Next i
    
        If modMain.ShowSplashScreen Then
            Sleep 200
            frmSplash.FadeOut
            Unload frmSplash
        End If
    
    End
End Sub

Private Sub cmdResetColors_Click(Index As Integer)
    If Index = 0 Then
        modMain.rgbCPUsUser = modMain.colDefaultCPUUser 'RGB(0, 64, 200)
        modMain.rgbCPUsUser2 = modMain.colDefaultCPUUser2 'RGB(0, 255, 255)
        SaveSetting "WinMeters", "CPU", "User", modMain.rgbCPUsUser
        SaveSetting "WinMeters", "CPU", "User2", modMain.rgbCPUsUser2
        Me.cmdSelectUsageColor(0).BackColor = modMain.rgbCPUsUser
        Me.cmdSelectUsageColor(1).BackColor = modMain.rgbCPUsUser2
    
        modMain.rgbCPUsKernel = modMain.colDefaultCPUKernel 'RGB(220, 0, 0)
        modMain.rgbCPUsKernel2 = modMain.colDefaultCPUKernel2 'RGB(255, 255, 0)
        SaveSetting "WinMeters", "CPU", "Kernel", modMain.rgbCPUsKernel
        SaveSetting "WinMeters", "CPU", "Kernel2", modMain.rgbCPUsKernel2
        Me.cmdSelectKernelColor(0).BackColor = modMain.rgbCPUsKernel
        Me.cmdSelectKernelColor(1).BackColor = modMain.rgbCPUsKernel2
        
        TuneCPUIndicator modMain.rgbCPUsUser, modMain.rgbCPUsUser2, modMain.rgbCPUsKernel, modMain.rgbCPUsKernel2
    ElseIf Index = 1 Then
        modMain.rgbMEM = modMain.colDefaultMEMUsage 'RGB(0, 64, 200)
        SaveSetting "WinMeters", "MEM", "Usage", modMain.rgbMEM
        Me.cmdSelectMemColor(0).BackColor = modMain.rgbMEM
        modMain.rgbMEMKernel = modMain.colDefaultMEMKernel 'RGB(220, 0, 0)
        SaveSetting "WinMeters", "MEM", "Kernel", modMain.rgbMEMKernel
        Me.cmdSelectMemColor(1).BackColor = modMain.rgbMEMKernel
        TuneMEMIndicator modMain.rgbMEM, modMain.rgbMEMKernel
    End If
    
    'modMain.rgbdevRead = RGB(0, 220, 0)
    'SaveSetting "WinMeters", "HDD", "Read", modMain.rgbdevRead
    'Me.cmdSelectHDDReadColor.BackColor = modMain.rgbdevRead
    
    'modMain.rgbdevWrite = RGB(220, 0, 0)
    'SaveSetting "WinMeters", "HDD", "Write", modMain.rgbdevWrite
    'Me.cmdSelectHDDWriteColor.BackColor = modMain.rgbdevWrite
    
    'modMain.rgbTemp = RGB(0, 64, 200)
    'SaveSetting "WinMeters", "Temperature", "Color", modMain.rgbTemp
    'Me.cmdSelectTemperatureColor.BackColor = modMain.rgbTemp
End Sub

Private Sub cmdSelectCPUFont_Click()
    Dim NewName As String, NewSize As Integer, NewBold As Boolean, NewItalic As Boolean, lRes As Long
    
    NewName = modMain.FontNameCPU
    NewSize = modMain.FontSizeCPU
    NewBold = modMain.FontBoldCPU
    NewItalic = modMain.FontItalicCPU
    
    'modFontSelector.ShowFont NewName, NewSize
    'Stop
    
    lRes = modFontSelector.GetFont(NewName, NewSize, NewBold, NewItalic, , , , Me.hwnd)
    If lRes = 0 Then Exit Sub
    
    modMain.FontNameCPU = NewName
    SaveSetting "WinMeters", "CPU", "FontName", modMain.FontNameCPU
    
    modMain.FontSizeCPU = NewSize
    SaveSetting "WinMeters", "CPU", "FontSize", modMain.FontSizeCPU
    
    modMain.FontBoldCPU = NewBold
    SaveSetting "WinMeters", "CPU", "FontBold", modMain.FontBoldCPU
    
    modMain.FontItalicCPU = NewItalic
    SaveSetting "WinMeters", "CPU", "FontItalic", modMain.FontItalicCPU
End Sub

'Private Sub cmdSelectHDDReadColor_Click()
'    Dim newRGBColor As OLE_COLOR
'    newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbdevRead)
'    If newRGBColor <> -1 Then
'        modMain.rgbdevRead = newRGBColor
'        Me.cmdSelectHDDReadColor.BackColor = newRGBColor
'        SaveSetting "WinMeters", "HDD", "Read", newRGBColor
'    End If
'End Sub

'Private Sub cmdSelectHDDWriteColor_Click()
'    Dim newRGBColor As OLE_COLOR
'    newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbdevWrite)
'    If newRGBColor <> -1 Then
'        modMain.rgbdevWrite = newRGBColor
'        Me.cmdSelectHDDWriteColor.BackColor = newRGBColor
'        SaveSetting "WinMeters", "HDD", "Write", newRGBColor
'    End If
'End Sub

Private Sub cmdSelectMemColor_Click(Index As Integer)
    Dim newRGBColor As OLE_COLOR
    If Index = 0 Then
        newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbMEM)
    Else
        newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbMEMKernel)
    End If
    If newRGBColor <> -1 Then
        If newRGBColor = &HFF00FF Then
            MsgBox "This color is used for system icon drawing. You cannot use it, sorry...", vbExclamation + vbOKOnly, "WinMeters warning"
            Exit Sub
        End If
        If Index = 0 Then
            modMain.rgbMEM = newRGBColor
            Me.cmdSelectMemColor(0).BackColor = newRGBColor
            SaveSetting "WinMeters", "MEM", "Usage", newRGBColor
            TuneMEMIndicator modMain.rgbMEM, modMain.rgbMEMKernel
        Else
            modMain.rgbMEMKernel = newRGBColor
            Me.cmdSelectMemColor(1).BackColor = newRGBColor
            SaveSetting "WinMeters", "MEM", "Kernel", newRGBColor
            TuneMEMIndicator modMain.rgbMEM, modMain.rgbMEMKernel
        End If
    End If
End Sub

Private Sub cmdSelectKernelColor_Click(Index As Integer)
    Dim newRGBColor As OLE_COLOR
    If Index = 0 Then
        newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbCPUsKernel)
    Else
        newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbCPUsKernel2)
    End If
    If newRGBColor <> -1 Then
        If newRGBColor = &HFF00FF Then
            MsgBox "This color is used for system drawing. You cannot use it, sorry...", vbExclamation + vbOKOnly, "WinMeters warning"
            Exit Sub
        End If
        If Index = 0 Then
            modMain.rgbCPUsKernel = newRGBColor
            SaveSetting "WinMeters", "CPU", "Kernel", newRGBColor
        Else
            modMain.rgbCPUsKernel2 = newRGBColor
            SaveSetting "WinMeters", "CPU", "Kernel2", newRGBColor
        End If
        Me.cmdSelectKernelColor(Index).BackColor = newRGBColor
        TuneCPUIndicator modMain.rgbCPUsUser, modMain.rgbCPUsUser2, modMain.rgbCPUsKernel, modMain.rgbCPUsKernel2
    End If
End Sub

Private Sub cmdSelectUsageColor_Click(Index As Integer)
    Dim newRGBColor As OLE_COLOR
    If Index = 0 Then
        newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbCPUsUser)
    Else
        newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbCPUsUser2)
    End If
    If newRGBColor <> -1 Then
        If newRGBColor = &HFF00FF Then
            MsgBox "This color is used for system drawing. You cannot use it, sorry...", vbExclamation + vbOKOnly, "WinMeters warning"
            Exit Sub
        End If
        If Index = 0 Then
            modMain.rgbCPUsUser = newRGBColor
            SaveSetting "WinMeters", "CPU", "User", newRGBColor
        Else
            modMain.rgbCPUsUser2 = newRGBColor
            SaveSetting "WinMeters", "CPU", "User2", newRGBColor
        End If
        Me.cmdSelectUsageColor(Index).BackColor = newRGBColor
        TuneCPUIndicator modMain.rgbCPUsUser, modMain.rgbCPUsUser2, modMain.rgbCPUsKernel, modMain.rgbCPUsKernel2
    End If
End Sub

'Private Sub cmdSelectTemperatureColor_Click()
'    Dim newRGBColor As OLE_COLOR
'    newRGBColor = ShowColorDialog(Me.hwnd, True, modMain.rgbTemp)
'    If newRGBColor <> -1 Then
'        modMain.rgbTemp = newRGBColor
'        Me.cmdSelectTemperatureColor.BackColor = newRGBColor
'        SaveSetting "WinMeters", "Temperature", "Color", newRGBColor
'    End If
'End Sub

Private Sub comboNet_Click()
    If Me.comboNet.Text = "No working network interfaces found!" Then Exit Sub
    SaveSetting "WinMeters", "NET", "Interface", Me.comboNet.Text
    modMain.sActiveNetwork = Me.comboNet.Text
End Sub

'Private Sub Form_Activate()
'    Me.slideCPU.Value = modMain.intCPUs
'    Me.slideMEM.Value = modMain.intMEM
'    Me.slideHDD.Value = modMain.intHDD
'End Sub

'Private Function LoadIconFromMultiRES(ResID, ResName, IconIndex, Optional PixelsX = 16, Optional PixelsY = 16, Optional bDefaultSize As Boolean = False) As Long
'    Const ICRESVER As Long = &H30000
'    Dim IconFile As Long
'    Dim IconRes() As Byte
'    Dim hIcon As Long
'    Dim lDirPos As Long, lSize As Long, lBMPpos As Long
'
'    'Load the icon from desired resource
'    IconRes = LoadResData(ResName, ResID)
'
'    'Grab the chosen icon from the file Index; 0 = 1st Icon
'    lDirPos = 6 + (IconIndex) * 16
'
'    CopyMemory lSize, IconRes(lDirPos + 8), 4&
'
'    CopyMemory lBMPpos, IconRes(lDirPos + 12), 4&
'
'    'Create the Icon File
'    If Not bDefaultSize Then
'        'hIcon = CreateIconFromResourceEx(IconRes(lBMPpos), UBound(IconRes) - 21&, True, ICRESVER, PixelsX, PixelsY, 0&)
'        hIcon = CreateIconFromResourceEx(IconRes(lBMPpos), lSize, True, ICRESVER, PixelsX, PixelsY, LR_DEFAULTCOLOR)
'    Else
'        'hIcon = CreateIconFromResource(IconRes(lBMPpos), UBound(IconRes) - 21&, True, ICRESVER)
'        hIcon = CreateIconFromResource(IconRes(lBMPpos), lSize, True, ICRESVER)
'    End If
'
'    'If there is data, set the icon to the desired source
'    If hIcon > 0 Then LoadIconFromMultiRES = hIcon
'End Function

Private Sub Form_Activate()
    Dim wsShell As Object, sRes As String, sPath As String
    On Error Resume Next
    sPath = Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34)
    Set wsShell = CreateObject("WScript.Shell")
    sRes = wsShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters")
    If sRes <> sPath Then
        wsShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters"
        Me.checkAutostart.Value = 0
    Else
        Me.checkAutostart.Value = 1
    End If
    Set wsShell = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_GotFocus()
    Call Form_Activate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub Form_Load()
'    Dim wsShell As Object, sRes As String, sPath As String
'    On Error Resume Next
'    sPath = Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34)
'    Set wsShell = CreateObject("WScript.Shell")
'    sRes = wsShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters")
'    If sRes <> sPath Then
'        wsShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinMeters"
'        Me.checkAutostart.Value = 0
'    Else
'        Me.checkAutostart.Value = 1
'    End If
'    Set wsShell = Nothing
'    Err.Clear
'    On Error GoTo 0

    Me.Caption = "WinMeters v" & App.Major & "." & App.Minor & "." & App.Revision & " - Settings"
    Call Form_Activate
    
    Me.Hide
    Me.slideCPU.Value = modMain.intCPUs
    Me.slideMEM.Value = modMain.intMEM
    Me.slideHDD.Value = modMain.intHDD
    Me.slideNET.Value = modMain.intNet
    'Me.cmdSelectHDDReadColor.BackColor = modMain.rgbdevRead
    'Me.cmdSelectHDDWriteColor.BackColor = modMain.rgbdevWrite
    Me.cmdSelectKernelColor(0).BackColor = modMain.rgbCPUsKernel
    Me.cmdSelectKernelColor(1).BackColor = modMain.rgbCPUsKernel2
    Me.cmdSelectKernelColor(1).Enabled = modMain.ShowSolidColors
    Me.cmdSelectMemColor(0).BackColor = modMain.rgbMEM
    Me.cmdSelectMemColor(1).BackColor = modMain.rgbMEMKernel
    Me.cmdSelectMemColor(1).Enabled = -Me.chkShowKernelMem.Value
    'Me.cmdSelectTemperatureColor.BackColor = modMain.rgbTemp
    Me.cmdSelectUsageColor(0).BackColor = modMain.rgbCPUsUser
    Me.cmdSelectUsageColor(1).BackColor = modMain.rgbCPUsUser2
    Me.cmdSelectUsageColor(1).Enabled = modMain.ShowSolidColors
    TuneCPUIndicator modMain.rgbCPUsUser, modMain.rgbCPUsUser2, modMain.rgbCPUsKernel, modMain.rgbCPUsKernel2
    TuneMEMIndicator modMain.rgbMEM, modMain.rgbMEMKernel
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    ElseIf UnloadMode <> vbFormCode Then
        If modMain.ShowSplashScreen Then
            Load frmSplash
            frmSplash.lblWarning.Caption = "WinMeters is stopping, please wait..."
            frmSplash.Show
            frmSplash.FadeIn
        End If
        frmTooltip.Timer1.Enabled = False
        Me.tmrCPUs.Enabled = False
        Me.tmrHDD.Enabled = False
        Me.tmrMem.Enabled = False
        Me.tmrNet.Enabled = False
        For i = UBound(vTrays) To LBound(vTrays) Step -1
            If Not (vTrays(i) Is Nothing) Then
                Unload vTrays(i)
                If Not (vTrays(i) Is Nothing) Then Set vTrays(i) = Nothing
            End If
        Next i
        Unload frmTooltip
        If modMain.ShowSplashScreen Then
            Sleep 200
            frmSplash.FadeOut
            Unload frmSplash
        End If
        Err.Clear
        End
    End If
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Set tmrCPUs = Nothing
'    Set tmrMem = Nothing
'    Set tmrHDD = Nothing
'End Sub

'Private Sub OKButton_Click()
''    modMain.ChangeCPUsInterval Me.slideCPU.Value
''    modMain.ChangeMemInterval Me.slideMEM.Value
''    modMain.ChangeHDDInterval Me.slideHDD.Value
''    modMain.ChangeNETInterval Me.slideNET.Value
'    Me.Hide
'End Sub

Private Sub slideCPU_Change()
    Me.slideCPU.Value = (Me.slideCPU.Value \ 100) * 100
    modMain.ChangeCPUsInterval Me.slideCPU.Value
    DoEvents
End Sub

Private Sub slideCPU_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.slideCPU.TooltipText = Me.slideCPU.Value
End Sub

Private Sub slideHDD_Change()
    Me.slideHDD.Value = (Me.slideHDD.Value \ 100) * 100
    modMain.ChangeHDDInterval Me.slideHDD.Value
    DoEvents
End Sub

Private Sub slideHDD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.slideHDD.TooltipText = Me.slideHDD.Value
End Sub

Private Sub slideMEM_Change()
    Me.slideMEM.Value = (Me.slideMEM.Value \ 500) * 500
    modMain.ChangeMemInterval Me.slideMEM.Value
    DoEvents
End Sub

Private Sub slideMEM_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.slideMEM.TooltipText = Me.slideMEM.Value
End Sub

Private Sub slideNET_Change()
    Me.slideNET.Value = (Me.slideNET.Value \ 100) * 100
    modMain.ChangeNETInterval Me.slideNET.Value
    DoEvents
End Sub

Private Sub slideNET_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.slideNET.TooltipText = Me.slideNET.Value
End Sub

Private Sub tmrCPUs_Timer()
    Me.tmrCPUs.Enabled = modMain.IndicatorEnabled(1)
    If Not modMain.IndicatorEnabled(1) Then Exit Sub
    'Call modMain.MonitorCPUs
    For i = 1 To modMain.nCPUs
        If modMain.ShowOnlyTotalCPULoad And (i > 1) Then Exit For
        If Not vTrays(i) Is Nothing Then vTrays(i).RefreshCPU
    Next i
    'If -(Me.checkTemp.Value) Then modMain.MonitorTemp
End Sub

'Private Sub tmrCPUs_Refresh()
'    Call modMain.MonitorCPUs
'    If -(Me.checkTemp.Value) Then modMain.MonitorTemp
'End Sub

Private Sub tmrHDD_Timer()
    Me.tmrHDD.Enabled = modMain.IndicatorEnabled(3)
    If Not modMain.IndicatorEnabled(3) Then Exit Sub
    'Call modMain.MonitorHDD
    If Not vTrays(modMain.nCPUs + 2) Is Nothing Then vTrays(modMain.nCPUs + 2).RefreshHDD
End Sub

'Private Sub tmrHDD_Refresh()
'    Call modMain.MonitorHDD
'End Sub

Private Sub tmrMem_Timer()
    Me.tmrMem.Enabled = modMain.IndicatorEnabled(2)
    If Not modMain.IndicatorEnabled(2) Then Exit Sub
    'Call modMain.MonitorMEM
    If Not vTrays(modMain.nCPUs + 1) Is Nothing Then vTrays(modMain.nCPUs + 1).RefreshMEM
End Sub

'Private Sub tmrMem_Refresh()
'    Call modMain.MonitorMEM
'End Sub
Private Sub tmrNet_Timer()
    Me.tmrNet.Enabled = modMain.IndicatorEnabled(4)
    If Not modMain.IndicatorEnabled(4) Then Exit Sub
    If Not vTrays(modMain.nCPUs + 4) Is Nothing Then vTrays(modMain.nCPUs + 4).RefreshNET
End Sub


Private Sub TuneCPUIndicator(ByVal colUsage1 As OLE_COLOR, ByVal colUsage2 As OLE_COLOR, _
            ByVal colKernel1 As OLE_COLOR, ByVal colKernel2 As OLE_COLOR)
    Dim currY As Integer, currU As Integer
    On Error Resume Next
    Me.picCPU.Cls
    'If currPercents > 85 Then Stop
    'currY = CInt((30 * currPercents) / 100)
    currY = (Me.picCPU.ScaleWidth - 2) * 90 / 100
    'Me.pic.Line (8, 30)-(21, 30 - currY), RGB(0, 64, 200), BF
    If modMain.ShowSolidColors Then
        '-------------------------
        Me.picCPU.Line (Me.picCPU.ScaleWidth / 4, Me.picCPU.ScaleHeight - 2)-((3 * Me.picCPU.ScaleWidth / 4) - 2, Me.picCPU.ScaleHeight - 1 - currY), modMain.rgbCPUsUser, BF
        '-------------------------
    Else
        DoGradient Me.picCPU, colUsage2, colUsage1, _
            Me.picCPU.ScaleWidth / 4, Me.picCPU.ScaleHeight - 1 - currY, (3 * Me.picCPU.ScaleWidth / 4) - 2, Me.picCPU.ScaleHeight - 2, gradHorizontal
    End If
        'currU = CInt((30 * (currPercents - iUserTime)) / 100)
        currU = (Me.picCPU.ScaleWidth - 2) * 35 / 100
        If modMain.ShowSolidColors Then
        'Me.pic.Line (8, 30)-(21, 30 - currU), vbRed, BF
            '-------------------------
            Me.picCPU.Line (Me.picCPU.ScaleWidth / 4, Me.picCPU.ScaleHeight - 2)-((3 * Me.picCPU.ScaleWidth / 4) - 2, Me.picCPU.ScaleHeight - 1 - currU), modMain.rgbCPUsKernel, BF
            '-------------------------
        Else
            DoGradient Me.picCPU, colKernel2, colKernel1, _
                 Me.picCPU.ScaleWidth / 4, Me.picCPU.ScaleHeight - 1 - currU, (3 * Me.picCPU.ScaleWidth / 4) - 2, Me.picCPU.ScaleHeight - 2, gradHorizontal
        End If
    'Me.pic.Line (8, 31 - currY + 1)-(22, 1), &HFF00FF, BF
    'Me.picCPU.Line (Me.picCPU.ScaleWidth \ 4, Me.picCPU.ScaleHeight - currY - 2)-((3 * Me.picCPU.ScaleWidth \ 4) - 2, 1), &HFF00FF, BF
    'Me.pic.Line (7, 0)-(22, 31), vbBlack, B
    Me.picCPU.Line (Me.picCPU.ScaleWidth / 4 - 1, 0)-((3 * Me.picCPU.ScaleWidth / 4) - 1, Me.picCPU.ScaleHeight - 1), vbBlack, B
    'SysTray.DrawIcon Me.pic, &HFF00FF
'    If iUserTime > 0 Then
'    Else
'        sCaption = IIf(Len(modMain.sExtraCPUsInfo) > 0, modMain.sExtraCPUsInfo, vbNullString) & "==========" & vbCrLf & _
'                "Core #" & currNumber & ": usage " & currPercents & "%"
'    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TuneMEMIndicator(ByVal colUsage As OLE_COLOR, colUsage2 As OLE_COLOR)
    Dim dPerc As Double, j As Integer, k As Integer
    On Error Resume Next
    Me.picMEM.Cls
    Me.picMEM.FillStyle = vbSolid
    
    dPerc = 2 * Pi * 40 / 100
    Me.picMEM.FillColor = colUsage 'RGB(0, 64, 200)
    Me.picMEM.Circle (Me.picMEM.ScaleWidth / 2 - 1, Me.picMEM.ScaleHeight / 2 - 1), Me.picMEM.ScaleWidth / 2 - 1, colUsage, -Pi / 2, -Pi / 2 - dPerc
    
    If Me.chkShowKernelMem = 1 Then
        dPerc = 2 * Pi * 12 / 100
        Me.picMEM.FillColor = colUsage2
        Me.picMEM.Circle (Me.picMEM.ScaleWidth / 2 - 1, Me.picMEM.ScaleHeight / 2 - 1), Me.picMEM.ScaleWidth / 2 - 1, colUsage2, -Pi / 2, -Pi / 2 - dPerc
    End If
    
    Me.picMEM.FillColor = vbBlack
    Me.picMEM.Circle (Me.picMEM.ScaleWidth / 2 - 1, Me.picMEM.ScaleHeight / 2 - 1), Me.picMEM.ScaleWidth / 2 - 1, 0, 2 * Pi, vbBlack
    Me.picMEM.Refresh
    Err.Clear
    On Error GoTo 0
End Sub


'Private Sub TuneMEMIndicator(ByVal colUsage As OLE_COLOR, colUsage2 As OLE_COLOR)
'    Dim tmpMemIcon As Long, tmpObjDraw As New LineGS
'    'tmpMemIcon = MyDrawCircle(40, 12)
'    Dim dPerc As Double, j As Integer, k As Integer, dKernPerc As Double
'    Set tmpObjDraw = New LineGS
'    With tmpObjDraw
'        .CreateEmptyDIB Me.picMEM.ScaleWidth, Me.picMEM.ScaleHeight
'        .CircleDIB Me.picMEM.ScaleWidth \ 2, Me.picMEM.ScaleHeight \ 2, Me.picMEM.ScaleWidth \ 2 - 1, Me.picMEM.ScaleWidth / 2 - 1, vbBlack, Thin
'            dPerc = 40 * 360 \ 100
'            dKernPerc = 12 * 360 \ 100
'            .PieDIB Me.picMEM.ScaleWidth \ 2, Me.picMEM.ScaleHeight \ 2, Me.picMEM.ScaleWidth \ 2 - 2, 0, dPerc, colUsage
'            If Me.chkShowKernelMem.Value = 1 Then
'                .PieDIB Me.picMEM.ScaleWidth \ 2, Me.picMEM.ScaleHeight \ 2, Me.picMEM.ScaleWidth \ 2 - 2, 0, dKernPerc, colUsage2
'            End If
'
'        '.CircleDIB Me.pic.ScaleWidth \ 2, Me.pic.ScaleHeight \ 2, Me.pic.ScaleWidth \ 2 - 1, Me.pic.ScaleWidth / 2 - 1, vbBlack, Thin
'        tmpMemIcon = .GetIconFromDIB
'    End With
'
'    If tmpMemIcon = 0 Then
'        TuneMEMIndicator_OLD colUsage, colUsage2
'    Else
'        Set Me.Image1.Picture = modMain.IconToPicture(tmpMemIcon)
'    End If
'    Set tmpObjDraw = Nothing
'End Sub
