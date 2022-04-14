VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6915
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Height          =   3735
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6375
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
         Left            =   360
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox checkDigits 
         Caption         =   "Show numerical indicator instead of thermometer"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   5535
      End
      Begin VB.CheckBox chkSolidColor 
         Caption         =   "Use solid colors"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox checkCPU 
         Caption         =   "CPU(s) usage refresh interval (milliseconds):"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.CommandButton cmdSelectUsageColor 
         Height          =   255
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmdSelectKernelColor 
         Height          =   255
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3225
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
         Left            =   240
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   2880
         Width           =   480
      End
      Begin VB.CommandButton cmdSelectKernelColor 
         Height          =   255
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3225
         Width           =   495
      End
      Begin VB.CommandButton cmdSelectUsageColor 
         Height          =   255
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
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
         Left            =   3480
         TabIndex        =   5
         Top             =   3240
         Width           =   1095
      End
      Begin MSComctlLib.Slider slideCPU 
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   500
         SmallChange     =   100
         Min             =   200
         Max             =   1500
         SelStart        =   200
         TickFrequency   =   100
         Value           =   200
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "<- Kernel usage"
         Height          =   195
         Left            =   2160
         TabIndex        =   20
         Top             =   3255
         Width           =   1110
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "<- Main usage"
         Height          =   195
         Left            =   2160
         TabIndex        =   19
         Top             =   2910
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         Height          =   195
         Left            =   285
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indicator colors"
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   2640
         Width           =   1080
      End
   End
   Begin VB.Timer tmrNet 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1680
      Top             =   4440
   End
   Begin VB.Timer tmrCPUs 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   4440
   End
   Begin VB.Timer tmrMem 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4440
   End
   Begin VB.Timer tmrHDD 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   4440
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6665
      _ExtentX        =   11748
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Global parameters"
            Key             =   "Global"
            Object.ToolTipText     =   "Change global program parameters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CPU"
            Key             =   "CPU"
            Object.ToolTipText     =   "Change options for CPU(s) indicator(s)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MEM"
            Key             =   "MEM"
            Object.ToolTipText     =   "Change options for memory indicator"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HDD"
            Key             =   "HDD"
            Object.ToolTipText     =   "Change options for HDD(s) indicator"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "NET"
            Key             =   "NET"
            Object.ToolTipText     =   "Change options for network indicator"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    MsgBox "Place code here to set options w/o closing dialog!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            Frame1(i).Left = 240
            Frame1(i).Enabled = True
        Else
            Frame1(i).Left = -20000
            Frame1(i).Enabled = False
        End If
    Next
    
End Sub
