VERSION 5.00
Begin VB.Form frmTooltip 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "frmTooltip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BorderColor     =   &H80000017&
      Height          =   2895
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblTooltipText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   630
   End
   Begin VB.Menu mnuTooltip 
      Caption         =   "TooltipMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyText 
         Caption         =   "Copy text to clipboard"
      End
      Begin VB.Menu mnuCloseMenu 
         Caption         =   "Close this menu"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public currentTipTray As Integer
'Private lIconHDD As Long

'Private Sub Form_Click()
'    Me.Hide
''    currentTipTray = 0
'End Sub

Private Sub Form_Load()
    Me.Hide
    currentTipTray = 0
'    lIconHDD = LoadIconFromMultiRES("HDDICON", 101, 3, , , True)
'    Me.picHD.Picture = IconToPicture(lIconHDD)
'    Me.picHD.Move 0, 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbRightButton) Then
        PopupMenu mnuTooltip
    Else
        Me.Hide
    End If
End Sub

Private Sub lblTooltipText_Change()
    Dim RightPos As Single, BottomPos As Single
    RightPos = Me.Left + Me.Width
    BottomPos = Me.Top + Me.Height
    Me.Width = lblTooltipText.Width + 120
    Me.Height = lblTooltipText.Height + 60
    Shape1.Width = Me.Width
    Shape1.Height = Me.Height
    Me.Move RightPos - Me.Width, BottomPos - Me.Height
End Sub

Private Sub lblTooltipText_Click()
    Me.Hide
'    currentTipTray = 0
End Sub

'Private Sub mnuCloseMenu_Click()
'    'modMain.SetOnTopWindow Me.hwnd, True
'    'mnuTooltip.
'End Sub

Private Sub mnuCopyText_Click()
    Clipboard.Clear
    Clipboard.SetText Me.lblTooltipText.Caption
End Sub

Private Sub lblTooltipText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
    Me.Timer1.Enabled = False
    modMain.SetOnTopWindow Me.hwnd, False
    Me.Hide
'    currentTipTray = 0
End Sub


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

