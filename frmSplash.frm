VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1140
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4545
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WinMeters is initializing, please wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4305
      End
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub ApplyTransparency(ByVal hwnd As Long, ByVal btPercentTrans As Integer)
' transparency (0 - 255)
    Dim lOldStyle  As Long
    Dim bTrans      As Byte

    bTrans = btPercentTrans * 2.55
    lOldStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, bTrans, LWA_ALPHA
End Sub

Public Sub FadeOut()
    Dim i As Integer
    For i = 99 To 1 Step -1
        Call ApplyTransparency(Me.hwnd, i)
        Sleep 5
        DoEvents
    Next
    Call ApplyTransparency(Me.hwnd, 0)
End Sub

Public Sub FadeIn()
    Dim i As Integer
    For i = 0 To 99
        Call ApplyTransparency(Me.hwnd, i)
        Sleep 5
        DoEvents
    Next
    Call ApplyTransparency(Me.hwnd, 100)
End Sub

Private Sub Form_Load()
    Call ApplyTransparency(Me.hwnd, 0)
End Sub


