VERSION 5.00
Begin VB.Form FrmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5940
   Icon            =   "FrmSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Flash1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2160
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer TmrFlash 
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mSettings 
         Caption         =   "&Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mMinimize 
         Caption         =   "&Minimize"
         Visible         =   0   'False
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mCloseMenu 
         Caption         =   "&Close this menu"
      End
      Begin VB.Menu mSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' Bitmap to Icon
Private Type IconInfo
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type Bitmap
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

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer

' Set tray icon to contents of 32x32 picturebox
Private Function DrawIcon(ppic As PictureBox, Optional plngTransparentColor As Long = -1) As Long
    ppic.AutoRedraw = True
    ppic.Picture = ppic.Image
    If plngTransparentColor < 0 Then
        nid.hIcon = BitmapToIcon(ppic.Picture.Handle)
    Else
        nid.hIcon = BitmapToIconTransparent(ppic.Picture.Handle, plngTransparentColor)
    End If
    If nid.hIcon Then
        UpdateIcon NIM_MODIFY
        'Me.Icon = nid.hIcon
        TmrFlash.Enabled = False
        RaiseEvent TIcon(Me)
    End If
End Function

Public Property Let Tooltip(Value As String)
        
        On Error GoTo Tooltip_Err

        nid.szTip = Value & vbNullChar

        
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Tooltip " & _
               "at line " & Erl
        End
        
End Property

Public Property Get Tooltip() As String
        
        On Error GoTo Tooltip_Err
        

        Tooltip = nid.szTip

        
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Tooltip " & _
               "at line " & Erl
        End
        
End Property

Public Property Let Interval(Value As Integer)
        
        On Error GoTo Interval_Err
        

        TmrFlash.Interval = Value
        UpdateIcon NIM_MODIFY

        
        Exit Property

Interval_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Interval " & _
               "at line " & Erl
        End
        
End Property

Public Property Get Interval() As Integer
        
        On Error GoTo Interval_Err
        

        Interval = TmrFlash.Interval

        
        Exit Property

Interval_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Interval " & _
               "at line " & Erl
        End
        
End Property

Public Property Let TrayIcon(Value)
        
        On Error GoTo TrayIcon_Err
        

        TmrFlash.Enabled = False
        On Error Resume Next
        ' Value can be a picturebox, image, form or string

        Select Case TypeName(Value)

            Case "PictureBox", "Image"
                Me.Icon = Value.Picture
                TmrFlash.Enabled = False
                RaiseEvent TIcon(Me)
                
            Case "String"

                If (UCase(Value) = "DEFAULT") Then

                    TmrFlash.Enabled = True
                    Me.Icon = Flash2.Picture
                    RaiseEvent TIcon(Me)

                Else

                    ' Sting is filename; load icon from picture file.
                    TmrFlash.Enabled = True
                    Me.Icon = LoadPicture(Value)
                   RaiseEvent TIcon(Me)

                End If

            Case Else
                ' It's a form ?
                Me.Icon = Value.Icon
                RaiseEvent TIcon(Me)

        End Select

        If Err.Number <> 0 Then TmrFlash.Enabled = True

        UpdateIcon NIM_MODIFY

        
        Exit Property

TrayIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.TrayIcon " & _
               "at line " & Erl
        End
        
End Property

Private Sub Form_Load()
        
        On Error GoTo Form_Load_Err
        

        Me.Icon = Flash1
        RaiseEvent TIcon(Me)
        Me.Visible = False
        TmrFlash.Enabled = True
        Tooltip = App.EXEName
        mAbout.Caption = "About " & App.EXEName
        UpdateIcon NIM_ADD

        
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Form_Load " & _
               "at line " & Erl
        End
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo Form_MouseMove_Err
        

        Dim result As Long
        Dim msg As Long
   
        ' The Form_MouseMove is intercepted to give systray mouse events.

        If Me.ScaleMode = vbPixels Then

            msg = X

        Else

            msg = X / Screen.TwipsPerPixelX

        End If
      
        Select Case msg

            Case WM_RBUTTONDBLCLK
                RaiseEvent Click("RBUTTONDBLCLK")

            Case WM_RBUTTONDOWN
                RaiseEvent Click("RBUTTONDOWN")

            Case WM_RBUTTONUP
                ' Popup menu: selectively enable items dependent on context.

'                Select Case FSys.Visible

'                    Case True

'                        Select Case FSys.WindowState

'                            Case vbMaximized
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = True
'                                mRestore.Enabled = False

'                            Case vbNormal
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = True
'                                mRestore.Enabled = False

'                            Case vbMinimized
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = False
'                                mRestore.Enabled = True

'                            Case Else
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = True
'                                mRestore.Enabled = True

'                        End Select

'                    Case Else
'                        mRestore.Enabled = True
'                        mMaximize.Enabled = False
'                        mMinimize.Enabled = False

'                End Select
         
                RaiseEvent Click("RBUTTONUP")
                PopupMenu mPopupMenu

            Case WM_LBUTTONDBLCLK
                RaiseEvent Click("LBUTTONDBLCLK")
                'mRestore_Click

            Case WM_LBUTTONDOWN
                RaiseEvent Click("LBUTTONDOWN")

            Case WM_LBUTTONUP
                RaiseEvent Click("LBUTTONUP")

            Case WM_MBUTTONDBLCLK
                RaiseEvent Click("MBUTTONDBLCLK")

            Case WM_MBUTTONDOWN
                RaiseEvent Click("MBUTTONDOWN")

            Case WM_MBUTTONUP
                RaiseEvent Click("MBUTTONUP")

            Case WM_MOUSEMOVE
                RaiseEvent Click("MOUSEMOVE")

            Case Else
                RaiseEvent Click("OTHER....: " & Format$(msg))

        End Select

        
        Exit Sub

Form_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Form_MouseMove " & _
               "at line " & Erl
        End
        
End Sub

Private Sub FSys_Resize()
    
    'On Error Resume Next
    

    ' Event generated my main form. WindowState is stored in LastWindowState, so that
    ' it may be re- set when the menu item "Restore" is selected.

    'If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState

End Sub

Private Sub FSys_Unload(Cancel As Integer)
        
        On Error GoTo FSys_Unload_Err
        

        ' Important: remove icon from tray, and unload this form when
        ' the main form is unloaded.
        UpdateIcon NIM_DELETE
        Unload Me

        
        Exit Sub

FSys_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.FSys_Unload " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mAbout_Click()
        
        On Error GoTo mAbout_Click_Err
        

        MsgBox "SuperStarter project  v." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
                "Selective and flexible launcher for many types of DesktopPublishing documents." & vbCrLf & vbCrLf & _
                "© Copyright 2008-2012, Alex Commandor (alex.commandor@gmail.com) ;)", vbInformation, "About SuperStarter"

        
        Exit Sub

mAbout_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mAbout_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mMinimize_Click()
        
        On Error GoTo mMinimize_Click_Err
        

        'FSys.WindowState = vbMinimized

        
        Exit Sub

mMinimize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mMinimize_Click " & _
               "at line " & Erl
        End
        
End Sub

Public Sub mExit_Click()
        
        On Error GoTo mExit_Click_Err
        

        UpdateIcon NIM_DELETE
        Unload FSys
        End

        
        Exit Sub

mExit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mExit_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mRestore_Click()
        
        On Error GoTo mRestore_Click_Err
        

        ' Don't "restore"  FSys is visible and not minimized.

        'If (FSys.Visible And FSys.WindowState <> vbMinimized) Then Exit Sub

        ' Restore LastWindowState
        'FSys.WindowState = LastWindowState
        'FSys.Visible = True
        'SetForegroundWindow FSys.hwnd

        
        Exit Sub

mRestore_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mRestore_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub UpdateIcon(Value As Long)
        
        On Error GoTo UpdateIcon_Err
        

        ' Used to add, modify and delete icon.

        With nid

            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uID = vbNull
            .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon

        End With

        Shell_NotifyIcon Value, nid

        
        Exit Sub

UpdateIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.UpdateIcon " & _
               "at line " & Erl
        End
        
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
        
        On Error GoTo MeQueryUnload_Err
        

'        If UnloadMode = vbFormControlMenu Then'

            ' Cancel by setting Cancel = 1, minimize and hide main window.
'            Cancel = 1
'            F.WindowState = vbMinimized
'            F.Hide

'        End If

        
        Exit Sub

MeQueryUnload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.MeQueryUnload " & _
               "at line " & Erl
        End
        
End Sub

Public Sub MeResize(ByRef F As Form)
        
        On Error GoTo MeResize_Err
        

'        Select Case F.WindowState

'            Case vbNormal, vbMaximized
                ' Store LastWindowState
'                LastWindowState = F.WindowState

'            Case vbMinimized
'                F.Hide

'        End Select

        
        Exit Sub

MeResize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.MeResize " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mStart_Click()
        
        On Error GoTo mStart_Click_Err
        

        'Call FSys.btnStart_Click

        
        Exit Sub

mStart_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mStart_Click " & _
               "at line " & Erl
        End
        
End Sub


Private Sub TmrFlash_Timer()
    
    On Error Resume Next
    

    ' Change icon.
    Static LastIconWasFlash1 As Boolean
    LastIconWasFlash1 = Not LastIconWasFlash1

    Select Case LastIconWasFlash1

        Case True
            Me.Icon = pic

        Case Else
            Me.Icon = Flash1

    End Select

    RaiseEvent TIcon(Me)
    UpdateIcon NIM_MODIFY

End Sub




' Birmap to icon functions provided by Edgemeal
' Special Thanks to Mike D. Sutton - Http://www.mvps.org/EDais/
Private Function BitmapToIcon(ByVal inBMP As Long) As Long
    Dim IconInf As IconInfo
    Dim BMInf As Bitmap
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
' Modified by Edgemeal for 32x32 pixel type Tray Icon programs.
Private Function BitmapToIconTransparent(ByVal hSrcBMP As Long, Optional ByVal inTransCol As Long = -1) As Long
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
        hMaskBMP = GetColMask(hSrcDC, inTransCol)
        hMaskOldBMP = SelectObject(hMaskDC, hMaskBMP)
        If (hMaskOldBMP) Then ' Overlay inverted mask over source
            Call SetTextColor(hSrcDC, vbWhite)
            Call SetBkColor(hSrcDC, vbBlack)
            Call BitBlt(hSrcDC, 0, 0, 32, 32, hMaskDC, 0, 0, vbSrcAnd)
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

Private Function GetColMask(ByVal inDC As Long, ByVal inMaskCol As Long) As Long
    Dim MaskDC As Long, MaskBMP As Long, OldMask As Long, OldBack As Long
    ' Create a new DC
    MaskDC = CreateCompatibleDC(inDC)
    If (MaskDC) Then ' Create a new 1-bpp Bitmap (DDB)
        MaskBMP = CreateBitmap(32, 32, 1, 1, ByVal 0&)
        If (MaskBMP) Then ' Select Bitmap into DC
            OldMask = SelectObject(MaskDC, MaskBMP)
            If (OldMask) Then ' Set mask colour
                OldBack = SetBkColor(inDC, inMaskCol)
                ' Generate mask image
                If (BitBlt(MaskDC, 0, 0, 32, 32, inDC, 0, 0, vbSrcCopy) <> 0) Then GetColMask = MaskBMP
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

