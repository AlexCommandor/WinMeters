VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Simple Timer that use API GetTickCount"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const CLASS_NAME As String = "MyTimer"
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent Refresh[(arg1, arg2, ... , argn)]
Public Event Refresh()
Attribute Refresh.VB_Description = "Event raised by Timer interval"

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private lStartTime As Long, lStopTime As Long, lTimeout As Long, bStop As Boolean

Public Sub StopTimer()
Attribute StopTimer.VB_Description = "Stop Timer processing"
    bStop = True
End Sub

Public Sub StartTimer(ByVal Interval As Long)
Attribute StartTimer.VB_Description = "Start Timer countdown with Interval in milliseconds"
    lTimeout = Interval
    bStop = False
    lStartTime = GetTickCount()
    DoTimer
End Sub

Private Sub DoTimer()
    Do
        lStopTime = GetTickCount()
        If lStopTime - lStartTime >= lTimeout Then
            RaiseEvent Refresh
            lStartTime = GetTickCount()
        End If
        DoEvents
    Loop Until bStop
End Sub

Private Sub Class_Initialize()
    bStop = True
End Sub

Private Sub Class_Terminate()
    bStop = True
End Sub
