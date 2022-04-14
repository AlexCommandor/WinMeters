Attribute VB_Name = "log"
Option Explicit

Public Opened As Boolean, LogFile As String, FileNum As Integer

Public Function FileExists(sFileName As String) As Boolean
  Dim iFNum As Integer
  On Error Resume Next
  Err.Clear: iFNum = FreeFile
  Open sFileName For Input As iFNum
  If Err.Number <> 0 Then 'File may be not exists or access denied!
    Err.Clear: FileExists = False
  Else
    Close #iFNum: FileExists = True
  End If
End Function

Public Sub OpenLog(Optional ByVal NewLog As Boolean = True)
    FileNum = FreeFile()
    If FileNum = 0 Then
        Opened = False
        Exit Sub
    End If
    LogFile = App.Path & "\" & App.EXEName & ".log"
    If FileExists(LogFile) Then
        On Error Resume Next
        Kill LogFile
        If Err Then ' something wrong while deleting logfile, creating new with date-time mark
            Err.Clear
            LogFile = App.Path & "\" & App.EXEName & Format$(Now(), "_yyyy-mm-dd_hh-nn-ss") & ".log"
        End If
        On Error GoTo 0
    End If
    On Error Resume Next
    If NewLog Then
        Open LogFile For Output As #FileNum
        If Err Then ' something wrong while deleting logfile, creating new with date-time mark
            Err.Clear
            LogFile = App.Path & "\" & App.EXEName & Format$(Now(), "_yyyy-mm-dd_hh-nn-ss") & ".log"
            Open LogFile For Output As #FileNum
        End If
    Else
        Open LogFile For Append As #FileNum
        If Err Then ' something wrong while deleting logfile, creating new with date-time mark
            Err.Clear
            LogFile = App.Path & "\" & App.EXEName & Format$(Now(), "_yyyy-mm-dd_hh-nn-ss") & ".log"
            Open LogFile For Append As #FileNum
        End If
    End If
    If Err Then
        Close #FileNum
        Opened = False
        Err.Clear
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    Opened = True
End Sub


Public Sub CloseLog()
    If Not Opened Then Exit Sub
    Close #FileNum
    Opened = False
End Sub

Public Sub WriteLog(ByVal Message As String, Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorDescription As String = vbNullString)
    If Not Opened Then Exit Sub
    Print #FileNum, Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "." & Right$(Format$(Timer, "#0.00"), 2) & vbTab & Message;
    If ErrorNumber <> 0 Then Print #FileNum, vbTab & "Err.Number=" & ErrorNumber;
    If Len(ErrorDescription) > 0 Then Print #FileNum, vbTab & "Err.Description=" & ErrorDescription;
    Print #FileNum, vbNullString
End Sub
