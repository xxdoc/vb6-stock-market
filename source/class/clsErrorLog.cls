VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum LogFileModeType
   LOG_CURRENT_DATE = 0
   LOG_PROJECT_NAME = 1
End Enum

Public Enum LogModeType
   LOG_MSGBOX = 0
   LOG_TO_FILE = 1
   LOG_FILE_MSGBOX = 2
   LOG_QUIET = 3
   LOG_TO_FILE_EX = 4
End Enum

Public ProjectName As String
Public FileName As String
Public DayKeepLog As Long
Public MsgBoxTitle As String

Public LogFileMode As LogFileModeType

'Public LocalError As Long
Public ModuleName As String
Public RoutineName As String
Public LocalErrorMsg As String
Public SystemErrorMsg As String

'YYYYMMDD
Private Function FormatEx(Data As Date, FormatDate As String) As String
Dim TempDay As Long
Dim TempMonth As Long
Dim TempYear As Long

On Error Resume Next

   TempDay = Day(Data)
   TempMonth = Month(Data)
   TempYear = Year(Data)
   
   If TempYear >= 2500 Then
      TempYear = TempYear - 543
   End If
   If FormatDate = "YYYYMMDD" Then
      FormatEx = Format(CStr(TempYear), "0000") & Format(CStr(TempMonth), "00") & Format(CStr(TempDay), "00")
   Else
      FormatEx = Format(CStr(TempYear), "0000") & "/" & Format(CStr(TempMonth), "00") & "/" & Format(CStr(TempDay), "00") & _
                 " " & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00")
   End If
End Function

Private Function StringToDate(DateStr As String) As Date
Dim TempDay As Long
Dim TempMonth As Long
Dim TempYear As Long

On Error Resume Next
   'YYYYMMDD
   TempDay = CLng(Mid(DateStr, 7, 2))
   TempMonth = CLng(Mid(DateStr, 5, 2))
   TempYear = CLng(Mid(DateStr, 1, 4))
   
   If TempYear >= 2500 Then
      TempYear = TempYear - 543
   End If
   StringToDate = DateSerial(TempYear, TempMonth, TempDay)
End Function

Public Function AskMessage() As VbMsgBoxResult
   AskMessage = MsgBox(LocalErrorMsg, vbQuestion Or vbYesNo, MsgBoxTitle)
End Function

Public Sub ShowUserError()
   Call MsgBox(LocalErrorMsg, vbExclamation Or vbOKOnly, MsgBoxTitle)
End Sub

Public Sub ShowUserErrorEx()
   Call ShowErrorLog(LOG_TO_FILE)
End Sub

Public Sub LogToDB(DBFlag As Boolean)
'On Error Resume Next
'Dim E As CErrorLog
'
'   If Not DBFlag Then
'      Exit Sub
'   End If
'
'   Set E = New CErrorLog
'   E.AddEditMode = SHOW_ADD
'   E.USER_ID = -1
'   E.ERROR_DATE = Now
'   E.FILE_NAME = ModuleName
'   E.MODULE = RoutineName
'   E.LOG_MSG = LocalErrorMsg
'   E.SQL = SystemErrorMsg
'   Call E.AddEditData
'   Set E = Nothing
End Sub

Public Sub ShowErrorLog(Mode As LogModeType, Optional DBFlag As Boolean = True)
On Error Resume Next
Dim ErrMsg As String
Dim Directory As String
Dim OldFile As String
Dim FileNum As Long
   Dim MyName As String
   Dim FD As Date

   FileNum = FreeFile()
      
   If LogFileMode = LOG_CURRENT_DATE Then
      FileName = FormatEx(Now, "YYYYMMDD") & ".LOG"
      If Right(App.PATH, 1) = "\" Then
         FileName = App.PATH & "Log\" & FileName
         Directory = App.PATH & "Log\"
      Else
         FileName = App.PATH & "\" & "Log\" & FileName
         Directory = App.PATH & "\" & "Log\"
      End If
                  
      MyName = Dir(Directory & "*.LOG")
      While MyName <> ""
         FD = FileDateTime(Directory & MyName)
         If DateDiff("d", FD, Now) > DayKeepLog Then
            Kill (Directory & MyName)
         End If
         MyName = Dir
      Wend
            
      If Dir(Directory, vbDirectory) = "" Then
         MkDir (Directory)
      End If
   Else
      FileName = App.EXEName & ".LOG"
      If Right(App.PATH, 1) = "\" Then
         FileName = App.PATH & FileName
      Else
         FileName = App.PATH & "\" & FileName
      End If
   End If
   
   ErrMsg = "" & FormatEx(Now, "YYYY/MM/DD HH:MM:SS") & "  Project name <" & ProjectName & "> : Module name <" & ModuleName & "> :"
   ErrMsg = ErrMsg & "Routine name <" & RoutineName & "> : Local error message <" & LocalErrorMsg & "> : "
   ErrMsg = ErrMsg & "System error message <" & SystemErrorMsg & ">"
'ErrMsg = LocalErrorMsg & "; In routine " & RoutineName
   Debug.Print SystemErrorMsg
   
   Select Case Mode
      Case LOG_MSGBOX:
         MsgBox ErrMsg, , MsgBoxTitle
         
      Case LOG_TO_FILE:
         Close #FileNum
         Open FileName For Append As #FileNum
         Print #FileNum, ErrMsg
         Close #FileNum
'MsgBox ErrMsg, , MsgBoxTitle
'Call LogToDB(DBFlag)

      Case LOG_TO_FILE_EX:
         Close #FileNum
         Open FileName For Append As #FileNum
         Print #FileNum, ErrMsg
         Close #FileNum
Call LogToDB(DBFlag)
         
      Case LOG_FILE_MSGBOX:
         Close #FileNum
         Open FileName For Append As #FileNum
         Print #FileNum, ErrMsg
         Close #FileNum
         MsgBox ErrMsg, , MsgBoxTitle
Call LogToDB(DBFlag)
      
      Case LOG_QUIET:
         Debug.Print ErrMsg
   End Select
End Sub

Private Sub Class_Initialize()
On Error Resume Next

   ProjectName = App.EXEName
   FileName = App.EXEName & "." & "LOG"
   DayKeepLog = 1
   LogFileMode = LOG_PROJECT_NAME
   MsgBoxTitle = "GeneticSoft Print Factory "
End Sub
