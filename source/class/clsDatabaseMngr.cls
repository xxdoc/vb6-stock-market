VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabaseMngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "clsDatabaseMngr"

Private Enum EVENT_TYPE
   EVENT_NONE = -1
   EVENT_CONNECTED = 0
   EVENT_ERROR = 1
   EVENT_CLOSE = 2
   EVENT_DATARECEIVED = 3
   EVENT_SENTCOMPLETE = 4
End Enum

Private m_ExpressADOConn As ADODB.Connection
Private m_ADOConn As ADODB.Connection
Private m_Rs As ADODB.Recordset
Public WithEvents m_Winsock As Winsock
Attribute m_Winsock.VB_VarHelpID = -1
Private m_Event As EVENT_TYPE
Private m_LastError As String
Private m_StringReceive As String
Private m_HasConnect As Boolean
Private m_IP As String
Private m_Port As Long

Private Type TProvince
   PROVINCE_ID As Long
   PROVINCE_NAME As String
End Type

Private Type TExpenseType
   EXPENSE_TYPE_ID As Long
   EXPENSE_TYPE_NAME As String
End Type

Private m_Province() As TProvince
Private m_ExpenseType() As TExpenseType

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
Dim RName As String

   RName = "Class_Initialize"
   
   Set m_ADOConn = New ADODB.Connection
   Set m_ExpressADOConn = New ADODB.Connection
   Set m_Rs = New ADODB.Recordset
   
   Load frmWinsock
   Set m_Winsock = frmWinsock.Winsock1
   
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Property Get DBConnection() As ADODB.Connection
   Set DBConnection = m_ADOConn
End Property
Public Property Get DBConnectionExpress() As ADODB.Connection
   Set DBConnectionExpress = m_ExpressADOConn
End Property
Public Function GetServerDateTime(ServerDateTime As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim ItemCount As Long

   RName = "GetServerDateTime"
   GetServerDateTime = False
   
'ServerDateTime = DateToStringInt(Now)
'GetServerDateTime = True
'Exit Function

   SQL = "SELECT CURRENT_TIMESTAMP || '' FROM SYSTEM_PARAM WHERE PARAM_NAME = 'DUMMY_PARAM' "
   If Not GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "End of file condition."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   If NVLS(m_Rs(0), "") = "" Then
      ErrorObj.LocalErrorMsg = "TIMESTAMP is null."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   ServerDateTime = m_Rs(0)
   
   GetServerDateTime = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL -> " & SQL
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetServerDateTime = False
End Function

Public Function GetSeqID(SequenceName As String, ID As Long, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim ItemCount As Long

   RName = "GetSeqID"
   GetSeqID = False
   
   SQL = "SELECT GEN_ID(" & SequenceName & ", 1) FROM SYSTEM_PARAM WHERE PARAM_NAME = 'DUMMY_PARAM' "
   If Not GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "End of file condition."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   If NVLI(m_Rs(0), 0) = 0 Then
      ErrorObj.LocalErrorMsg = "ID is null."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   ID = m_Rs(0)
   
   GetSeqID = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL -> " & SQL
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetSeqID = False
End Function

Public Function GetRs(SQL1 As String, SQL2 As String, MustCount As Boolean, ItemCount As Long, Rs As ADODB.Recordset, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "GetRs"
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   If MustCount Then
      Call Rs.Open(SQL2, m_ADOConn, adOpenKeyset, adLockOptimistic)
      If Rs.EOF Then
         ItemCount = 0
      Else
         ItemCount = NVLI(Rs(0), 0)
      End If
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   m_ADOConn.CursorLocation = adUseClient
   Call Rs.Open(SQL1, m_ADOConn, adOpenKeyset, adLockOptimistic)

   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   If Rs.State <> adStateOpen Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   GetRs = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetRs = False
End Function
Public Function GetRsExpress(SQL1 As String, SQL2 As String, MustCount As Boolean, ItemCount As Long, Rs As ADODB.Recordset, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "GetRsExpress"
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRsExpress = False
      Exit Function
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   If MustCount Then
      Call Rs.Open(SQL2, m_ExpressADOConn, adOpenKeyset, adLockOptimistic)
      If Rs.EOF Then
         ItemCount = 0
      Else
         ItemCount = NVLI(Rs(0), 0)
      End If
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   m_ExpressADOConn.CursorLocation = adUseClient
   Call Rs.Open(SQL1, m_ExpressADOConn, adOpenKeyset, adLockOptimistic)

   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRsExpress = False
      Exit Function
   End If
   
   If Rs.State <> adStateOpen Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRsExpress = False
      Exit Function
   End If
   
   GetRsExpress = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetRsExpress = False
End Function
Public Function GetComputerName() As String
   GetComputerName = m_Winsock.LocalHostName
End Function

Private Function Connet(IP As String, Port As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "Connect"
   Connet = False
#If WIN_DEBUG = 1 Then
   Connet = True
   Exit Function
#End If

   Call m_Winsock.Connect(IP, Port)
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_CONNECTED Then
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.RoutineName = RName
      ErrorObj.LocalErrorMsg = "Can't connect to " & IP & " : " & Port & ""
      ErrorObj.SystemErrorMsg = m_LastError
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      
'      m_Winsock.Close
      Exit Function
   End If
   
   Connet = True
   Exit Function
   
ErrorHandler:
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Connet = False
End Function

Private Function GetResult(Command As String, FullCommand As String, Result As String, ErrorStr As String) As Boolean
Dim Token As TOKENLIB
Dim Buf As String * 255
Dim tokenstr As String * 255
Dim TokenStrLen As Long
Dim NewToken As String
Dim RName As String

   RName = "GetResult"
   GetResult = False
   
   If TOKInitial(Token) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf
      Exit Function
   End If
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 1, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & FullCommand
      Exit Function
   End If
   NewToken = Mid(tokenstr, 1, TokenStrLen)
   If UCase(NewToken) <> "RESPONSE" Then
      Call TOKTerminate(Token)
      ErrorStr = "Not valid valid response command -> " & FullCommand & " "
      Exit Function
   End If
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 2, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & "& fullcommand &"
      Exit Function
   End If
   NewToken = Trim(Mid(tokenstr, 1, TokenStrLen))
   If UCase(Command) <> UCase(NewToken) Then
      Call TOKTerminate(Token)
      ErrorStr = "Not valid valid response command -> " & FullCommand & " "
      Exit Function
   End If
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 3, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & FullCommand
      Exit Function
   End If
   NewToken = Mid(tokenstr, 1, TokenStrLen)
   If Not IsNumeric(NewToken) Then
      Call TOKTerminate(Token)
      ErrorStr = "Not numeric result value -> " & FullCommand & " "
      Exit Function
   End If
   
   Result = NewToken
   
   TokenStrLen = Len(tokenstr)
   If TOKGetTokenStr(Token, FullCommand, Len(FullCommand), 4, tokenstr, TokenStrLen) <> 0 Then
      Call TOKLastErrorStr(Token, Buf, Len(Buf))
      Call TOKTerminate(Token)
      ErrorStr = Buf & "-> " & FullCommand
      Exit Function
   End If
   NewToken = Mid(tokenstr, 1, TokenStrLen)
   
   ErrorStr = NewToken
   
   GetResult = True
End Function

Public Function Login(UserName As String, UserID As String, CanLogin As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "Login"
   Login = False
   
#If WIN_DEBUG = 1 Then
   CanLogin = True
   Login = True
   Exit Function
#End If

   Command = "(login, " & UserName & ", " & UserID & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("login", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Result <> 0 Then
      ErrorObj.LocalErrorMsg = ErrorStr
      ErrorObj.SystemErrorMsg = m_StringReceive
      
      Login = True
      CanLogin = False
      Exit Function
   End If
   
   CanLogin = True
   Login = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   Login = False
End Function

Private Function IsLogin(UserID As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "IsLogin"
   IsLogin = False
   
   Command = "(islogin,  " & UserID & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("islogin", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Trim(Result) = 0 Then
      IsLogin = False
   Else
      IsLogin = True
   End If
   
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   IsLogin = False
End Function

Public Function LockTable(TableName As String, KEY As Long, IsCanLock As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "LockTable"
   LockTable = False

#If WIN_DEBUG = 1 Then
   LockTable = True
   IsCanLock = True
   Exit Function
#End If
   
   Command = "(lock, " & TableName & ", " & KEY & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("lock", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Result <> 0 Then
      ErrorObj.LocalErrorMsg = ErrorStr
      ErrorObj.SystemErrorMsg = m_StringReceive
      
      LockTable = True
      IsCanLock = False
      Exit Function
   End If
   
   IsCanLock = True
   LockTable = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   LockTable = False
End Function

Public Function UnLockTable(TableName As String, KEY As Long, IsCanUnlock As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim Command As String
Dim Result As String
Dim ErrorStr As String

   RName = "UnLockTable"
   UnLockTable = False
   
#If WIN_DEBUG = 1 Then
   UnLockTable = True
   IsCanUnlock = True
   Exit Function
#End If
   
   Command = "(unlock, " & TableName & ", " & KEY & ")"
'   Call m_Winsock.SendData(Command)
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event = EVENT_DATARECEIVED Then
      GoTo XXX
   End If
   If m_Event <> EVENT_SENTCOMPLETE Then
      ErrorObj.LocalErrorMsg = "Send complete event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
   
   m_Event = EVENT_NONE
   While m_Event = EVENT_NONE
      DoEvents
   Wend
   If m_Event <> EVENT_DATARECEIVED Then
      ErrorObj.LocalErrorMsg = "Data receive event not found."
      ErrorObj.SystemErrorMsg = m_LastError
      Exit Function
   End If
XXX:
   If Not GetResult("unlock", m_StringReceive, Result, ErrorStr) Then
      ErrorObj.LocalErrorMsg = "Can't extract value from response command."
      ErrorObj.SystemErrorMsg = ErrorStr
      Exit Function
   End If
   
   If Result <> 0 Then
      ErrorObj.LocalErrorMsg = ErrorStr
      ErrorObj.SystemErrorMsg = m_StringReceive
      
      UnLockTable = True
      IsCanUnlock = False
      Exit Function
   End If
   
   IsCanUnlock = True
   UnLockTable = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   
   UnLockTable = False
End Function

Public Function ConnectDatabase(FileDb As String, UserName As String, Password As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String
   
   RName = "ConnectDatabase"
      
'   If Not (Right(FileDb, 13) = "\DATABASE.GDB") Then
'      GoTo ErrorHandler
'   End If
   ConnStr = "data source=" & FileDb & ";ctype=UNICODE_FSS"
   m_ADOConn.Provider = "LCPI.IBProvider"
   Call m_ADOConn.Open(ConnStr, UserName, Password)
   
   m_ADOConn.Properties("Session AutoCommit") = True   ' enable auto commit
   m_ADOConn.Properties("Autocommit Isolation Levels") = 4096  ' read commited=0x1000
      
   ConnectDatabase = True
   m_HasConnect = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "กรุณาใส่การเชื่อมต่อของฐานข้อมูลให้ถูกต้อง"
   glbErrorLog.ShowUserError
   
'   Unload frmSplash
'   Set frmSplash = Nothing
   
   ConnectDatabase = False
   m_HasConnect = False
End Function
Public Function ConnectDatabaseExpress(FileDb As String, UserName As String, Password As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String

   RName = "ConnectDatabaseExpress"
'FileDb = "D:\express\DTS2004"
   ConnStr = "Provider=VFPOLEDB.1; Data Source='" & FileDb & "'"
   Call m_ExpressADOConn.Open(ConnStr, UserName, Password)
   
   ConnectDatabaseExpress = True
   
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)

   Unload frmSplash
   Set frmSplash = Nothing
   
   ConnectDatabaseExpress = False
End Function

Public Function DisConnectDatabase() As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String
   
   RName = "DisConnectDatabase"
   Call m_ADOConn.Close
      
   DisConnectDatabase = True
   m_HasConnect = False
   Exit Function
   
ErrorHandler:
'   glbErrorLog.LocalErrorMsg = "Runtime error."
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.RoutineName = RName
'   glbErrorLog.ModuleName = MODULE_NAME
'   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)
   
   DisConnectDatabase = False
   m_HasConnect = False
End Function
Public Function DisConnectDatabaseExpress() As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String

   RName = "DisConnectDatabaseExpress"
   Call m_ExpressADOConn.Close
   
   DisConnectDatabaseExpress = True
   Exit Function

ErrorHandler:
'   glbErrorLog.LocalErrorMsg = "Runtime error."
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.RoutineName = RName
'   glbErrorLog.ModuleName = MODULE_NAME
'   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)
   
   DisConnectDatabaseExpress = False
End Function

Public Function IsConnect() As Boolean
   IsConnect = m_HasConnect
End Function
Public Function ConnectAgentServer(IP As String, Port As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "ConnectAgentServer"
   ConnectAgentServer = False
   
   If Not Connet(IP, Port, ErrorObj) Then
      Exit Function
   End If
   
   ConnectAgentServer = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

   Unload frmSplash
   Set frmSplash = Nothing
   
   ConnectAgentServer = False
End Function

Private Function ExtractDate(DateStr As String) As String
   ExtractDate = Mid(DateStr, 1, 10)
End Function

Private Function ExtractTime(DateStr As String) As String
   ExtractTime = Mid(DateStr, 12, 8)
End Function

Private Sub Class_Terminate()
   Set m_Winsock = Nothing
   
   Unload frmWinsock
   Set frmWinsock = Nothing
End Sub

Private Sub m_Winsock_Close()
   m_Event = EVENT_CLOSE
   m_HasConnect = False
End Sub

Private Sub m_Winsock_Connect()
   m_Event = EVENT_CONNECTED
   m_HasConnect = True
End Sub

Private Sub m_Winsock_DataArrival(ByVal bytesTotal As Long)
   m_Event = EVENT_DATARECEIVED
'   Call m_Winsock.GetData(m_StringReceive, vbString)
End Sub

Private Sub m_Winsock_Error(ByVal Number As Integer, DESCRIPTION As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   m_Event = EVENT_ERROR
   m_LastError = DESCRIPTION
   m_HasConnect = False
End Sub

Private Sub m_Winsock_SendComplete()
   m_Event = EVENT_SENTCOMPLETE
End Sub

Public Function CountRecord(TableName As String, FieldName1 As String, FieldName2 As String, FieldValue As String, ID As Long, ErrorObj As clsErrorLog, Optional FieldNameExTend As String, Optional FieldNameExTendValue As String, Optional FieldNameExTendEX As String, Optional FieldNameExTendValueEX As String, Optional NullFlag As Boolean = False) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim ItemCount As Long

   RName = "CountRecord"
   CountRecord = False
    
    SQL1 = "SELECT COUNT(*) RC FROM " & TableName & " WHERE (" & FieldName1 & " = '" & FieldValue & "') AND (" & FieldName2 & " <> " & ID & ") "
    
    If Len(FieldNameExTendValue) > 0 Then
      SQL1 = SQL1 & " AND (" & FieldNameExTend & " = '" & FieldNameExTendValue & "')"
   ElseIf NullFlag Then
      SQL1 = SQL1 & " AND (" & FieldNameExTend & " IS NULL)"
   End If
   If Len(FieldNameExTendValueEX) > 0 Then
      SQL1 = SQL1 & " AND (" & FieldNameExTendEX & " = '" & FieldNameExTendValueEX & "')"
   ElseIf NullFlag Then
      SQL1 = SQL1 & " AND (" & FieldNameExTendEX & " IS NULL)"
   End If
   
   If Not GetRs(SQL1, "", False, ItemCount, m_Rs, ErrorObj) Then
      Exit Function
   End If
   
   If m_Rs.EOF Then
      CountRecord = 0
   Else
      CountRecord = NVLI(m_Rs("RC"), 0)
   End If

   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   CountRecord = False
End Function
Public Function CountRecordKey(TableName As String, FieldName1 As String, FieldValue As String, ErrorObj As clsErrorLog) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim ItemCount As Long

   RName = "CountRecord"
   CountRecordKey = False
    
    SQL1 = "SELECT COUNT(*) RC FROM " & TableName & " WHERE (" & FieldName1 & " = '" & FieldValue & "')"
    
   If Not GetRs(SQL1, "", False, ItemCount, m_Rs, ErrorObj) Then
      Exit Function
   End If
   
   If m_Rs.EOF Then
      CountRecordKey = 0
   Else
      CountRecordKey = NVLI(m_Rs("RC"), 0)
   End If

   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   CountRecordKey = False
End Function

'======================== END OF GENERIC FUNCTION HERE ======================
