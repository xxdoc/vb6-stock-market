VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub
Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReportConfig = False
End Function

Public Function DeleteMasterRef(Ug As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteMasterRef"
   DeleteMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteMasterRef = True
End Function
Public Function AddEditMasterRef(Ua As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterRef"
   AddEditMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterRef = False
End Function
Public Function QueryChequeConfig(Ua As CChequeConfig, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
   
   RName = "QueryChequeConfig"
   QueryChequeConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryChequeConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryChequeConfig = False
End Function
Public Function DeleteChequeConfig(Ug As CChequeConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteChequeConfig"
   DeleteChequeConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteChequeConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteChequeConfig = True
End Function
Public Function AddEditChequeConfig(Ug As CChequeConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditChequeConfig"
   AddEditChequeConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditChequeConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditChequeConfig = False
End Function
Public Function QueryStockMarket(Ua As CStockMarket, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
   
   RName = "QueryStockMarket"
   QueryStockMarket = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CStockMarketDetail
      Set Gr = New CStockMarketDetail
      Gr.STOCK_MARKET_ID = Ua.STOCK_MARKET_ID
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing
      
      Set Ua.StockMarketDetails = Nothing
      Set Ua.StockMarketDetails = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CStockMarketDetail
         Call Gr.PopulateFromRS(1, m_Rs1)
         Call Ua.StockMarketDetails.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryStockMarket = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryStockMarket = False
End Function
Public Function QueryUserGroup(Ua As CUserGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean

   RName = "QueryUserGroup"
   QueryUserGroup = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CGroupRight
      Set Gr = New CGroupRight
      Gr.GROUP_ID = Ua.GROUP_ID
      Call Gr.QueryData(2, m_Rs1, iCount)
      Set Gr = Nothing
      
      Set Ua.RightItems = Nothing
      Set Ua.RightItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CGroupRight
         Call Gr.PopulateFromRS(2, m_Rs1)
         Call Ua.RightItems.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryUserGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserGroup = False
End Function

Public Function AddEditUserGroup(Ua As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserGroup"
   AddEditUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CGroupRight
   For Each Gr In Ua.RightItems
      Gr.GROUP_ID = Ua.GROUP_ID
      If Gr.Flag = "A" Then
         Gr.AddEditMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.AddEditMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserGroup = False
End Function

Public Function DeleteUserGroup(Ug As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteUserGroup"
   DeleteUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   
   IsOK = False

   DeleteUserGroup = True
End Function
Public Function DeleteUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteUserAccount"
   DeleteUserAccount = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteUserAccount = True
End Function
Public Function AddEditUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserAccount"
   AddEditUserAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserAccount = False
End Function
Public Function QueryUserAccount(Ua As CUserAccount, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryUserAccount"
   QueryUserAccount = False

   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryUserAccount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserAccount = False
End Function
Public Function CopyChequeConfig(Ug As CChequeConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyChequeConfig"
   CopyChequeConfig = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryChequeConfig(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.CHEQUE_NAME = "***" & Ug.CHEQUE_NAME
       
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditChequeConfig(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyChequeConfig = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   CopyChequeConfig = False
End Function

Public Function DBLogin(UserName As String, Password As String, CanLogin As Boolean, User As clsUser, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim ItemCount As Long
Dim LoginID As Long

   RName = "Login"
   DBLogin = False
'
   SQL1 = "SELECT COUNT(*) FROM USER_ACCOUNT "
   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   If NVLI(m_Rs1(0), 0) <= 0 Then 'ถ้าไม่มีผู้ใช้อยุในดาตาเบสให้ล็อคอินอะไรก็ได้
      If UserName <> "" Then
         ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

         DBLogin = True
         CanLogin = False
         Exit Function
      End If

      User.GROUP_ID = 0
      User.GROUP_NAME = ""
      User.USER_NAME = UserName

      DBLogin = True
      CanLogin = True
      Exit Function
   End If

   SQL1 = "SELECT UD.*, GD.* FROM USER_ACCOUNT UD, USER_GROUP GD "
   SQL1 = SQL1 & "WHERE (GD.GROUP_ID = UD.GROUP_ID) "
   SQL1 = SQL1 & "AND (UD.USER_NAME = '" & ChangeQuote(UserName) & "' ) "
   SQL1 = SQL1 & "AND (UD.USER_PASSWORD = '" & ChangeQuote(EncryptText(Password)) & "' ) "

   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

      DBLogin = True
      CanLogin = False
      Exit Function
   End If

   If NVLS(m_Rs1("USER_STATUS"), "Y") <> "Y" Then
      ErrorObj.LocalErrorMsg = "รายชื่อของท่านขณะนี้ไม่สามารถเข้าใช้งานระบบได้ โปรดติดต่อผู้บริหารระบบ "

      DBLogin = True
      CanLogin = False
      Exit Function
   End If

   User.GROUP_ID = NVLI(m_Rs1("GROUP_ID"), -1)
   User.GROUP_NAME = NVLS(m_Rs1("GROUP_NAME"), "")
   User.USER_NAME = UserName
   
   
   CanLogin = True
   DBLogin = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   DBLogin = False
End Function
Public Function UpDateUserStatus(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "UpDateUserStatus"
   UpDateUserStatus = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.UpDateStatus
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   UpDateUserStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   UpDateUserStatus = False
End Function
Public Function AddEditStockMarket(Ua As CStockMarket, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditStockMarket"
   AddEditStockMarket = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Smd As CStockMarketDetail
   For Each Smd In Ua.StockMarketDetails
      Smd.STOCK_MARKET_ID = Ua.STOCK_MARKET_ID
      If Smd.Flag = "A" Then
         Smd.AddEditMode = SHOW_ADD
         Call Smd.AddEditData
      ElseIf Smd.Flag = "E" Then
         Smd.AddEditMode = SHOW_EDIT
         Call Smd.AddEditData
      ElseIf Smd.Flag = "D" Then
         Call Smd.DeleteData
      End If
   Next Smd
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditStockMarket = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditStockMarket = False
End Function
Public Function DeleteStockMarket(Ug As CStockMarket, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteStockMarket"
   DeleteStockMarket = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteStockMarket = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteStockMarket = True
End Function

