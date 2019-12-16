VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public KEY_ID As Long
Public KEY_CODE As String
Public KEY_NAME As String
Public MASTER_AREA As Long
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM MASTER_REF WHERE (KEY_ID = " & KEY_ID & ")"
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT MR.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      
      WhereStr = ""
   End If
   
   If KEY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.KEY_ID = " & KEY_ID & ") " & vbCrLf
   End If
   
   If MASTER_AREA > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.MASTER_AREA = " & MASTER_AREA & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 2 Then
         OrderByStr = "ORDER BY KEY_NAME "
      Else
         OrderByStr = "ORDER BY KEY_CODE "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_CODE = NVLS(Rs("KEY_CODE"), "")
      KEY_NAME = NVLS(Rs("KEY_NAME"), "")
      MASTER_AREA = NVLI(Rs("MASTER_AREA"), -1)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

'   TempID1 = PARENT_ID
'   If PARENT_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("MASTER_REF_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO MASTER_REF " & _
                     "( " & _
                     "   KEY_CODE, " & _
                     "   KEY_NAME, " & _
                     "   MASTER_AREA, " & _
                     "   CREATE_DATE, " & _
                    "    MODIFY_DATE " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         "'" & ChangeQuote(Trim(KEY_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(KEY_NAME)) & "'," & _
                        MASTER_AREA & "," & _
                        "'" & ServerDateTime & "'," & _
                         "'" & ServerDateTime & "'" & _
                     ") "
         KEY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MASTER_REF SET " & _
                     "    KEY_CODE = '" & ChangeQuote(Trim(KEY_CODE)) & "', " & _
                     "    KEY_NAME = '" & ChangeQuote(Trim(KEY_NAME)) & "', " & _
                     "    MASTER_AREA = " & MASTER_AREA & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "' " & _
                     "WHERE KEY_ID = " & KEY_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
