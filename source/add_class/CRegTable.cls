VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public Flag As String
' End of generic data member

Public REGISTER_KEY As String
Public ACTIVATED_KEY As String
Public TIME_USED As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM REGTABLE UG "
   End If
   WhereStr = ""
   
   If Len(Trim(REGISTER_KEY)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.REGISTER_KEY = '" & ChangeQuote(Trim(REGISTER_KEY)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.REGISTER_KEY = '" & ChangeQuote(Trim(REGISTER_KEY)) & "')"
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      REGISTER_KEY = NVLS(Rs("REGISTER_KEY"), "")
      ACTIVATED_KEY = NVLS(Rs("ACTIVATED_KEY"), "")
      TIME_USED = NVLI(Rs("TIME_USED"), -1)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
      SQL1 = "INSERT INTO REGTABLE " & _
                     "( " & _
                     "   REGISTER_KEY, " & _
                     "   ACTIVATED_KEY, " & _
                     "   TIME_USED  " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(REGISTER_KEY)) & "'," & _
                     "'" & ChangeQuote(Trim(ACTIVATED_KEY)) & "'," & _
                     TIME_USED & "" & _
                     ") "
   
   Call m_Conn.Execute(SQL1)  'FOR ADD ONLY
   
End Function
Public Function UpdateTime()
Dim SQL1 As String
   
   SQL1 = "UPDATE REGTABLE SET " & _
                        "    TIME_USED = " & TIME_USED & " " & _
                        " WHERE REGISTER_KEY = '" & REGISTER_KEY & "'"
   
   Call m_Conn.Execute(SQL1)
   
End Function
Public Function UpdateAccKey()
Dim SQL1 As String
   
   SQL1 = "UPDATE REGTABLE SET " & _
                        "   ACTIVATED_KEY = '" & ChangeQuote(Trim(ACTIVATED_KEY)) & "'" & _
                        " WHERE REGISTER_KEY = '" & REGISTER_KEY & "'"
                     
   Call m_Conn.Execute(SQL1)
   
End Function


