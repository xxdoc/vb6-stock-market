VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserAccount"
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

Public USER_NAME As String
Public USER_PASSWORD As String
Public USER_DESC As String
Public USER_STATUS As String
Public GROUP_ID As Long

Public GROUP_NAME As String

Public PREV_NAME  As String
Public OrderBy  As Long
Public OrderType As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim TempID1 As String
   
   TempID1 = GROUP_ID
   If GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO USER_ACCOUNT " & _
                     "( " & _
                     "   USER_NAME, " & _
                     "   USER_PASSWORD, " & _
                     "   USER_DESC, " & _
                     "   USER_STATUS, " & _
                     "   GROUP_ID, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(USER_NAME)) & "'," & _
                     "'" & ChangeQuote(Trim(USER_PASSWORD)) & "'," & _
                     "'" & ChangeQuote(Trim(USER_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(USER_STATUS)) & "'," & _
                     TempID1 & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE USER_ACCOUNT SET " & _
                        "   USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "', " & _
                        "   USER_PASSWORD = '" & ChangeQuote(Trim(USER_PASSWORD)) & "', " & _
                        "   USER_DESC = '" & ChangeQuote(Trim(USER_DESC)) & "', " & _
                        "   USER_STATUS = '" & ChangeQuote(Trim(USER_STATUS)) & "', " & _
                        "    GROUP_ID = " & TempID1 & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE USER_NAME = '" & PREV_NAME & "'"
   End If
   
   Call m_Conn.Execute(SQL1)
   
End Function
'In this class we will not handle any error or any transaction
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
      SelectStr1 = SelectStr1 & ",UP.GROUP_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM USER_ACCOUNT UG "
      FromStr = FromStr & "LEFT OUTER JOIN USER_GROUP UP ON (UG.GROUP_ID = UP.GROUP_ID) "
   End If
   
   WhereStr = ""
   
   If Len(Trim(USER_NAME)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   OrderByStr = "ORDER BY UG.USER_NAME "
   
   OrderTypeStr = "DESC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      USER_NAME = NVLS(Rs("USER_NAME"), "")
      USER_PASSWORD = NVLS(Rs("USER_PASSWORD"), "")
      USER_DESC = NVLS(Rs("USER_DESC"), "")
      GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
      USER_STATUS = NVLS(Rs("USER_STATUS"), "")
            
      GROUP_NAME = NVLS(Rs("GROUP_NAME"), "")
      
      PREV_NAME = NVLS(Rs("USER_NAME"), "")
   End If
End Sub
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM USER_ACCOUNT WHERE USER_NAME = '" & USER_NAME & "'"
      m_Conn.Execute (SQL1)
End Function
Public Function UpDateStatus()
Dim ServerDateTime As String
Dim SQL1 As String
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
            SQL1 = "UPDATE USER_ACCOUNT SET " & _
                        "   USER_STATUS = '" & ChangeQuote(Trim(USER_STATUS)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE USER_NAME = '" & USER_NAME & "'"
   
   Call m_Conn.Execute(SQL1)
   
End Function

