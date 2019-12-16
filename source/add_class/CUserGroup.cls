VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserGroup"
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

Public GROUP_ID As Long
Public GROUP_NAME As String
Public GROUP_DESC As String
Public GROUP_STATUS As String
Public MAX_USER As Long

Public OrderBy  As Long
Public OrderType As Long
Public RightItems As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
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
   
      FromStr = "FROM USER_GROUP UG "
   End If
   WhereStr = ""
   
   If GROUP_ID >= 0 Then    ' ต้องเป็น >= 0 เท่านั้นเพื่อที่เวลา Query ออกมาจากแสดงได้ถูกต้อง
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.GROUP_ID = " & GROUP_ID & ")"
      Else
         WhereStr = WhereStr & "AND (UG.GROUP_ID = " & GROUP_ID & ")"
      End If
   End If
   
   If Len(Trim(GROUP_STATUS)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.GROUP_STATUS = '" & ChangeQuote(Trim(GROUP_STATUS)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.GROUP_STATUS = '" & ChangeQuote(Trim(GROUP_STATUS)) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY UG.GROUP_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY UG.GROUP_ID "
   Else
      OrderByStr = "ORDER BY UG.GROUP_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
      GROUP_NAME = NVLS(Rs("GROUP_NAME"), "")
      GROUP_DESC = NVLS(Rs("GROUP_DESC"), "")
      GROUP_STATUS = NVLS(Rs("GROUP_STATUS"), "")
      MAX_USER = NVLI(Rs("MAX_USER"), -1)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO USER_GROUP " & _
                     "( " & _
                     "   GROUP_NAME, " & _
                     "   GROUP_DESC, " & _
                     "   GROUP_STATUS, " & _
                     "   MAX_USER, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(GROUP_NAME)) & "'," & _
                     "'" & ChangeQuote(Trim(GROUP_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(GROUP_STATUS)) & "'," & _
                     MAX_USER & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING GROUP_ID "
   Else
            SQL1 = "UPDATE USER_GROUP SET " & _
                        "   GROUP_NAME = '" & ChangeQuote(Trim(GROUP_NAME)) & "', " & _
                        "   GROUP_DESC = '" & ChangeQuote(Trim(GROUP_DESC)) & "', " & _
                        "   GROUP_STATUS = '" & ChangeQuote(Trim(GROUP_STATUS)) & "', " & _
                        "    MAX_USER = " & MAX_USER & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE GROUP_ID = " & GROUP_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      GROUP_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM GROUP_RIGHT WHERE GROUP_ID = " & GROUP_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM USER_GROUP WHERE GROUP_ID = " & GROUP_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set RightItems = Nothing
End Sub
