VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPayName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public PREV_NAME As String
Public PAY_NAME As String
Public PAY_NAME_CHANGE As String

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempStr  As String
Dim SubLen  As Long
Dim NewStr  As String
   
   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PAY_NAME UG "
   
   WhereStr = ""
   
   If Len(Trim(PAY_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(PAY_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(UG.PAY_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (UG.PAY_NAME = '" & ChangeQuote(Trim(PAY_NAME)) & "')"
      End If
   End If
   
   If glbUser.ENTERPRISE_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.ENTERPRISE_ID = " & glbUser.ENTERPRISE_ID & ")"
      Else
         WhereStr = WhereStr & "AND (UG.ENTERPRISE_ID = " & glbUser.ENTERPRISE_ID & ")"
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   OrderByStr = "ORDER BY UG.PAY_NAME "
   
   OrderTypeStr = "DESC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   PAY_NAME = NVLS(Rs("PAY_NAME"), "")
   
   PREV_NAME = NVLS(Rs("PAY_NAME"), "")
   PAY_NAME_CHANGE = NVLS(Rs("PAY_NAME_CHANGE"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PAY_NAME " & _
                     "( " & _
                     "   PAY_NAME, " & _
                     "   PAY_NAME_CHANGE, " & _
                     "   ENTERPRISE_ID, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(PAY_NAME)) & "'," & _
                     "'" & ChangeQuote(Trim(PAY_NAME_CHANGE)) & "'," & _
                     "" & glbUser.ENTERPRISE_ID & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE PAY_NAME SET " & _
                        "   PAY_NAME = '" & ChangeQuote(Trim(PAY_NAME)) & "', " & _
                        "   PAY_NAME_CHANGE = '" & ChangeQuote(Trim(PAY_NAME_CHANGE)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE PAY_NAME = '" & ChangeQuote(Trim(PREV_NAME)) & "'"
   End If
   
   
   Call m_Conn.Execute(SQL1)
   
End Function

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PAY_NAME WHERE PAY_NAME = '" & PAY_NAME & "'"
      m_Conn.Execute (SQL1)
End Function
