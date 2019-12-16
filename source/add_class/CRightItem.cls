VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRightItem"
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

Public RIGHT_ID As Long
Public RIGHT_ITEM_DESC As String
Public DEFAULT_VALUE As String
Public RIGHT_ITEM_NAME As String
Public PARENT_ID As Long

Public OrderBy  As Long
Public OrderType As Long
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
   
      FromStr = "FROM RIGHT_ITEM UG "
   End If
   WhereStr = ""
   
   If RIGHT_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.RIGHT_ID = " & RIGHT_ID & ")"
      Else
         WhereStr = WhereStr & "AND (UG.RIGHT_ID = " & RIGHT_ID & ")"
      End If
   End If
   
   If Len(Trim(RIGHT_ITEM_NAME)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.RIGHT_ITEM_NAME = '" & ChangeQuote(Trim(RIGHT_ITEM_NAME)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.RIGHT_ITEM_NAME = '" & ChangeQuote(Trim(RIGHT_ITEM_NAME)) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY UG.RIGHT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY UG.RIGHT_ID "
   Else
      OrderByStr = "ORDER BY UG.RIGHT_ID "
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
      RIGHT_ID = NVLI(Rs("RIGHT_ID"), -1)
      RIGHT_ITEM_DESC = NVLS(Rs("RIGHT_ITEM_DESC"), "")
      DEFAULT_VALUE = NVLS(Rs("DEFAULT_VALUE"), "")
      RIGHT_ITEM_NAME = NVLS(Rs("RIGHT_ITEM_NAME"), "")
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim TempID1 As String
   
   TempID1 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO RIGHT_ITEM " & _
                     "( " & _
                     "   RIGHT_ITEM_DESC, " & _
                     "   DEFAULT_VALUE, " & _
                     "   RIGHT_ITEM_NAME, " & _
                     "   PARENT_ID, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(RIGHT_ITEM_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(DEFAULT_VALUE)) & "'," & _
                     "'" & ChangeQuote(Trim(RIGHT_ITEM_NAME)) & "'," & _
                     TempID1 & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING RIGHT_ID "
   Else
            SQL1 = "UPDATE RIGHT_ITEM SET " & _
                        "   RIGHT_ITEM_DESC = '" & ChangeQuote(Trim(RIGHT_ITEM_DESC)) & "', " & _
                        "   DEFAULT_VALUE = '" & ChangeQuote(Trim(DEFAULT_VALUE)) & "', " & _
                        "   RIGHT_ITEM_NAME = '" & ChangeQuote(Trim(RIGHT_ITEM_NAME)) & "', " & _
                        "    PARENT_ID = " & TempID1 & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE RIGHT_ID = " & RIGHT_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      RIGHT_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM RIGHT_ITEM WHERE RIGHT_ID = " & RIGHT_ID
      m_Conn.Execute (SQL1)
End Function


