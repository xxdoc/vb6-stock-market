VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupRight"
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

Public GROUP_RIGHT_ID As Long
Public GROUP_ID As Long
Public RIGHT_ID As Long
Public RIGHT_STATUS As String
Public PARENT_ID  As Long
Public RIGHT_ITEM_DESC As String
Public RIGHT_ITEM_NAME As String

Public OrderBy As Long
Public OrderType As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM GROUP_RIGHT WHERE GROUP_RIGHT_ID = " & GROUP_RIGHT_ID
      m_Conn.Execute (SQL1)
End Function
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
      SelectStr1 = "SELECT EN.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM RIGHT_ITEM EN "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT RI.*, EN.*, EN.RIGHT_ID RIGHT_ID2, EN.RIGHT_STATUS RIGHT_STATUS2, RI.RIGHT_ID RIGHT_ID1, RI.DEFAULT_VALUE RIGHT_STATUS1 "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM RIGHT_ITEM RI "
      FromStr = FromStr & "LEFT OUTER JOIN GROUP_RIGHT EN ON ((RI.RIGHT_ID = EN.RIGHT_ID) AND (EN.GROUP_ID = " & GROUP_ID & "))"
   
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT EN.*, EN.RIGHT_ID RIGHT_ID_ORG, RI.RIGHT_ITEM_NAME, RI.DEFAULT_VALUE  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM GROUP_RIGHT EN "
      FromStr = FromStr & "LEFT OUTER JOIN RIGHT_ITEM RI ON (EN.RIGHT_ID = RI.RIGHT_ID) "
   End If
   
   If GROUP_RIGHT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EN.GROUP_RIGHT_ID = " & GROUP_RIGHT_ID & ") " & vbCrLf
   End If
   
   If (GROUP_ID > 0) And (Ind <> 2) Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EN.GROUP_ID = " & GROUP_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY GROUP_ID "
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RI.RIGHT_ITEM_NAME "
      Else
         OrderByStr = "ORDER BY RI.RIGHT_ITEM_NAME "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY GROUP_ID "
   End If
   
   If OrderType = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      GROUP_RIGHT_ID = NVLI(Rs("GROUP_RIGHT_ID"), -1)
      GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
      RIGHT_ID = NVLI(Rs("RIGHT_ID"), -1)
      RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS"), "")
   ElseIf Ind = 2 Then
      GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
      GROUP_RIGHT_ID = NVLI(Rs("GROUP_RIGHT_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), 0)
      If NVLI(Rs("RIGHT_ID1"), 0) <> NVLI(Rs("RIGHT_ID2"), 0) Then
         RIGHT_ID = NVLI(Rs("RIGHT_ID1"), 0)
         RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS1"), "")
         Flag = "A"
      Else
         RIGHT_ID = NVLI(Rs("RIGHT_ID2"), 0)
         RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS2"), "")
         Flag = "I"
      End If
      RIGHT_ITEM_DESC = NVLS(Rs("RIGHT_ITEM_DESC"), "")
      RIGHT_ITEM_NAME = NVLS(Rs("RIGHT_ITEM_NAME"), "")
   ElseIf Ind = 3 Then
      GROUP_RIGHT_ID = NVLI(Rs("GROUP_RIGHT_ID"), -1)
      GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
      RIGHT_ID = NVLI(Rs("RIGHT_ID_ORG"), -1)
      RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS"), "")
      If RIGHT_ID <= 0 Then
         RIGHT_STATUS = "Y"
      End If
      RIGHT_ITEM_NAME = NVLS(Rs("RIGHT_ITEM_NAME"), "")
   End If
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = GROUP_ID
   If GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = RIGHT_ID
   If RIGHT_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO GROUP_RIGHT " & _
                     "( " & _
                     "   GROUP_ID, " & _
                     "   RIGHT_ID, " & _
                     "   RIGHT_STATUS, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     "'" & ChangeQuote(Trim(RIGHT_STATUS)) & "'," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE GROUP_RIGHT SET " & _
                        "    GROUP_ID = " & TempID1 & ", " & _
                        "    RIGHT_ID = " & TempID2 & ", " & _
                        "   RIGHT_STATUS = '" & ChangeQuote(Trim(RIGHT_STATUS)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE GROUP_RIGHT_ID = " & GROUP_RIGHT_ID
   End If
   
   Call m_Conn.Execute(SQL1)
   
End Function
