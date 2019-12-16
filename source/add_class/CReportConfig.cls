VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportConfig"
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

Public REPORT_CONFIG_ID As Long
Public REPORT_KEY As String
Public PAPER_SIZE As Double
Public PAPER_WIDTH As Double
Public PAPER_HEIGHT As Double
Public ORIENTATION As Double
Public MARGIN_BOTTOM As Double
Public MARGIN_FOOTER As Double
Public MARGIN_HEADER As Double
Public MARGIN_LEFT As Double
Public MARGIN_RIGHT As Double
Public MARGIN_TOP As Double
Public FONT_NAME As String
Public FONT_SIZE As Double
Public COMPUTER_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM REPORT_CONFIG UG "
   
   WhereStr = ""
   If REPORT_CONFIG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (REPORT_CONFIG_ID = " & REPORT_CONFIG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (REPORT_CONFIG_ID = " & REPORT_CONFIG_ID & ") "
      End If
   End If
   
   If Len(REPORT_KEY) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (REPORT_KEY = '" & ChangeQuote(Trim(REPORT_KEY)) & "') "
      Else
         WhereStr = WhereStr & "AND (REPORT_KEY = '" & ChangeQuote(Trim(REPORT_KEY)) & "') "
      End If
   End If
   
   If Len(COMPUTER_NAME) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMPUTER_NAME = '" & ChangeQuote(Trim(COMPUTER_NAME)) & "') "
      Else
         WhereStr = WhereStr & "AND (COMPUTER_NAME = '" & ChangeQuote(Trim(COMPUTER_NAME)) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY REPORT_CONFIG_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PAPER_SIZE "
   Else
      OrderByStr = "ORDER BY REPORT_CONFIG_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   REPORT_CONFIG_ID = NVLI(Rs("REPORT_CONFIG_ID"), -1)
   REPORT_KEY = NVLS(Rs("REPORT_KEY"), "")
   PAPER_SIZE = NVLD(Rs("PAPER_SIZE"), 0)
   PAPER_WIDTH = NVLD(Rs("PAPER_WIDTH"), 0)
   PAPER_HEIGHT = NVLD(Rs("PAPER_HEIGHT"), 0)
   ORIENTATION = NVLD(Rs("ORIENTATION"), 0)
   MARGIN_BOTTOM = NVLD(Rs("MARGIN_BOTTOM"), 0)
   MARGIN_FOOTER = NVLD(Rs("MARGIN_FOOTER"), 0)
   MARGIN_HEADER = NVLD(Rs("MARGIN_HEADER"), 0)
   MARGIN_LEFT = NVLD(Rs("MARGIN_LEFT"), 0)
   MARGIN_RIGHT = NVLD(Rs("MARGIN_RIGHT"), 0)
   MARGIN_TOP = NVLD(Rs("MARGIN_TOP"), 0)
   FONT_NAME = NVLS(Rs("FONT_NAME"), "")
   FONT_SIZE = NVLD(Rs("FONT_SIZE"), 0)
   COMPUTER_NAME = NVLS(Rs("COMPUTER_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO REPORT_CONFIG " & _
                     "( " & _
                     "   REPORT_KEY, " & _
                     "   PAPER_SIZE, " & _
                     "   PAPER_WIDTH, " & _
                     "   PAPER_HEIGHT, " & _
                     "   ORIENTATION, " & _
                     "   MARGIN_BOTTOM, " & _
                     "   MARGIN_FOOTER, " & _
                     "   MARGIN_HEADER, " & _
                     "   MARGIN_LEFT, " & _
                     "   MARGIN_RIGHT, " & _
                     "   MARGIN_TOP, " & _
                     "   FONT_NAME, "
         SQL1 = SQL1 & _
                     "   FONT_SIZE, " & _
                     "   COMPUTER_NAME " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         "'" & ChangeQuote(Trim(REPORT_KEY)) & "'," & _
                         PAPER_SIZE & "," & _
                         PAPER_WIDTH & "," & _
                         PAPER_HEIGHT & "," & _
                         ORIENTATION & "," & _
                         MARGIN_BOTTOM & "," & _
                         MARGIN_FOOTER & "," & _
                         MARGIN_HEADER & "," & _
                         MARGIN_LEFT & "," & _
                         MARGIN_RIGHT & ","
               SQL1 = SQL1 & _
                         MARGIN_TOP & "," & _
                         "'" & ChangeQuote(Trim(FONT_NAME)) & "'," & _
                         FONT_SIZE & "," & _
                         "'" & ChangeQuote(Trim(COMPUTER_NAME)) & "'" & _
                     ") "
         REPORT_CONFIG_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE REPORT_CONFIG SET " & _
                     "   REPORT_KEY = '" & ChangeQuote(Trim(REPORT_KEY)) & "', " & _
                     "   PAPER_SIZE = " & PAPER_SIZE & ", " & _
                     "   PAPER_WIDTH = " & PAPER_WIDTH & ", " & _
                     "   PAPER_HEIGHT = " & PAPER_HEIGHT & ", " & _
                     "   ORIENTATION = " & ORIENTATION & ", " & _
                     "   MARGIN_BOTTOM = " & MARGIN_BOTTOM & ", " & _
                     "   MARGIN_FOOTER = " & MARGIN_FOOTER & ", " & _
                     "   MARGIN_HEADER = " & MARGIN_HEADER & ", " & _
                     "   MARGIN_LEFT = " & MARGIN_LEFT & ", " & _
                     "   MARGIN_RIGHT = " & MARGIN_RIGHT & ", " & _
                     "   MARGIN_TOP = " & MARGIN_TOP & ", " & _
                     "   FONT_NAME = '" & ChangeQuote(Trim(FONT_NAME)) & "', " & _
                     "   COMPUTER_NAME = '" & ChangeQuote(Trim(COMPUTER_NAME)) & "', " & _
                     "   FONT_SIZE = " & FONT_SIZE & " " & _
                     "WHERE REPORT_CONFIG_ID = " & REPORT_CONFIG_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM REPORT_CONFIG WHERE REPORT_CONFIG_ID = " & REPORT_CONFIG_ID
      m_Conn.Execute (SQL1)
End Function


