VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChequeConfig"
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
Public CHEQUE_NAME As String
Public DATE_LEFT As Double
Public DATE_TOP As Double
Public DATE_TYPE As Long
Public PAY_LEFT As Double
Public PAY_TOP As Double
Public THAI_LEFT As Double
Public THAI_TOP As Double
Public AMOUNT_LEFT As Double
Public AMOUNT_TOP As Double

Public LOGO1_FLAG As String
Public LOGO1_LEFT As Double
Public LOGO1_TOP As Double
Public LOGO1_PATH As String

Public LOGO2_FLAG As String
Public LOGO2_LEFT As Double
Public LOGO2_TOP As Double
Public LOGO2_PATH As String

Public LOGO1_HEIGHT As Double
Public LOGO1_WEIGHT As Double
Public LOGO2_HEIGHT As Double
Public LOGO2_WEIGHT As Double

Public NAME1_DESC As String
Public NAME1_DESC_LEFT As Double
Public NAME1_DESC_TOP As Double

Public NAME2_DESC As String
Public NAME2_DESC_LEFT As Double
Public NAME2_DESC_TOP As Double

Public AC_LEFT As Double
Public AC_TOP As Double
Public BEARER_LEFT As Double
Public BEARER_TOP As Double

Public BBANK_ID As Long
Public BBANK_NAME As String
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
   
   SelectStr1 = "SELECT UG.*,MR1.KEY_NAME BBANK_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM CHEQUE_CONFIG UG "
   FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (UG.BBANK_ID = MR1.KEY_ID) "
   
   WhereStr = ""
   
   If Len(Trim(CHEQUE_NAME)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.CHEQUE_NAME = '" & ChangeQuote(Trim(CHEQUE_NAME)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.CHEQUE_NAME = '" & ChangeQuote(Trim(CHEQUE_NAME)) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   OrderByStr = "ORDER BY UG.CHEQUE_NAME "
   
   OrderTypeStr = "DESC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   CHEQUE_NAME = NVLS(Rs("CHEQUE_NAME"), "")
   DATE_LEFT = NVLD(Rs("DATE_LEFT"), 0)
   DATE_TOP = NVLD(Rs("DATE_TOP"), 0)
   DATE_TYPE = NVLI(Rs("DATE_TYPE"), 1)
   PAY_LEFT = NVLD(Rs("PAY_LEFT"), 0)
   PAY_TOP = NVLD(Rs("PAY_TOP"), 0)
   THAI_LEFT = NVLD(Rs("THAI_LEFT"), 0)
   THAI_TOP = NVLD(Rs("THAI_TOP"), 0)
   AMOUNT_LEFT = NVLD(Rs("AMOUNT_LEFT"), 0)
   AMOUNT_TOP = NVLD(Rs("AMOUNT_TOP"), 0)
   
   LOGO1_FLAG = NVLS(Rs("LOGO1_FLAG"), "")
   LOGO1_LEFT = NVLD(Rs("LOGO1_LEFT"), 0)
   LOGO1_TOP = NVLD(Rs("LOGO1_TOP"), 0)
   LOGO1_PATH = NVLS(Rs("LOGO1_PATH"), "")
   
   LOGO2_FLAG = NVLS(Rs("LOGO2_FLAG"), "")
   LOGO2_LEFT = NVLD(Rs("LOGO2_LEFT"), 0)
   LOGO2_TOP = NVLD(Rs("LOGO2_TOP"), 0)
   LOGO2_PATH = NVLS(Rs("LOGO2_PATH"), "")
   
   LOGO1_HEIGHT = NVLD(Rs("LOGO1_HEIGHT"), 0)
   LOGO1_WEIGHT = NVLD(Rs("LOGO1_WEIGHT"), 0)
   LOGO2_HEIGHT = NVLD(Rs("LOGO2_HEIGHT"), 0)
   LOGO2_WEIGHT = NVLD(Rs("LOGO2_WEIGHT"), 0)
   
   NAME1_DESC = NVLS(Rs("NAME1_DESC"), "")
   NAME1_DESC_LEFT = NVLD(Rs("NAME1_DESC_LEFT"), 0)
   NAME1_DESC_TOP = NVLD(Rs("NAME1_DESC_TOP"), 0)
   
   NAME2_DESC = NVLS(Rs("NAME2_DESC"), "")
   NAME2_DESC_LEFT = NVLD(Rs("NAME2_DESC_LEFT"), 0)
   NAME2_DESC_TOP = NVLD(Rs("NAME2_DESC_TOP"), 0)
   
   AC_LEFT = NVLD(Rs("AC_LEFT"), 0)
   AC_TOP = NVLD(Rs("AC_TOP"), 0)
   BEARER_LEFT = NVLD(Rs("BEARER_LEFT"), 0)
   BEARER_TOP = NVLD(Rs("BEARER_TOP"), 0)
   
   PREV_NAME = NVLS(Rs("CHEQUE_NAME"), "")
   
   BBANK_ID = NVLI(Rs("BBANK_ID"), -1)
   
   BBANK_NAME = NVLS(Rs("BBANK_NAME"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   TempID1 = BBANK_ID
   If BBANK_ID <= 0 Then
      TempID1 = "NULL"
   End If

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO CHEQUE_CONFIG " & _
                     "( " & _
                     "   CHEQUE_NAME, " & _
                     "   DATE_LEFT,DATE_TOP,DATE_TYPE, " & _
                     "   PAY_LEFT,PAY_TOP, " & _
                     "   THAI_LEFT,THAI_TOP, " & _
                     "   AMOUNT_LEFT,AMOUNT_TOP, " & _
                     "   LOGO1_FLAG,LOGO1_PATH, " & _
                     "   LOGO1_LEFT,LOGO1_TOP, " & _
                     "   LOGO2_FLAG,LOGO2_PATH, " & _
                     "   LOGO2_LEFT,LOGO2_TOP, " & _
                     "   LOGO1_HEIGHT,LOGO1_WEIGHT, " & _
                     "   LOGO2_HEIGHT,LOGO2_WEIGHT, " & _
                     "   NAME1_DESC,NAME1_DESC_LEFT,NAME1_DESC_TOP, " & _
                     "   NAME2_DESC,NAME2_DESC_LEFT,NAME2_DESC_TOP, " & _
                     "   BBANK_ID, " & _
                     "   AC_LEFT,AC_TOP, " & _
                     "   BEARER_LEFT,BEARER_TOP, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(CHEQUE_NAME)) & "'," & _
                     DATE_LEFT & "," & DATE_TOP & "," & DATE_TYPE & "," & _
                     PAY_LEFT & "," & PAY_TOP & "," & _
                     THAI_LEFT & "," & THAI_TOP & "," & _
                     AMOUNT_LEFT & "," & AMOUNT_TOP & "," & _
                     "'" & ChangeQuote(Trim(LOGO1_FLAG)) & "'," & "'" & ChangeQuote(Trim(LOGO1_PATH)) & "'," & _
                     LOGO1_LEFT & "," & LOGO1_TOP & "," & _
                     "'" & ChangeQuote(Trim(LOGO2_FLAG)) & "'," & "'" & ChangeQuote(Trim(LOGO2_PATH)) & "'," & _
                     LOGO2_LEFT & "," & LOGO2_TOP & "," & _
                     LOGO1_HEIGHT & "," & LOGO1_WEIGHT & "," & _
                     LOGO2_HEIGHT & "," & LOGO2_WEIGHT & "," & _
                     "'" & ChangeQuote(Trim(NAME1_DESC)) & "'," & "" & NAME1_DESC_LEFT & "," & "" & NAME1_DESC_TOP & "," & _
                     "'" & ChangeQuote(Trim(NAME2_DESC)) & "'," & "" & NAME2_DESC_LEFT & "," & "" & NAME2_DESC_TOP & "," & _
                     TempID1 & "," & _
                     AC_LEFT & "," & AC_TOP & "," & _
                     BEARER_LEFT & "," & BEARER_TOP & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE CHEQUE_CONFIG SET " & _
                        "   CHEQUE_NAME = '" & ChangeQuote(Trim(CHEQUE_NAME)) & "', " & _
                        "   DATE_LEFT = " & DATE_LEFT & ", " & "   DATE_TOP = " & DATE_TOP & ", " & _
                        "   DATE_TYPE = " & DATE_TYPE & ", " & _
                        "   PAY_LEFT = " & PAY_LEFT & ", " & "   PAY_TOP = " & PAY_TOP & ", " & _
                        "   THAI_LEFT = " & THAI_LEFT & ", " & "   THAI_TOP = " & THAI_TOP & ", " & _
                        "   AMOUNT_LEFT = " & AMOUNT_LEFT & ", " & _
                        "   AMOUNT_TOP = " & AMOUNT_TOP & ", " & _
                        "   LOGO1_FLAG = '" & ChangeQuote(Trim(LOGO1_FLAG)) & "', " & "   LOGO1_PATH = '" & ChangeQuote(Trim(LOGO1_PATH)) & "', " & _
                        "   LOGO1_LEFT = " & LOGO1_LEFT & ", " & " LOGO1_TOP = " & LOGO1_TOP & ", " & _
                        "   LOGO2_FLAG = '" & ChangeQuote(Trim(LOGO2_FLAG)) & "', " & "   LOGO2_PATH = '" & ChangeQuote(Trim(LOGO2_PATH)) & "', " & _
                        "   LOGO2_LEFT = " & LOGO2_LEFT & ", " & " LOGO2_TOP = " & LOGO2_TOP & ", " & _
                        "   LOGO1_HEIGHT = " & LOGO1_HEIGHT & ", " & " LOGO1_WEIGHT = " & LOGO1_WEIGHT & ", " & _
                        "   LOGO2_HEIGHT = " & LOGO2_HEIGHT & ", " & " LOGO2_WEIGHT = " & LOGO2_WEIGHT & ", " & _
                        "   NAME1_DESC = '" & ChangeQuote(Trim(NAME1_DESC)) & "', " & _
                        "   NAME2_DESC = '" & ChangeQuote(Trim(NAME2_DESC)) & "', " & _
                        "   NAME1_DESC_LEFT = " & NAME1_DESC_LEFT & ", " & " NAME1_DESC_TOP = " & NAME1_DESC_TOP & ", " & _
                        "   NAME2_DESC_LEFT = " & NAME2_DESC_LEFT & ", " & " NAME2_DESC_TOP = " & NAME2_DESC_TOP & ", " & _
                        "   BBANK_ID = " & TempID1 & ", " & _
                        "  AC_LEFT = " & AC_LEFT & ", " & " AC_TOP = " & AC_TOP & ", " & _
                        "  BEARER_LEFT = " & BEARER_LEFT & ", " & " BEARER_TOP = " & BEARER_TOP & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE CHEQUE_NAME = '" & ChangeQuote(Trim(PREV_NAME)) & "'"
   End If
   
   
   Call m_Conn.Execute(SQL1)
   
End Function

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM CHEQUE_CONFIG WHERE CHEQUE_NAME = '" & CHEQUE_NAME & "'"
      m_Conn.Execute (SQL1)
End Function
