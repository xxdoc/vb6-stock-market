VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrintChequeLog"
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

Public PRINT_CHEQUE_ID As Long
Public PAYEE_NAME As String
Public CHEQUE_DATE As Date
Public CHEQUE_AMOUNT As Double
Public CHEQUE_TYPE As String
Public CHEQUE_NO As String

Public PASS_FLAG As String
Public CREATE_DATE As Date
Public CREATE_BY As String
Public MODIFY_DATE  As Date
Public BANK_DATE  As Date
Public MODIFY_BY As String
Public CANCEL_FLAG As String
Public CANCEL_DATE  As Date
Public CANCEL_BY As String
Public CHEQUE_YYYYMM As String
Public CANCEL_REASON As String

Public FROM_CHEQUE_DATE As Date
Public TO_CHEQUE_DATE As Date
Public FROM_PRINT_DATE As Date
Public TO_PRINT_DATE As Date
Public FROM_CHEQUE_NO As String
Public TO_CHEQUE_NO As String
Public FROM_BANK_DATE As Date
Public TO_BANK_DATE As Date

Public BBANK_ID As Long
Public BBANK_NAME  As String

Public ENTERPRISE_ID As Long
'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False, Optional AllEnterprise As Boolean)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim GroupByStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim TempStr As String
Dim SubLen As Long
Dim NewStr As String
   If Ind = 1 Or Ind = 2 Or Ind = 10 Or Ind = 11 Or Ind = 12 Then  'frmChequeLog    'CheckChequeNoCanPrint'CReportCheque001,'CReportCheque001_1,CReportCheque001_2,CReportCheque001_3
      SelectStr1 = "SELECT UG.*,MR1.KEY_NAME BBANK_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (CC.BBANK_ID = MR1.KEY_ID) "
      
   ElseIf Ind = 3 Then   ''CReportCheque002
      SelectStr1 = "SELECT SUM(UG.CHEQUE_AMOUNT) CHEQUE_AMOUNT "
      SelectStr1 = SelectStr1 & " ,UG.PAYEE_NAME,UG.CHEQUE_YYYYMM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
   ElseIf Ind = 4 Then     ''CReportCheque002,'CReportCheque003,'CReportCheque004
      SelectStr1 = "SELECT DISTINCT UG.PAYEE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
   ElseIf Ind = 5 Then   'CReportCheque002
      SelectStr1 = "SELECT DISTINCT UG.CHEQUE_YYYYMM "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
   ElseIf Ind = 6 Then     'CReportCheque003
      SelectStr1 = "SELECT SUM(UG.CHEQUE_AMOUNT) CHEQUE_AMOUNT "
      SelectStr1 = SelectStr1 & " ,UG.PAYEE_NAME,UG.CHEQUE_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
   ElseIf Ind = 7 Then        'CReportCheque003
      SelectStr1 = "SELECT DISTINCT UG.CHEQUE_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
   ElseIf Ind = 8 Then        'CReportCheque004
      SelectStr1 = "SELECT SUM(UG.CHEQUE_AMOUNT) CHEQUE_AMOUNT "
      SelectStr1 = SelectStr1 & " ,UG.PAYEE_NAME,CC.BBANK_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
   ElseIf Ind = 9 Then        'CReportCheque004
      SelectStr1 = "SELECT DISTINCT MR1.KEY_NAME BBANK_NAME,CC.BBANK_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM PRINT_CHEQUE_LOG UG "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_CONFIG CC ON (UG.CHEQUE_TYPE = CC.CHEQUE_NAME) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (CC.BBANK_ID = MR1.KEY_ID) "
      
   End If
   WhereStr = ""
   
   If Not AllEnterprise Then
      ENTERPRISE_ID = glbUser.ENTERPRISE_ID
   End If
   
   If Len(Trim(PAYEE_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(PAYEE_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(UG.PAYEE_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (UG.PAYEE_NAME = '" & ChangeQuote(Trim(PAYEE_NAME)) & "')"
      End If
   End If
   
   
   If Len(Trim(PASS_FLAG)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.PASS_FLAG = '" & ChangeQuote(Trim(PASS_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.PASS_FLAG = '" & ChangeQuote(Trim(PASS_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(CHEQUE_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_NO)) & "')"
      End If
   End If
   
   If Len(Trim(CHEQUE_TYPE)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.CHEQUE_TYPE = '" & ChangeQuote(Trim(CHEQUE_TYPE)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.CHEQUE_TYPE = '" & ChangeQuote(Trim(CHEQUE_TYPE)) & "')"
      End If
   End If
   
   If BBANK_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CC.BBANK_ID = " & BBANK_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CC.BBANK_ID = " & BBANK_ID & ")"
      End If
   End If
   
   If Len(Trim(FROM_CHEQUE_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.CHEQUE_NO >= '" & ChangeQuote(Trim(FROM_CHEQUE_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.CHEQUE_NO >= '" & ChangeQuote(Trim(FROM_CHEQUE_NO)) & "')"
      End If
   End If
   
   If Len(Trim(TO_CHEQUE_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.CHEQUE_NO <= '" & ChangeQuote(Trim(TO_CHEQUE_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.CHEQUE_NO <= '" & ChangeQuote(Trim(TO_CHEQUE_NO)) & "')"
      End If
   End If
   
   If FROM_CHEQUE_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_CHEQUE_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (UG.CHEQUE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_CHEQUE_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_CHEQUE_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (UG.CHEQUE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

   If FROM_PRINT_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_PRINT_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (UG.CREATE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_PRINT_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_PRINT_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (UG.CREATE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If FROM_BANK_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_BANK_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (UG.BANK_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_BANK_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_BANK_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (UG.BANK_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If ENTERPRISE_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.ENTERPRISE_ID = " & ENTERPRISE_ID & ")"
      Else
         WhereStr = WhereStr & "AND (UG.ENTERPRISE_ID = " & ENTERPRISE_ID & ")"
      End If
   End If
   
   If Ind = 3 Then
      GroupByStr = " GROUP BY UG.PAYEE_NAME,UG.CHEQUE_YYYYMM  "
   ElseIf Ind = 6 Then
      GroupByStr = " GROUP BY UG.PAYEE_NAME,UG.CHEQUE_DATE  "
   ElseIf Ind = 8 Then
      GroupByStr = " GROUP BY UG.PAYEE_NAME,CC.BBANK_ID  "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Or Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY UG.CREATE_DATE "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE "
      ElseIf OrderBy = 14 Then
         OrderByStr = "ORDER BY UG.CHEQUE_NO "
      ElseIf OrderBy = 15 Then
         OrderByStr = "ORDER BY UG.CHEQUE_AMOUNT "
      ElseIf OrderBy = 16 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE "
      Else
         OrderByStr = "ORDER BY UG.PRINT_CHEQUE_ID "
      End If
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY UG.PAYEE_NAME "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY UG.CHEQUE_YYYYMM "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY UG.CHEQUE_DATE "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE "
   ElseIf Ind = 10 Then      'CReportCheque001_1 สำหรับการจัดเรียงตามการจ่ายให้
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.CREATE_DATE "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.CHEQUE_DATE "
      ElseIf OrderBy = 14 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.CHEQUE_NO "
      ElseIf OrderBy = 15 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.CHEQUE_AMOUNT "
      ElseIf OrderBy = 16 Then
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,MR1.KEY_CODE "
      Else
         OrderByStr = "ORDER BY UG.PAYEE_NAME ASC,UG.PRINT_CHEQUE_ID "
      End If
   ElseIf Ind = 11 Then      'CReportCheque001_2 สำหรับการจัดเรียงตามสาขาธนาคาร
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.PAYEE_NAME "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.CREATE_DATE "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.CHEQUE_DATE "
      ElseIf OrderBy = 14 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.CHEQUE_NO "
      ElseIf OrderBy = 15 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.CHEQUE_AMOUNT "
      ElseIf OrderBy = 16 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE "
      Else
         OrderByStr = "ORDER BY MR1.KEY_CODE ASC,UG.PRINT_CHEQUE_ID "
      End If
   ElseIf Ind = 12 Then      'CReportCheque001_3 สำหรับการจัดเรียงตามวันที่เช็ค
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.PRINT_CHEQUE_ID "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.PAYEE_NAME "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.CREATE_DATE "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE "
      ElseIf OrderBy = 14 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.CHEQUE_NO "
      ElseIf OrderBy = 15 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.CHEQUE_AMOUNT "
      ElseIf OrderBy = 16 Then
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,MR1.KEY_CODE "
      Else
         OrderByStr = "ORDER BY UG.CHEQUE_DATE ASC,UG.PRINT_CHEQUE_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   If OrderBy = 2 Then
      OrderTypeStr = OrderTypeStr & "ROWS 1 TO 20 "     'สำหรับเอา 20 Record แรก
   End If
   If OrderByStr = "" Then
      OrderTypeStr = ""
   End If
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Or Ind = 2 Or Ind = 10 Or Ind = 11 Or Ind = 12 Then
      PRINT_CHEQUE_ID = NVLI(Rs("PRINT_CHEQUE_ID"), -1)
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      CHEQUE_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), ""))
      CHEQUE_TYPE = NVLS(Rs("CHEQUE_TYPE"), "")
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      
      CREATE_DATE = InternalDateToDate(NVLS(Rs("CREATE_DATE"), ""))
      CREATE_BY = NVLS(Rs("CREATE_BY"), "")
      
      PASS_FLAG = NVLS(Rs("PASS_FLAG"), "")
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
      MODIFY_BY = NVLS(Rs("MODIFY_BY"), "")
      
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      CANCEL_DATE = InternalDateToDate(NVLS(Rs("CANCEL_DATE"), ""))
      CANCEL_BY = NVLS(Rs("CANCEL_BY"), "")
      CANCEL_REASON = NVLS(Rs("CANCEL_REASON"), "")
      
      BBANK_NAME = NVLS(Rs("BBANK_NAME"), "")
      BANK_DATE = InternalDateToDate(NVLS(Rs("BANK_DATE"), ""))
      
   ElseIf Ind = 3 Then
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      CHEQUE_YYYYMM = NVLS(Rs("CHEQUE_YYYYMM"), "")
   ElseIf Ind = 4 Then
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
   ElseIf Ind = 5 Then
      CHEQUE_YYYYMM = NVLS(Rs("CHEQUE_YYYYMM"), "")
   ElseIf Ind = 6 Then
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      CHEQUE_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), ""))
   ElseIf Ind = 7 Then
      CHEQUE_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), ""))
   ElseIf Ind = 8 Then
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      BBANK_ID = NVLI(Rs("BBANK_ID"), -1)
   ElseIf Ind = 9 Then
      BBANK_ID = NVLI(Rs("BBANK_ID"), -1)
      BBANK_NAME = NVLS(Rs("BBANK_NAME"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PRINT_CHEQUE_LOG " & _
                     "( " & _
                     "   PAYEE_NAME, " & _
                     "   CHEQUE_AMOUNT, " & _
                     "   CHEQUE_DATE, " & _
                     "   CHEQUE_TYPE,CHEQUE_NO, " & _
                     "   PASS_FLAG, " & _
                     "   ENTERPRISE_ID, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(PAYEE_NAME)) & "'," & _
                     CHEQUE_AMOUNT & "," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(CHEQUE_DATE))) & "'," & _
                     "'" & ChangeQuote(Trim(CHEQUE_TYPE)) & "'," & "'" & ChangeQuote(Trim(CHEQUE_NO)) & "'," & _
                     "'N'," & _
                     "" & glbUser.ENTERPRISE_ID & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE PRINT_CHEQUE_LOG SET " & _
                        "   PAYEE_NAME = '" & ChangeQuote(Trim(PAYEE_NAME)) & "', " & _
                        "   CHEQUE_AMOUNT = " & CHEQUE_AMOUNT & ", " & _
                        "   CHEQUE_DATE= '" & ChangeQuote(Trim(DateToStringInt(CHEQUE_DATE))) & "', " & _
                        "   CHEQUE_TYPE = '" & ChangeQuote(Trim(CHEQUE_TYPE)) & "', " & "   CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_NO)) & "', " & _
                        "   PASS_FLAG = '" & ChangeQuote(Trim(PASS_FLAG)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                        " WHERE PRINT_CHEQUE_ID = " & PRINT_CHEQUE_ID
   End If
   
   Call m_Conn.Execute(SQL1)
   
End Function

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PRINT_CHEQUE_LOG WHERE PRINT_CHEQUE_ID = " & PRINT_CHEQUE_ID
      m_Conn.Execute (SQL1)
End Function
Public Function UpdatePassFlag()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE PRINT_CHEQUE_LOG SET " & _
                  "   PASS_FLAG = '" & ChangeQuote(Trim(PASS_FLAG)) & "', " & _
                  "   BANK_DATE= '" & ChangeQuote(Trim(DateToStringInt(BANK_DATE))) & "', " & _
                  "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                  "    MODIFY_BY = '" & glbUser.USER_NAME & "' " & _
                  " WHERE PRINT_CHEQUE_ID = " & PRINT_CHEQUE_ID
   
   Call m_Conn.Execute(SQL1)
End Function
Public Function UpdateCancelFlag()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE PRINT_CHEQUE_LOG SET " & _
                  "   CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "', " & _
                  "    CANCEL_DATE = '" & ServerDateTime & "', " & _
                  "   CANCEL_REASON = '" & ChangeQuote(Trim(CANCEL_REASON)) & "', " & _
                  "    CANCEL_BY = '" & glbUser.USER_NAME & "' " & _
                  " WHERE PRINT_CHEQUE_ID = " & PRINT_CHEQUE_ID
   
   Call m_Conn.Execute(SQL1)
End Function
Public Function UpdateChequeDate()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE PRINT_CHEQUE_LOG SET " & _
                  "   CHEQUE_DATE= '" & ChangeQuote(Trim(DateToStringInt(CHEQUE_DATE))) & "'," & _
                  "   CHEQUE_YYYYMM = '" & Left(ChangeQuote(Trim(DateToStringInt(CHEQUE_DATE))), 7) & "' " & _
                  " WHERE PRINT_CHEQUE_ID = " & PRINT_CHEQUE_ID
   Call m_Conn.Execute(SQL1)
End Function


