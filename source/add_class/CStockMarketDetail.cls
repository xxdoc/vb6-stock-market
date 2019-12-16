VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockMarketDetail"
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

Public STOCK_MARKET_DETAIL_ID As Long
Public STOCK_MARKET_ID As Long
Public ORDER_NO As Long
Public PAY_NAME As String
Public AMOUNT As Double
Public PRICE As Double
Public CHEQUE_NO As String
Public CHEQUE_NAME As String

Public STOCK_MARKET_DATE As Date
'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
   
   If Ind = 1 Then  'frmStockMarket
      SelectStr1 = "SELECT SMD.* "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM STOCK_MARKET_DETAIL SMD "
   ElseIf Ind = 2 Then  'ReportForm01
      SelectStr1 = "SELECT SMD.*,SM.STOCK_MARKET_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM STOCK_MARKET_DETAIL SMD "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_MARKET SM ON (SMD.STOCK_MARKET_ID = SM.STOCK_MARKET_ID) "
   End If
   WhereStr = ""
   
   If STOCK_MARKET_DETAIL_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SMD.STOCK_MARKET_DETAIL_ID = " & STOCK_MARKET_DETAIL_ID & ")"
      Else
         WhereStr = WhereStr & "AND (SMD.STOCK_MARKET_DETAIL_ID = " & STOCK_MARKET_DETAIL_ID & ")"
      End If
   End If
   
   If STOCK_MARKET_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SMD.STOCK_MARKET_ID = " & STOCK_MARKET_ID & ")"
      Else
         WhereStr = WhereStr & "AND (SMD.STOCK_MARKET_ID = " & STOCK_MARKET_ID & ")"
      End If
   End If
   
   If Ind = 3 Then
      'GroupByStr = " GROUP BY UG.PAYEE_NAME,UG.CHEQUE_YYYYMM  "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SMD.STOCK_MARKET_DETAIL_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SMD.STOCK_MARKET_DETAIL_ID "
      Else
         OrderByStr = "ORDER BY SMD.ORDER_NO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY SMD.ORDER_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
'   If OrderBy = 2 Then
'      OrderTypeStr = OrderTypeStr & "ROWS 1 TO 20 "     'สำหรับเอา 20 Record แรก
'   End If
   If OrderByStr = "" Then
      OrderTypeStr = ""
   End If
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      STOCK_MARKET_DETAIL_ID = NVLI(Rs("STOCK_MARKET_DETAIL_ID"), -1)
      STOCK_MARKET_ID = NVLI(Rs("STOCK_MARKET_ID"), -1)
      
      ORDER_NO = NVLI(Rs("ORDER_NO"), -1)
      PAY_NAME = NVLS(Rs("PAY_NAME"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      PRICE = NVLD(Rs("PRICE"), 0)
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      CHEQUE_NAME = NVLS(Rs("CHEQUE_NAME"), "")
   ElseIf Ind = 2 Then
      STOCK_MARKET_DETAIL_ID = NVLI(Rs("STOCK_MARKET_DETAIL_ID"), -1)
      STOCK_MARKET_ID = NVLI(Rs("STOCK_MARKET_ID"), -1)
      
      ORDER_NO = NVLI(Rs("ORDER_NO"), -1)
      PAY_NAME = NVLS(Rs("PAY_NAME"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      PRICE = NVLD(Rs("PRICE"), 0)
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      CHEQUE_NAME = NVLS(Rs("CHEQUE_NAME"), "")
      
      STOCK_MARKET_DATE = InternalDateToDate(NVLS(Rs("STOCK_MARKET_DATE"), ""))
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO STOCK_MARKET_DETAIL " & _
                     "( " & _
                     "   STOCK_MARKET_ID, " & _
                     "   ORDER_NO, " & _
                     "   PAY_NAME, " & _
                     "   AMOUNT, " & _
                     "   PRICE, " & _
                     "   CHEQUE_NO, " & _
                     "   CHEQUE_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     STOCK_MARKET_ID & "," & _
                     ORDER_NO & "," & _
                     "'" & ChangeQuote(Trim(PAY_NAME)) & "'," & _
                     AMOUNT & "," & _
                     PRICE & "," & _
                     "'" & ChangeQuote(Trim(CHEQUE_NO)) & "'," & _
                     "'" & ChangeQuote(Trim(CHEQUE_NAME)) & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE STOCK_MARKET_DETAIL SET " & _
                       "   ORDER_NO = " & ORDER_NO & ", " & _
                        "   PAY_NAME = '" & ChangeQuote(Trim(PAY_NAME)) & "', " & _
                        "   AMOUNT = " & AMOUNT & ", " & _
                        "   PRICE = " & PRICE & ", " & _
                        "   CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_NO)) & "', " & _
                        "   CHEQUE_NAME = '" & ChangeQuote(Trim(CHEQUE_NAME)) & "' " & _
                        " WHERE STOCK_MARKET_DETAIL_ID = " & STOCK_MARKET_DETAIL_ID
   End If
   
   Call m_Conn.Execute(SQL1)
   
End Function

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM STOCK_MARKET_DETAIL WHERE STOCK_MARKET_DETAIL_ID = " & STOCK_MARKET_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function
