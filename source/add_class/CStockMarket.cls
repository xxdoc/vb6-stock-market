VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockMarket"
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

Public STOCK_MARKET_ID As Long
Public STOCK_MARKET_DATE As Date
Public STOCK_MARKET_NAME As String

Public StockMarketDetails As Collection
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
      SelectStr1 = "SELECT UG.* "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = " FROM STOCK_MARKET UG "
      
   End If
   WhereStr = ""
   
   If STOCK_MARKET_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.STOCK_MARKET_ID = " & STOCK_MARKET_ID & ")"
      Else
         WhereStr = WhereStr & "AND (UG.STOCK_MARKET_ID = " & STOCK_MARKET_ID & ")"
      End If
   End If
   
   If Ind = 3 Then
      'GroupByStr = " GROUP BY UG.PAYEE_NAME,UG.CHEQUE_YYYYMM  "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY UG.STOCK_MARKET_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY UG.STOCK_MARKET_ID "
      Else
         OrderByStr = "ORDER BY UG.STOCK_MARKET_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
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
      STOCK_MARKET_ID = NVLI(Rs("STOCK_MARKET_ID"), -1)
      STOCK_MARKET_DATE = InternalDateToDate(NVLS(Rs("STOCK_MARKET_DATE"), ""))
      STOCK_MARKET_NAME = NVLS(Rs("STOCK_MARKET_NAME"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO STOCK_MARKET " & _
                     "( " & _
                     "   STOCK_MARKET_DATE, " & _
                     "   STOCK_MARKET_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(DateToStringInt(STOCK_MARKET_DATE))) & "'," & _
                     "'" & ChangeQuote(Trim(STOCK_MARKET_NAME)) & "'" & _
                     ")  RETURNING STOCK_MARKET_ID "
   Else
            SQL1 = "UPDATE STOCK_MARKET SET " & _
                        "   STOCK_MARKET_DATE= '" & ChangeQuote(Trim(DateToStringInt(STOCK_MARKET_DATE))) & "', " & _
                        "   STOCK_MARKET_NAME = '" & ChangeQuote(Trim(STOCK_MARKET_NAME)) & "' " & _
                        " WHERE STOCK_MARKET_ID = " & STOCK_MARKET_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      STOCK_MARKET_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Set StockMarketDetails = New Collection
End Sub
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM STOCK_MARKET_DETAIL WHERE STOCK_MARKET_ID = " & STOCK_MARKET_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM STOCK_MARKET WHERE STOCK_MARKET_ID = " & STOCK_MARKET_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set StockMarketDetails = Nothing
End Sub
