VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBkTrn"
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

Public TRNDAT  As Date
Public CHQNUM As String
Public CHQDAT As Date         'วันที่ DUE เช็ค
Public GETDAT As Date            'วันที่รับเช็ค
Public BNKCOD As String
Public CUSCOD As String
Public DEPCOD As String
Public PAYINDAT  As Date
Public REFNUM As String
Public AMOUNT As Double
Public REMARK  As String
Public VOUCHER As String

Public BNKNAM As String
Public BRANCH As String
Public SUPNAM As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public FROM_CHQDAT As Date
Public TO_CHQDAT As Date
Public FROM_GETDAT As Date
Public TO_GETDAT As Date

Public DOCDAT As Date
Public SUPCOD As String
Public TYPDES As String
Public SUPTYP As String
Public Credit As Long

Public ACCNUM As String
Public DESCRP As String

Public BNKNUM1  As String
Public BNKNUM2 As String
Public JNLTRNTYP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnectionExpress
   
   FROM_CHQDAT = -1
   TO_CHQDAT = -1
   FROM_GETDAT = -1
   TO_GETDAT = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT BKT.*, DTOS(BKT.CHQDAT) CHQ_CHQDAT, DTOS(BKT.GETDAT) CHQ_GETDAT, BKM.BRANCH BNKBRANCH, IST.TYPDES BNKNAM, APM.SUPNAM,BKM.BNKNAM BNKNUM1,BKM.BNKNUM BNKNUM2 "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN BKMAS BKM ON (BKT.BNKACC = BKM.BNKACC) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
   End If
   '
   If Trim(VOUCHER) <> "" Then
      If WildCard(VOUCHER, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (VOUCHER LIKE '" & ChangeQuote(Trim(VOUCHER)) & "')"
         Else
            WhereStr = WhereStr & "AND (VOUCHER LIKE '" & ChangeQuote(Trim(VOUCHER)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         Else
            WhereStr = WhereStr & "AND (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         End If
      End If
   End If
                                       
   If Trim(CHQNUM) <> "" Then
      If WildCard(CHQNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CHQNUM LIKE '" & ChangeQuote(Trim(CHQNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (CHQNUM LIKE '" & ChangeQuote(Trim(CHQNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CHQNUM = '" & ChangeQuote(Trim(CHQNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (CHQNUM = '" & ChangeQuote(Trim(CHQNUM)) & "')"
         End If
      End If
   End If
                                       
   If FROM_GETDAT > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_GETDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(GETDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(GETDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_GETDAT > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_GETDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(GETDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(GETDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
                                       
   If FROM_CHQDAT > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_CHQDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_CHQDAT > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CHQDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(SUPCOD) <> "" Then
      If WildCard(SUPCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      End If
   End If
                  
   If Len(SUPTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      End If
   End If
   
   If Len(ACCNUM) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GL.ACCNUM = '" & Trim(ChangeQuote(ACCNUM)) & "') "
      Else
         WhereStr = WhereStr & "AND (GL.ACCNUM = '" & Trim(ChangeQuote(ACCNUM)) & "') "
      End If
   End If
   
   If Trim(REMARK) <> "" Then
      If WildCard(REMARK, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (BKT.REMARK LIKE '" & ChangeQuote(Trim(REMARK)) & "')"
         Else
            WhereStr = WhereStr & "AND (BKT.REMARK LIKE '" & ChangeQuote(Trim(REMARK)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BKT.REMARK = '" & ChangeQuote(Trim(REMARK)) & "')"
         Else
            WhereStr = WhereStr & "AND (BKT.REMARK = '" & ChangeQuote(Trim(REMARK)) & "')"
         End If
      End If
   End If
   
   If Len(JNLTRNTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BKT.JNLTRNTYP = '" & Trim(ChangeQuote(JNLTRNTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (BKT.JNLTRNTYP = '" & Trim(ChangeQuote(JNLTRNTYP)) & "') "
      End If
   End If
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CHQNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CHQNUM "
      Else
         OrderByStr = "ORDER BY CHQNUM "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRsExpress(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TRNDAT = InternalDateToDateExpress(NVLS(Rs("TRNDAT"), ""))
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      CHQDAT = InternalDateToDateExpress(NVLS(Rs("CHQ_CHQDAT"), ""))
      GETDAT = InternalDateToDateExpress(NVLS(Rs("CHQ_GETDAT"), ""))
      BNKCOD = NVLS(Rs("BNKCOD"), "")
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      DEPCOD = NVLS(Rs("DEPCOD"), "")
      PAYINDAT = InternalDateToDateExpress(NVLS(Rs("PAYINDAT"), ""))
      REFNUM = NVLS(Rs("REFNUM"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      REMARK = NVLS(Rs("REMARK"), "")
      VOUCHER = NVLS(Rs("VOUCHER"), "")
      
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      BNKNAM = NVLS(Rs("BNKNAM"), "")
      BRANCH = NVLS(Rs("BRANCH"), "")
      
      BNKNUM1 = NVLS(Rs("BNKNUM1"), "")
      BNKNUM2 = NVLS(Rs("BNKNUM2"), "")
   End If
End Function
