VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCheque004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCheque004"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mColparam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList

Private AmountColl As Collection
Private BBankColl As Collection
Private Sub Class_Initialize()
   Set mColparam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set AmountColl = New Collection
   Set BBankColl = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mColparam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set AmountColl = Nothing
   Set BBankColl = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mColparam.Add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function
Private Property Get CReportInterface_ReadParam() As Collection
   Set CReportInterface_ReadParam = mColparam
End Property

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mColparam.Count
      mColparam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Pcl As CPrintChequeLog

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "NO", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "�������", "<")
   Call m_Headers1.AddField(Cf)
   
   For Each Pcl In BBankColl
      Call Cf.SetFieldValue(2, "^", Pcl.BBANK_NAME, ">")
      Call m_Headers1.AddField(Cf)
   Next Pcl
   
   Call Cf.SetFieldValue(2, "^", "���", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = mColparam("REPORT_NAME")
                                          
   FromDate = "�ҡ�ѹ����� " & EmptyToString(DateToStringExtEx2(mColparam("FROM_CHEQUE_DATE")), "N/A")
   ToDate = "�֧�ѹ����� " & EmptyToString(DateToStringExtEx2(mColparam("TO_CHEQUE_DATE")), "N/A")
   
   Call LoadPrintChequeLogBBank(AmountColl, mColparam("PAY"), mColparam("FROM_CHEQUE_NO"), mColparam("TO_CHEQUE_NO"), mColparam("FROM_CHEQUE_DATE"), mColparam("TO_CHEQUE_DATE"), mColparam("BBANK_ID"), mColparam("PASS_FLAG"))
   Call LoadBBankFromPCL(BBankColl, mColparam("PAY"), mColparam("FROM_CHEQUE_NO"), mColparam("TO_CHEQUE_NO"), mColparam("FROM_CHEQUE_DATE"), mColparam("TO_CHEQUE_DATE"), mColparam("BBANK_ID"), mColparam("PASS_FLAG"))
   
   MODULE_DESC = glbUser.ENTERPRISE_NAME & "    ��§ҹ " & mColparam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " �ҡ�� " & EmptyToString(mColparam("FROM_CHEQUE_NO"), "N/A") & " �֧�� " & EmptyToString(mColparam("TO_CHEQUE_NO"), "N/A") & vbCrLf & _
                                       " �ҢҸ�Ҥ�� " & EmptyToString(mColparam("BBANK_NAME"), "N/A") & " " & " " & " ��ҹ�� " & EmptyToString(mColparam("PASS_FLAG_NAME"), "N/A") & " " & " " & " ���§���������� " & " ���§�ҡ " & EmptyToString(mColparam("ORDER_TYPE_NAME"), "N/A")
                                       
                                       
   Call SetReportConfig(Vsp, mColparam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim m_PrintChequeLog As CPrintChequeLog
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim iCount As Long
Dim TempStr As String
Dim Total1(100) As Double
Dim Temp As Long
Dim Pcl As CPrintChequeLog
Dim TempPcl As CPrintChequeLog
Dim SumRow As Double
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
      
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField
      
      Set m_PrintChequeLog = New CPrintChequeLog
      
      m_PrintChequeLog.PRINT_CHEQUE_ID = -1
      m_PrintChequeLog.PAYEE_NAME = PatchWildCard(mColparam("PAY"))
      m_PrintChequeLog.FROM_CHEQUE_NO = mColparam("FROM_CHEQUE_NO")
      m_PrintChequeLog.TO_CHEQUE_NO = mColparam("TO_CHEQUE_NO")
      m_PrintChequeLog.FROM_CHEQUE_DATE = mColparam("FROM_CHEQUE_DATE")
      m_PrintChequeLog.TO_CHEQUE_DATE = mColparam("TO_CHEQUE_DATE")
      m_PrintChequeLog.BBANK_ID = mColparam("BBANK_ID")
      
      m_PrintChequeLog.PASS_FLAG = PassChequeToString(mColparam("PASS_FLAG"))
      m_PrintChequeLog.CANCEL_FLAG = "N"
      m_PrintChequeLog.OrderType = mColparam("ORDER_TYPE")

      Call m_PrintChequeLog.QueryData(4, Rs, iCount)
      
      I = 0
      While Not Rs.EOF
         I = I + 1
         j = 0
         SumRow = 0
         Call m_PrintChequeLog.PopulateFromRS(4, Rs)
         
         If I = 1 Then
            Call GenerateNewPage
         End If
         
         Call m_Details1.ClearField
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                  
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_PrintChequeLog.PAYEE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         For Each Pcl In BBankColl
            Set TempPcl = GetObject("CPrintChequeLog", AmountColl, Trim(m_PrintChequeLog.PAYEE_NAME & "-" & Pcl.BBANK_ID))
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(TempPcl.CHEQUE_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + TempPcl.CHEQUE_AMOUNT
            SumRow = SumRow + TempPcl.CHEQUE_AMOUNT
         Next Pcl
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(SumRow)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + SumRow
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Rs.MoveNext
      Wend
       
       Call GenerateFooter("�ط��", "", "", Total1, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      Else
         TempStr = FormatNumberToNull(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
   If Vsp.CurrentPage > 1 Then
      Call GenerateNewPage
   End If
End Sub
Private Sub GenerateNewPage()
Dim talnT As TextAlignSettings
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taLeftBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp = ""
   
   Vsp.FontSize = iSize
   
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
