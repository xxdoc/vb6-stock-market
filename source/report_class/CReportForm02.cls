VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportForm02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CReportInterface

Private Const MODULE_NAME = "CReportForm02"
Private MODULE_DESC As String

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
Private Const HEADER_SIZE = 16
Private Const DETAIL_SIZE = 16

Private Const TITLE_FORMAT = "<52|<3|<5;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_ReportConfig As CReportConfig
Private Sub Class_Initialize()
   Set mColparam = New Collection
   Set m_ReportConfig = New CReportConfig
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mColparam = Nothing
   Set Vsp = Nothing
   Set m_ReportConfig = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mColparam.Add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property
Private Property Get CReportInterface_ReadParam() As Collection
   Set CReportInterface_ReadParam = mColparam
End Property
Private Function CReportInterface_Preview() As Boolean
   'If mcolParam("PREVIEW_TYPE") = 2 Then
    '  CReportInterface_Preview = genDoc(False)
   'Else
      CReportInterface_Preview = genDoc(True)
   'End If
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
Private Function initDoc() As Boolean
Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   
   Vsp.PaperSize = pprA4
   
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = "1cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginTop = "0cm"
   Vsp.MarginLeft = "5cm"
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("ใบส่งสินค้า")
   
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
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CChequeConfig
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim j As Long
Dim DateL As Double
Dim DateT As Double
Dim PayL As Double
Dim PayT As Double
Dim ThaiL As Double
Dim ThaiT As Double
Dim AmountL As Double
Dim AmountT As Double
Dim Logo1Flag As String
Dim Logo1L As Double
Dim Logo1T As Double
Dim Logo1Path As String
Dim Logo1Height As Double
Dim Logo1Weight As Double
Dim Logo2Flag As String
Dim Logo2L As Double
Dim Logo2T As Double
Dim Logo2Path As String
Dim Logo2Height As Double
Dim Logo2Weight As Double
Dim Name1Desc As String
Dim Name1DescL As Double
Dim Name1DescT As Double
Dim Name2Desc As String
Dim Name2DescL As Double
Dim Name2DescT As Double
Dim DateType As Long
Dim Bkt As CBkTrn
Dim Pos As Long
Dim ACL As Double
Dim ACT As Double
Dim BEARERL As Double
Dim BEARERT  As Double

   RName = "genDoc"

'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set Bkt = New CBkTrn
   Bkt.CHQNUM = mColparam("CHEQUE_NO")
   Call Bkt.QueryData(1, Rs, iCount)
   If Rs.EOF Then
      mstrErrMsg = "ไม่พบเลขที่เช็คที่ระบุ"
      genDoc = False
      Exit Function
   End If
   Call Bkt.PopulateFromRS(1, Rs)
   
   If Len(mColparam("PAY")) <= 0 Then
      Call mColparam.Remove("PAY")
      Call mColparam.Add(Bkt.SUPNAM, "PAY")
      'mColparam("PAY") = Bkt.SUPNAM
      If Len(mColparam("PAY")) <= 0 Then
         'ให้ตัดเอา ข้อความที่อยู่ข้างหน้า * มาเท่านั้น
         Pos = InStr(1, Bkt.REMARK, "*")
         Call mColparam.Remove("PAY")
         If Pos > 0 Then
            Call mColparam.Add(Mid(Bkt.REMARK, 1, Pos - 1), "PAY")
         Else
            Call mColparam.Add(Bkt.REMARK, "PAY")
         End If
      End If
   End If
   
   Call mColparam.Add(Bkt.AMOUNT, "AMOUNT")
   If mColparam("NO_DUE_DATE") = 0 Then
      Call mColparam.Add(DateToStringExtEx2(Bkt.CHQDAT), "DUE_DATE")
   Else
      Call mColparam.Add(-2, "DUE_DATE")
   End If
   
   Set Rs = New ADODB.Recordset
   Set Poi = New CChequeConfig
      
   Poi.CHEQUE_NAME = mColparam("CHEQUE_NAME")
   Call Poi.QueryData(Rs, iCount)
   
   
   Call Poi.PopulateFromRS(1, Rs)
   
   DateL = Poi.DATE_LEFT
   DateT = Poi.DATE_TOP
   DateType = Poi.DATE_TYPE
   PayL = Poi.PAY_LEFT
   PayT = Poi.PAY_TOP
   ThaiL = Poi.THAI_LEFT
   ThaiT = Poi.THAI_TOP
   AmountL = Poi.AMOUNT_LEFT
   AmountT = Poi.AMOUNT_TOP
   
   Logo1Flag = Poi.LOGO1_FLAG
   Logo1L = Poi.LOGO1_LEFT
   Logo1T = Poi.LOGO1_TOP
   Logo1Path = Poi.LOGO1_PATH
   Logo1Height = Poi.LOGO1_HEIGHT
   Logo1Weight = Poi.LOGO1_WEIGHT
   Logo2Flag = Poi.LOGO2_FLAG
   Logo2L = Poi.LOGO2_LEFT
   Logo2T = Poi.LOGO2_TOP
   Logo2Path = Poi.LOGO2_PATH
   Logo2Height = Poi.LOGO2_HEIGHT
   Logo2Weight = Poi.LOGO2_WEIGHT
   Name1Desc = Poi.NAME1_DESC
   Name1DescL = Poi.NAME1_DESC_LEFT
   Name1DescT = Poi.NAME1_DESC_TOP
   Name2Desc = Poi.NAME2_DESC
   Name2DescL = Poi.NAME2_DESC_LEFT
   Name2DescT = Poi.NAME2_DESC_TOP
   
   ACL = Poi.AC_LEFT
   ACT = Poi.AC_TOP
   BEARERL = Poi.BEARER_LEFT
   BEARERT = Poi.BEARER_TOP
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   Vsp.StartDoc

   '--------------------------------------------------------------------------------------------------------------------------------
   Vsp.TableBorder = tbNone

   Vsp.CurrentX = DateL * 567
   Vsp.CurrentY = DateT * 567

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 1) = DateL & "cm"
   Vsp.TableCell(tcColWidth, , 2) = ((Vsp.PageWidth / 567) - DateL) & "cm"
   
   If DateType = 1 Then
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(mColparam("DUE_DATE"))
   ElseIf DateType = 2 Then
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx22(mColparam("DUE_DATE"))
   ElseIf DateType = 3 Then
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx222(mColparam("DUE_DATE"))
   End If
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle

   Vsp.EndTable

   '--------------------------------------------------------------------------------------------------------------------------------
   Vsp.CurrentX = PayL * 567
   Vsp.CurrentY = PayT * 567

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 1) = PayL & "cm"
   Vsp.TableCell(tcColWidth, , 2) = ((Vsp.PageWidth / 567) - PayL) & "cm"

   Vsp.TableCell(tcText, 1, 2) = mColparam("PAY")
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle

   Vsp.EndTable

   '--------------------------------------------------------------------------------------------------------------------------------
   Vsp.CurrentX = ThaiL * 567
   Vsp.CurrentY = ThaiT * 567

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 1) = ThaiL & "cm"
   Vsp.TableCell(tcColWidth, , 2) = ((Vsp.PageWidth / 567) - ThaiL) & "cm"

   Vsp.TableCell(tcText, 1, 2) = "(" & ThaiBaht(mColparam("AMOUNT")) & ")"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle

   Vsp.EndTable

   '--------------------------------------------------------------------------------------------------------------------------------
   Vsp.CurrentX = AmountL * 567
   Vsp.CurrentY = AmountT * 567

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 1) = AmountL & "cm"
   Vsp.TableCell(tcColWidth, , 2) = ((Vsp.PageWidth / 567) - AmountL) & "cm"

   Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(mColparam("AMOUNT")) & "-"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Vsp.EndTable
   
   '--------------------------------------------------------------------------------------------------------------------------------
   
   If Logo1Flag = "Y" Then
      Vsp.CurrentX = Logo1L * 567
      Vsp.CurrentY = Logo1T * 567
            
      frmReport.Picture1.Picture = LoadPicture(Logo1Path)
      frmReport.Picture1.Visible = False
      Call Vsp.DrawPicture(frmReport.Picture1, Logo1L & "cm", Logo1T & "cm", Logo1Weight & "cm", Logo1Height & "cm", "Top")
   End If
   
   If Logo2Flag = "Y" Then
      Vsp.CurrentX = Logo2L * 567
      Vsp.CurrentY = Logo2T * 567
            
      frmReport.Picture1.Picture = LoadPicture(Logo2Path)
      frmReport.Picture1.Visible = False
      Call Vsp.DrawPicture(frmReport.Picture1, Logo2L & "cm", Logo2T & "cm", Logo2Weight & "cm", Logo2Height & "cm", "Top")
      
   End If
   
   If mColparam("SHOW_AC") = 1 Then
      Vsp.CurrentX = ACL * 567
      Vsp.CurrentY = ACT * 567
            
      frmReport.Picture1.Picture = LoadPicture(glbParameterObj.Acpayeeonly)
      frmReport.Picture1.Visible = False
      'Call Vsp.DrawPicture(frmReport.Picture1, Logo2L & "cm", Logo2T & "cm", Logo2Weight & "cm", Logo2Height & "cm", "Top")
      Call Vsp.DrawPicture(frmReport.Picture1, ACL & "cm", ACT & "cm", 3 & "cm", 3 & "cm", "Top")
   End If
   If mColparam("SHOW_BEARER") = 1 Then
      Vsp.CurrentX = BEARERL * 567
      Vsp.CurrentY = BEARERT * 567
      
      Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.CurrentX + 600, Vsp.CurrentY - 300)
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.CurrentX + 700, Vsp.CurrentY - 300)
   End If
   
   '--------------------------------------------------------------------------------------------------------------------------------
   Vsp.CurrentX = Name1DescL * 567
   Vsp.CurrentY = Name1DescT * 567

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 1) = Name1DescL & "cm"
   Vsp.TableCell(tcColWidth, , 2) = ((Vsp.PageWidth / 567) - Name1DescL) & "cm"

   Vsp.TableCell(tcText, 1, 2) = Name1Desc
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------
   Vsp.CurrentX = Name2DescL * 567
   Vsp.CurrentY = Name2DescT * 567
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 1) = Name2DescL & "cm"
   Vsp.TableCell(tcColWidth, , 2) = ((Vsp.PageWidth / 567) - Name2DescL) & "cm"

   Vsp.TableCell(tcText, 1, 2) = Name2Desc
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------
   
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   '--------------------------------------------------------------------------------------------------------------------------------
   Rs.Close
   Set Rs = Nothing
   
   genDoc = True
   
   'If mcolParam("PREVIEW_TYPE") = 2 Then
   'Vsp.PrintDoc (True)
   'End If
   
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   Set Rs = Nothing
End Function


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
   
'   Vsp.DrawPicture mcolParam("PICTURE"), "0.75in", "0.75in", "100%", "100%"
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
      
   'mdY = Vsp.CurrentY
   
   'Vsp.TextAlign = talnT
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
