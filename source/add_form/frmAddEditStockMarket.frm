VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditStockMarket 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditStockMarket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   1
         Top             =   1880
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5325
         Left            =   150
         TabIndex        =   2
         Top             =   2400
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   9393
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditStockMarket.frx":27A2
         Column(2)       =   "frmAddEditStockMarket.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditStockMarket.frx":290E
         FormatStyle(2)  =   "frmAddEditStockMarket.frx":2A6A
         FormatStyle(3)  =   "frmAddEditStockMarket.frx":2B1A
         FormatStyle(4)  =   "frmAddEditStockMarket.frx":2BCE
         FormatStyle(5)  =   "frmAddEditStockMarket.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditStockMarket.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1320
         Width           =   8325
         _extentx        =   5001
         _extenty        =   767
      End
      Begin StockMarket.uctlDate uctlDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   11
         Top             =   840
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin Threed.SSCommand cmdPrintReport 
         Height          =   525
         Left            =   6720
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditStockMarket.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   5040
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditStockMarket.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   12
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   390
         TabIndex        =   10
         Top             =   1440
         Width           =   1365
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditStockMarket.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   7
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   4
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   3
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditStockMarket.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   5
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditStockMarket.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditStockMarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_StockMarket As CStockMarket

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_StockMarket.STOCK_MARKET_ID = ID
      If Not glbDaily.QueryStockMarket(m_StockMarket, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_StockMarket.PopulateFromRS(1, m_Rs)
      
      txtDesc.Text = m_StockMarket.STOCK_MARKET_NAME
      uctlDate.ShowDate = m_StockMarket.STOCK_MARKET_DATE
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyDate(lblDate, uctlDate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
'   If Not CheckUniqueNs(StockMarket_NO, txtStockMarketNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtStockMarketNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
      
   m_StockMarket.AddEditMode = ShowMode
   m_StockMarket.STOCK_MARKET_ID = ID
   m_StockMarket.STOCK_MARKET_NAME = txtDesc.Text
   m_StockMarket.STOCK_MARKET_DATE = uctlDate.ShowDate
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditStockMarket(m_StockMarket, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditStockMarketDetail.ParentForm = Me
      Set frmAddEditStockMarketDetail.TempCollection = m_StockMarket.StockMarketDetails
      frmAddEditStockMarketDetail.ShowMode = SHOW_ADD
      frmAddEditStockMarketDetail.HeaderText = MapText("เพิ่มรายการ")
      Load frmAddEditStockMarketDetail
      frmAddEditStockMarketDetail.Show 1

      OKClick = frmAddEditStockMarketDetail.OKClick

      Unload frmAddEditStockMarketDetail
      Set frmAddEditStockMarketDetail = Nothing
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   Call RefreshGrid
   
   m_HasModify = True
   
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_StockMarket.StockMarketDetails.Remove (ID2)
      Else
         m_StockMarket.StockMarketDetails.Item(ID2).Flag = "D"
      End If
   
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   
   Call RefreshGrid
   m_HasModify = True

End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditStockMarketDetail.ParentForm = Me
      frmAddEditStockMarketDetail.ID = ID
      Set frmAddEditStockMarketDetail.TempCollection = m_StockMarket.StockMarketDetails
      frmAddEditStockMarketDetail.HeaderText = MapText("แก้ไขสินค้า")
      frmAddEditStockMarketDetail.ShowMode = SHOW_EDIT
      Load frmAddEditStockMarketDetail
      frmAddEditStockMarketDetail.Show 1
      
      OKClick = frmAddEditStockMarketDetail.OKClick

      Unload frmAddEditStockMarketDetail
      Set frmAddEditStockMarketDetail = Nothing
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Call RefreshGrid
   
   m_HasModify = True
   
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_StockMarket.STOCK_MARKET_ID
      m_StockMarket.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
End Sub

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
   
   ReportMode = 1
   
   If m_HasModify Or (m_StockMarket.STOCK_MARKET_ID <= 0) Then
      If Not SaveData Then
         Exit Sub
      End If
      ShowMode = SHOW_EDIT
      ID = m_StockMarket.STOCK_MARKET_ID
      m_StockMarket.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   End If
   
   ReportFlag = False
   
   ReportKey = "CReportForm01"
      
   Set Report = New CReportForm01
   ReportFlag = True
   
   If Not Report Is Nothing Then
      Call Report.AddParam(m_StockMarket.STOCK_MARKET_ID, "STOCK_MARKET_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1

      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub cmdPrintReport_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
   
   ReportMode = 1
   
   If m_HasModify Or (m_StockMarket.STOCK_MARKET_ID <= 0) Then
      If Not SaveData Then
         Exit Sub
      End If
      ShowMode = SHOW_EDIT
      ID = m_StockMarket.STOCK_MARKET_ID
      m_StockMarket.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   End If
   
   ReportFlag = False
   
   ReportKey = "CReportCheque001"
      
   Set Report = New CReportCheque001
   ReportFlag = True
   
   If Not Report Is Nothing Then
      Call Report.AddParam(m_StockMarket.STOCK_MARKET_ID, "STOCK_MARKET_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1

      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_StockMarket.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_StockMarket.QueryFlag = 0
         Call QueryData(False)
         uctlDate.ShowDate = Now
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_StockMarket = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 700
   Col.Caption = MapText("ลำดับ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3000
   Col.Caption = MapText("ชื่อผู้ถือหุ้น")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนหุ้น")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 3000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ผลตอบแทน")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 1000
   Col.Caption = MapText("เลขที่เช็ค")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 3000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ชื่อออกเช็ค")
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblDate, MapText("วันที่"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrintReport.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitMainButton(cmdPrint, MapText("พิมพ์เช็ค"))
   Call InitMainButton(cmdPrintReport, MapText("พิมพ์รายงาน"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายละเอียด")
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_StockMarket = New CStockMarket
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_StockMarket.StockMarketDetails Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CStockMarketDetail
      If m_StockMarket.StockMarketDetails.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_StockMarket.StockMarketDetails, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.STOCK_MARKET_DETAIL_ID
      Values(2) = RealIndex
      Values(3) = CR.ORDER_NO
      Values(4) = CR.PAY_NAME
      Values(5) = CR.AMOUNT
      Values(6) = CR.PRICE
      Values(7) = CR.CHEQUE_NO
      Values(8) = CR.CHEQUE_NAME
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
  If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_StockMarket.StockMarketDetails)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_StockMarket.StockMarketDetails)
   GridEX1.Rebind
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub


Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdPrintReport.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Private Sub uctlDate_HasChange()
   m_HasModify = True
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
