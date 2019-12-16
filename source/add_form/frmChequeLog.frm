VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmChequeLog 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmChequeLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   975
         Left            =   180
         TabIndex        =   0
         Top             =   840
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   1720
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MultiSelect     =   -1  'True
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
         Column(1)       =   "frmChequeLog.frx":27A2
         Column(2)       =   "frmChequeLog.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmChequeLog.frx":290E
         FormatStyle(2)  =   "frmChequeLog.frx":2A6A
         FormatStyle(3)  =   "frmChequeLog.frx":2B1A
         FormatStyle(4)  =   "frmChequeLog.frx":2BCE
         FormatStyle(5)  =   "frmChequeLog.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmChequeLog.frx":2D5E
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   4410
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmChequeLog.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   2760
         TabIndex        =   5
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   165
         TabIndex        =   1
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   2
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmChequeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_StockMarket As CStockMarket
Private m_TempStockMarket As CStockMarket
Private m_Rs As ADODB.Recordset
Public OKClick As Boolean
Private Sub cmdAdd_Click()
On Error GoTo ErrorHandler

   ' Load frmImportExcel
   Load frmExportChartAccount
   frmExportChartAccount.Show 1
   
   Unload frmExportChartAccount
   Set frmExportChartAccount = Nothing
   
   Call QueryData(True)

ErrorHandler:
   Call glbDaily.RollbackTransaction
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim ID As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   m_StockMarket.STOCK_MARKET_ID = ID
   If Not glbDaily.DeleteStockMarket(m_StockMarket, IsOK, True, glbErrorLog) Then
      m_StockMarket.STOCK_MARKET_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call EnableForm(Me, True)
End Sub


Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim TempStr As String
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   frmAddEditStockMarket.ID = ID
   frmAddEditStockMarket.HeaderText = MapText("แก้ไขข้อมูลการจ่ายเงิน")
   frmAddEditStockMarket.ShowMode = SHOW_EDIT
   Load frmAddEditStockMarket
   frmAddEditStockMarket.Show 1

   OKClick = frmAddEditStockMarket.OKClick

   Unload frmAddEditStockMarket
   Set frmAddEditStockMarket = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_StockMarket.STOCK_MARKET_ID = -1
      
      If Not glbDaily.QueryStockMarket(m_StockMarket, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call InitGrid
   
   'GridEX1.ItemCount = m_Rs.RecordCount
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempStockMarket.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempStockMarket.STOCK_MARKET_ID
   Values(2) = DateToStringExtEx2(m_TempStockMarket.STOCK_MARKET_DATE)
   Values(3) = m_TempStockMarket.STOCK_MARKET_NAME
   
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR

   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2300
   Col.Caption = MapText("วันที่เช็ค")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5450
   Col.Caption = MapText("รายละเอียด")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      'Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      'Call ShowSearch
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      'Call cmdCommit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      'Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      'Call cmdCancel_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   End If
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลการจ่ายเช็คสำหรับผู้ถือหุ้น")
   pnlHeader.Caption = MapText("ข้อมูลการจ่ายเช็คสำหรับผู้ถือหุ้น")
   
   Call InitGrid
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่มข้อมูล (F7)"))
   
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   
   Set m_StockMarket = New CStockMarket
   Set m_TempStockMarket = New CStockMarket
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = cmdAdd.Top
   cmdDelete.Top = cmdAdd.Top
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      'Call ShowSearch
      KeyCode = 0
   End If
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(2)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
