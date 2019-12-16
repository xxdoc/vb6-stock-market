VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExportChartAccount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   Icon            =   "frmExportChartAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5325
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   9393
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin StockMarket.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1740
         TabIndex        =   13
         Top             =   3420
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1740
         TabIndex        =   14
         Top             =   3750
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   11280
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin StockMarket.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   810
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtCollumn 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtRow 
         Height          =   435
         Left            =   5040
         TabIndex        =   3
         Top             =   1320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtSheet 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtCollumn2 
         Height          =   435
         Left            =   5040
         TabIndex        =   6
         Top             =   1800
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtCollumn3 
         Height          =   435
         Left            =   7800
         TabIndex        =   7
         Top             =   1800
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtCollumn4 
         Height          =   435
         Left            =   10800
         TabIndex        =   8
         Top             =   1800
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtRow2 
         Height          =   435
         Left            =   7800
         TabIndex        =   4
         Top             =   1320
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtName 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   2760
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   767
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblRow2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6000
         TabIndex        =   28
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblCollumn4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   8760
         TabIndex        =   27
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   720
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblCollumn3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   6000
         TabIndex        =   25
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblCollumn2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3240
         TabIndex        =   24
         Top             =   1920
         Width           =   1695
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   435
         Left            =   8600
         TabIndex        =   1
         Top             =   810
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccount.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblSheet 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   720
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3480
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblCollumn 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   20
         Top             =   930
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1740
         TabIndex        =   11
         Top             =   4380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccount.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3480
         TabIndex        =   19
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   18
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   17
         Top             =   3900
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9495
         TabIndex        =   12
         Top             =   4380
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExportChartAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
Dim MaxSheet As Long
   
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn, txtCollumn) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn2, txtCollumn2) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn3, txtCollumn3) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn4, txtCollumn4) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblRow, txtRow) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblRow2, txtRow2) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblSheet, txtSheet) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   MaxSheet = m_ExcelApp.Sheets.Count
   
   If Val(txtSheet.Text) > MaxSheet Then
      Call MsgBox("กรุณากรอกข้อมูล ชีดให้ถูกต้องโดยไม่สามารถมากกว่า  " & MaxSheet, vbOKOnly, PROJECT_NAME)
      Exit Sub
   End If
   
   Call ExportAccount
   
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
 
End Sub
Private Sub ExportAccount()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim iCount As Long
Dim I As Long
Dim TempNo As String
Dim j As Long
Dim MaxRow As Long
Dim Sm As CStockMarket
Dim Smd As CStockMarketDetail

   prgProgress.Max = 100
   prgProgress.Min = 0
   
   prgProgress.Value = 0
   
   I = 0
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet.Text))
   
   j = 0
   
   Set Sm = New CStockMarket
   Sm.STOCK_MARKET_DATE = uctlFromDate.ShowDate
   Sm.STOCK_MARKET_NAME = txtName.Text
   Sm.AddEditMode = SHOW_ADD
   
   iCount = Val(txtRow2.Text) - Val(txtRow.Text)
   While (j < iCount)
      j = j + 1
      prgProgress.Value = MyDiff(j, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      TempNo = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn.Text)).Value
      If (TempNo <> "") And (Val(m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Value) > 0) Then
         
         Set Smd = New CStockMarketDetail
         Smd.Flag = "A"
         Smd.ORDER_NO = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn.Text)).Value
         Smd.PAY_NAME = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Value
         Smd.AMOUNT = Val(m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Value)
         Smd.PRICE = Val(m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Value)
         
         Call Sm.StockMarketDetails.Add(Smd)
      End If
   Wend
   
   Call glbDaily.StartTransaction
   
   If Not glbDaily.AddEditStockMarket(Sm, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call glbDaily.CommitTransaction
   prgProgress.Value = 100
   txtPercent.Text = 100
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      'Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub
Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "Import Data From Excel"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblName, "รายการ")
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   
   Call InitNormalLabel(lblRow, "แถวเริ่ม")
   Call InitNormalLabel(lblRow2, "แถวจบ")
   Call InitNormalLabel(lblSheet, "ชีด")
   Call InitNormalLabel(lblCollumn, "คอลัมน์ลำดับ")
   Call InitNormalLabel(lblCollumn2, "คอลัมน์ชื่อผู้ถือหุ้น")
   Call InitNormalLabel(lblCollumn3, "คอลัมน์จำนวนหุ้น")
   Call InitNormalLabel(lblCollumn4, "คอลัมน์ผลตอบแทน")
   
   Call InitNormalLabel(lblFromDate, "วันที่เช็ค")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   'txtFileName.Enabled = False
   Call txtCollumn.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSheet.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCollumn2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCollumn3.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCollumn4.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   txtSheet.Text = "6"
   txtRow.Text = "5"
   txtRow2.Text = "161"
   txtCollumn.Text = "2"
   txtCollumn2.Text = "3"
   txtCollumn3.Text = "15"
   txtCollumn4.Text = "16"
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Call m_ExcelApp.Workbooks.Close
   'Call m_ExcelApp.Close
End Sub
