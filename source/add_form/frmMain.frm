VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   9000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15875
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame3 
         Height          =   6015
         Left            =   4275
         TabIndex        =   9
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   10610
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboGeneric 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   540
            Visible         =   0   'False
            Width           =   2850
         End
         Begin StockMarket.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   1590
            TabIndex        =   11
            Top             =   930
            Visible         =   0   'False
            Width           =   2850
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin StockMarket.uctlDate uctlGenericDate 
            Height          =   405
            Index           =   0
            Left            =   1590
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   2850
            _ExtentX        =   5689
            _ExtentY        =   291
         End
         Begin Threed.SSCheck chkCommit 
            Height          =   300
            Index           =   0
            Left            =   1440
            TabIndex        =   17
            Top             =   2400
            Visible         =   0   'False
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   529
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   435
            Left            =   1440
            TabIndex        =   15
            Top             =   1860
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   767
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMain.frx":1CCA
            ButtonStyle     =   3
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin Threed.SSCheck chkGeneric 
            Height          =   465
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Visible         =   0   'False
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   820
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   7920
         Top             =   7560
      End
      Begin VB.PictureBox Picture1 
         Height          =   765
         Left            =   4560
         ScaleHeight     =   705
         ScaleWidth      =   825
         TabIndex        =   2
         Top             =   7440
         Visible         =   0   'False
         Width           =   885
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   495
         Left            =   4275
         TabIndex        =   5
         Top             =   0
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   873
         _Version        =   131073
         BackStyle       =   1
         Begin Threed.SSCommand SSCommand1 
            Height          =   555
            Left            =   9660
            TabIndex        =   7
            Top             =   6390
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   979
            _Version        =   131073
            PictureFrames   =   1
            Picture         =   "frmMain.frx":1FE4
            Caption         =   "SSCommand1"
            ButtonStyle     =   3
         End
         Begin VB.Label lblDateTime 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   525
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   3000
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   495
         Left            =   4275
         TabIndex        =   8
         Top             =   480
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   873
         _Version        =   131073
         BackStyle       =   1
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   6060
         Left            =   0
         TabIndex        =   0
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10689
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cordia New"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5640
         Top             =   7440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3074
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":394E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4228
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4B02
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4C5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5536
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5E10
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":612A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6A04
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":72DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7BB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8892
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblUserName 
         Caption         =   "Label1"
         Height          =   525
         Left            =   3120
         TabIndex        =   18
         Top             =   8280
         Width           =   4440
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   1560
         TabIndex        =   16
         Top             =   8160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9780
         TabIndex        =   4
         Top             =   8130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdConfig 
         Height          =   525
         Left            =   8070
         TabIndex        =   3
         Top             =   8130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Const ROOT_TREE = "CHEQUE"
Private Const ROOT_TREE1 = "USER"
Private Const ROOT_TREE2 = "REPORT"
Private Const ROOT_TREE3 = "BBANK"

Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset
Private m_TableName As String

Public HeaderText As String
Private m_MustAsk As Boolean

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_TextLookups As Collection
Private m_Dates As Collection
Private m_CheckBoxes As Collection
Private m_Labels As Collection
Private m_Checks As Collection

Private m_Combos As Collection

Private m_ReportParams As Collection
Private m_FromDate As Date
Private m_ToDate As Date
Private m_DBPath As String
Private C As CReportControl

Private m_MonthID As Long
Private m_YearNo As String

Private m_StockMarket As CStockMarket
Private m_Rs1  As ADODB.Recordset
Private Result As Long
Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long

   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If

   ReportKey = trvMain.SelectedItem.Key
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   Rc.COMPUTER_NAME = glbParameterObj.ComputerName
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If

   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMain.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String
Dim C As CReportControl
   
   Key = trvMain.SelectedItem.Key
   Name = trvMain.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Not (trvMain.SelectedItem Is Nothing) Then
      Call Report.AddParam(trvMain.SelectedItem.Text, "REPORT_TEXT")
   End If
   
   If Key = ROOT_TREE & " 1-0-1" Then
      'Set Report = New CReportForm01
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 1-0-2" Then
      'Set Report = New CReportForm02
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-1" Then
      'Set Report = New CReportCheque001
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-1-1" Then
      'Set Report = New CReportCheque001_1
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-1-2" Then
      'Set Report = New CReportCheque001_2
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-1-3" Then
      'Set Report = New CReportCheque001_3
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-2" Then
      'Set Report = New CReportCheque002
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-3" Then
      'Set Report = New CReportCheque003
      SelectFlag = True
   ElseIf Key = ROOT_TREE2 & " 1-4" Then
      'Set Report = New CReportCheque004
      SelectFlag = True
   
   End If

   If SelectFlag Then
      Call FillReportInput(Report)
      Call Report.AddParam(Name, "REPORT_NAME")
      Call Report.AddParam(Key, "REPORT_KEY")
      
      Set frmReport.ReportObject = Report
      frmReport.ReportID = Key
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
   End If
End Sub
Private Sub Form_Activate()
Dim OKClick As Boolean
Dim DBPath As String

   If m_HasActivate Then
      Exit Sub
   End If
   m_HasActivate = True
   
   Load frmLogin
   frmLogin.Show 1
   
   OKClick = frmLogin.OKClick

   Unload frmLogin
   Set frmLogin = Nothing
   
   If OKClick Then
      Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
      Call InitNormalLabel(lblUsername, "USER NAME : " & glbUser.USER_NAME & " (" & glbUser.GROUP_NAME & ")", RGB(0, 0, 255))
      SSFrame2.Visible = True
      Call InitMainTreeview
      
      Me.Caption = Me.Caption
   Else
      m_MustAsk = False
      Call cmdExit_Click
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
   Call InitFormLayout
   Set m_Rs = New ADODB.Recordset
   
   Set m_ReportControls = New Collection
   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
  Set m_TextLookups = New Collection
   Set m_Combos = New Collection
   Set m_ReportParams = New Collection
   Set m_CheckBoxes = New Collection
   Set m_Checks = New Collection
   
   Set m_StockMarket = New CStockMarket
   Set m_Rs1 = New ADODB.Recordset

End Sub

Private Sub InitFormLayout()
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.FontSize = 20
   lblDateTime.BackStyle = 1
   
   lblDateTime.BackColor = RGB(255, 255, 255)
   SSFrame2.Visible = False
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Me.Caption = MapText("MITTRAPHAP STOCK-MARKET " & App.Major & "." & App.Revision)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
      
   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม"))
   
   
   SSFrame3.Visible = False
   cmdConfig.Visible = False
   
End Sub

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String

   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT_EX
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
   
  If VerifyAccessRight("STOCK-MARKET", , False) Then
      Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, MapText("โปรแกรมพิมพ์เช็คสำหรับผู้ถือหุ้น"), 8, 8)
      Node.Expanded = True
      Node.Selected = True
      
      If VerifyAccessRight("STOCK-MARKET-DATA", , False) Then
         Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0-1", MapText("1.ข้อมูลการจ่ายเช็ค"), 3, 3)
         Node.Expanded = False
      End If
      If VerifyAccessRight("CHEQUE_CONFIG", , False) Then
         Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2-1", MapText("2.ตั้งค่าเช็ค"), 3, 3)
         Node.Expanded = False
      End If
   End If
   
   If VerifyAccessRight("REPORT", , False) Then
      Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE2, MapText("รายงาน"), 11, 11)
      Node.Expanded = True
      Node.Selected = True
      
      If VerifyAccessRight("REPORT_1", MapText("ประวัติเช็ค"), False) Then
         Set Node = trvMain.Nodes.Add(ROOT_TREE2, tvwChild, ROOT_TREE2 & " 1-1", MapText("ประวัติการจ่ายเงินผู้ถือหุ้น"), 3, 3)
         Node.Expanded = False
      End If
   End If
   
   If VerifyAccessRight("ADMIN", , False) Then
      Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE1, MapText("ระบบผู้ใช้"), 4, 4)
      Node.Expanded = True
      Node.Selected = True
      If VerifyAccessRight("ADMIN_GROUP", , False) Then
         Set Node = trvMain.Nodes.Add(ROOT_TREE1, tvwChild, ROOT_TREE1 & " -1", MapText("กลุ่มผู้ใช้"), 3, 3)
         Node.Expanded = False
      End If
      If VerifyAccessRight("ADMIN_USER", , False) Then
         Set Node = trvMain.Nodes.Add(ROOT_TREE1, tvwChild, ROOT_TREE1 & " -2", MapText("ผู้ใช้"), 3, 3)
         Node.Expanded = False
      End If
   End If
   
   If VerifyAccessRight("BBANK", MapText("สาขาธนาคาร"), False) Then
      Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE3, MapText("สาขาธนาคาร"), 1, 1)
      Node.Expanded = False
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านต้องการออกจากโปรแกรมใช่หรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
      Else
         Cancel = True
      End If
   Else
      Cancel = False
   End If
End Sub

Private Sub FillReportInput(R As CReportInterface)
On Error Resume Next
   
   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
         
         If C.Param2 = "MONTH_ID" Then
            m_MonthID = cboGeneric(C.ControlIndex).ListIndex
         End If
         
         
         
      End If
      
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
         
         If Len(txtGeneric(C.ControlIndex).Text) = 0 Then
            If C.Param2 = "YEAR_NO" Then
               txtGeneric(C.ControlIndex).Text = Year(Now)
            End If
         End If
         
         If C.Param2 = "YEAR_NO" Then
            m_YearNo = txtGeneric(C.ControlIndex).Text
         End If
         
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_STOCK_MARKET_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_STOCK_MARKET_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               End If
            End If
            If C.Param2 = "FROM_STOCK_MARKET_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_STOCK_MARKET_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
      
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param2)
         End If
      End If
      
   Next C
End Sub
Private Function VerifyReportInput() As Boolean
Dim C As CReportControl

   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
      
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
         If C.Param2 = "AMOUNT" Then
            m_Texts(C.ControlIndex).Text = Replace(m_Texts(C.ControlIndex).Text, ",", "")
            If Not IsNumeric(m_Texts(C.ControlIndex).Text) Then
                  glbErrorLog.LocalErrorMsg = "กรุณาใส่ยอดเช็คที่เป็นตัวเลขเท่านั้น"
                  glbErrorLog.ShowUserError
                  Exit Function
            End If
         End If
      End If
   
      If (C.ControlType = "D") Then
         If C.Param2 = "DUE_DATE" Then
            If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
               Exit Function
            Else
               If Not m_Dates(C.ControlIndex).VerifyDate(False) Then
                  glbErrorLog.LocalErrorMsg = "ไม่มีการใส่วันที่เช็ค กรุณาใส่วันที่เช็คในภายหลัง"
                  glbErrorLog.ShowUserError
               Else 'มีการใส่วันที่ ต้องเช็คว่าเป็นวันที่ย้อนหลังหรือไม่
                  If Not VerifyDateToDay(Nothing, m_Dates(C.ControlIndex)) Then
                     
                  End If
               End If
            End If
            
            
         Else
            If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
               Exit Function
            End If
         End If
      End If
   Next C
   VerifyReportInput = True
End Function
Private Function VerifyReportInputDrCr() As Boolean
Dim C As CReportControl

   VerifyReportInputDrCr = False

   For Each C In m_ReportControls
         If (C.ControlType = "D") Then
            If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
               Exit Function
            End If
         End If
   Next C
   VerifyReportInputDrCr = True
End Function
Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "", Optional KeySearch As String, Optional OldLine As Boolean = False, Optional ToolTipText As String)
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim ChIdx As Long

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   ChIdx = m_Checks.Count + 1
   
   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.Add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
      lblGeneric(LblIdx).ToolTipText = ToolTipText
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.Add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
      C.OldLine = OldLine
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.Add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
      C.OldLine = OldLine
      txtGeneric(TxtIdx).SetKeySearch (KeySearch)
      
      If Param1 = "YEAR_NO" Then
         If Len(m_YearNo) > 0 Then
            txtGeneric(TxtIdx).Text = m_YearNo
         Else
            txtGeneric(TxtIdx).Text = Year(Now) + 543
         End If
      End If
      
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.Add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx
      
      If Param1 = "FROM_STOCK_MARKET_DATE" Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      ElseIf Param1 = "TO_STOCK_MARKET_DATE" Then
         If m_ToDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         End If
      End If
      
   ElseIf ControlType = "LU" Then
'         Load uctlTextLookup(LkupIdx)
'         Call m_TextLookups.Add(uctlTextLookup(LkupIdx))
'         C.ControlIndex = LkupIdx
   ElseIf ControlType = "CH" Then
      Load chkCommit(ChIdx)
      Call m_Checks.Add(chkCommit(ChIdx))
      Call InitCheckBox(chkCommit(ChIdx), TextMsg)
      C.ControlIndex = ChIdx
   End If
   
   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.Add(C)
   Set C = Nothing
End Sub
Private Sub UnloadAllControl()
Dim I As Long
Dim j As Long

   I = m_Labels.Count
   While I > 0
      Call Unload(m_Labels(I))
      Call m_Labels.Remove(I)
      I = I - 1
   Wend
   
   I = m_Texts.Count
   While I > 0
      Call Unload(m_Texts(I))
      Call m_Texts.Remove(I)
      I = I - 1
   Wend

   I = m_Dates.Count
   While I > 0
      Call Unload(m_Dates(I))
      Call m_Dates.Remove(I)
      I = I - 1
   Wend

   I = m_Combos.Count
   While I > 0
      Call Unload(m_Combos(I))
      Call m_Combos.Remove(I)
      I = I - 1
   Wend
   
   I = m_TextLookups.Count
   While I > 0
      Call Unload(m_TextLookups(I))
      Call m_TextLookups.Remove(I)
      I = I - 1
   Wend
   
   I = m_Checks.Count
   While I > 0
      Call Unload(m_Checks(I))
      Call m_Checks.Remove(I)
      I = I - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub
Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long


   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "LU") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            If C.OldLine Then
               m_Combos(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Combos(C.ControlIndex).Top = PrevTop - m_Combos(C.ControlIndex - 1).Height
            Else
               m_Combos(C.ControlIndex).Left = PrevLeft
               m_Combos(C.ControlIndex).Top = PrevTop
            End If
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
            If C.OldLine Then
               PrevLeft = m_Combos(C.ControlIndex).Left - CurWidth - 20
            Else
               PrevLeft = m_Combos(C.ControlIndex).Left
            End If
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).Height
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            If C.OldLine Then
               m_Texts(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Texts(C.ControlIndex).Top = PrevTop - txtGeneric(0).Height
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
               m_Texts(C.ControlIndex).Width = C.Width
            Else
               m_Texts(C.ControlIndex).Left = PrevLeft
               m_Texts(C.ControlIndex).Top = PrevTop
               m_Texts(C.ControlIndex).Width = C.Width
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
                              
               CurTop = PrevTop
               CurLeft = PrevLeft
               CurWidth = PrevWidth
               
               PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).Height
               PrevLeft = m_Texts(C.ControlIndex).Left
               PrevWidth = C.Width
            End If
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).Height
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CH" Then
            m_Checks(C.ControlIndex).Left = PrevLeft
            m_Checks(C.ControlIndex).Top = PrevTop + 100
            m_Checks(C.ControlIndex).Width = C.Width
            m_Checks(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Checks(C.ControlIndex).Top + m_Checks(C.ControlIndex).Height
            PrevLeft = m_Checks(C.ControlIndex).Left
            PrevWidth = C.Width
         End If
   Else 'Label
         
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            If C.AllowNull Then
               Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            Else
               Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg, RGB(255, 0, 0))
            End If
            m_Labels(C.ControlIndex).Visible = True
   End If
   Next C
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_ReportParams = Nothing
   Set m_CheckBoxes = Nothing
   Set m_Rs = Nothing
   Set m_Checks = Nothing
   
   Set m_StockMarket = Nothing
   Set m_Rs1 = Nothing
   
   Call ReleaseAll
End Sub
Private Sub Timer1_Timer()
   
   Timer1.Enabled = False
   
   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   
   Timer1.Enabled = True
End Sub
Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean

   If LastKey = Node.Key Then
      Exit Sub
   End If

   pnlHeader.Caption = Node.Text
   
   Status = True
   QueryFlag = False

   Call UnloadAllControl
   
   cmdConfig.Visible = False
   cmdAdd.Visible = False
   SSFrame3.Visible = False
   If Node.Key = ROOT_TREE & " 1-0-1" Then
      Load frmChequeLog
      frmChequeLog.Show 1
      
      Unload frmChequeLog
      Set frmChequeLog = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-0-2" Then
      Call InitReport1_0_2
      SSFrame3.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-2-1" Then
      Load frmChequeConfig
      frmChequeConfig.Show 1
      
      Unload frmChequeConfig
      Set frmChequeConfig = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-2-2" Then
'      Load frmStockMarket
'      frmStockMarket.Show 1
'
'      Unload frmStockMarket
'      Set frmStockMarket = Nothing
   ElseIf Node.Key = ROOT_TREE1 & " -1" Then
      Load frmUserGroup
      frmUserGroup.Show 1
      
      Unload frmUserGroup
      Set frmUserGroup = Nothing
   ElseIf Node.Key = ROOT_TREE1 & " -2" Then
      Load frmUser
      frmUser.Show 1

      Unload frmUser
      Set frmUser = Nothing
   ElseIf Node.Key = ROOT_TREE3 Then
      frmMasterMain.m_TempArea = MASTER_BBANK
      frmMasterMain.HeaderText = Node.Text
      Load frmMasterMain
      frmMasterMain.Show 1

      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Node.Key = ROOT_TREE2 & " 1-1" Or Node.Key = ROOT_TREE2 & " 1-1-1" Or Node.Key = ROOT_TREE2 & " 1-1-2" Or Node.Key = ROOT_TREE2 & " 1-1-3" Then
      Call InitReportRoot2_1_1
      SSFrame3.Visible = True
      cmdConfig.Visible = True
   ElseIf Node.Key = ROOT_TREE2 & " 1-2" Or Node.Key = ROOT_TREE2 & " 1-3" Or Node.Key = ROOT_TREE2 & " 1-4" Then
      Call InitReportRoot2_1_2
      SSFrame3.Visible = True
      cmdConfig.Visible = True
   End If
End Sub
Private Sub LoadComboData()
Dim C As CReportControl

'   Me.Refresh
'   DoEvents
'   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "CB") Then
         
         If trvMain.SelectedItem.Key = ROOT_TREE & " 1-0-1" Or trvMain.SelectedItem.Key = ROOT_TREE & " 1-0-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadChequeConfig(m_Combos(C.ControlIndex))
            End If
         ElseIf trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-1" Or trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-1-1" Or trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-1-2" Or trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadChequeConfig(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMasterRef(m_Combos(C.ControlIndex), Nothing, MASTER_BBANK)
            ElseIf C.ComboLoadID = 3 Then
               Call InitPassFlag(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCancelFlag(m_Combos(C.ControlIndex))
               m_Combos(C.ControlIndex).ListIndex = 2    'เพื่อให้แสดงรายการไม่ยกเลิกเช็ค
            ElseIf C.ComboLoadID = 5 Then
               Call InitChequeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
            
         ElseIf trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-2" Or trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-3" Or trvMain.SelectedItem.Key = ROOT_TREE2 & " 1-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadMasterRef(m_Combos(C.ControlIndex), Nothing, MASTER_BBANK)
            ElseIf C.ComboLoadID = 2 Then
               Call InitPassFlag(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      End If 'C.ControlType = "C"
   Next C
'   Call EnableForm(Me, True)
End Sub
Private Sub Form_Resize()
On Error Resume Next
'On Error GoTo ErrorHandler
   SSFrame2.Width = ScaleWidth
   SSFrame2.Height = ScaleHeight
   
   SSPanel1.Width = ScaleWidth - 4395
   lblDateTime.Width = ScaleWidth - 4395
   
   If ScaleWidth > 0 Then
      'trvMain.Width = ScaleWidth - SSFrame3.Width
   End If
   
   If ScaleHeight > 0 Then
      cmdExit.Top = ScaleHeight - cmdExit.Height - 100
      cmdConfig.Top = cmdExit.Top
      cmdOK.Top = cmdExit.Top
      lblUsername.Top = cmdExit.Top + 100
   End If
   
   If ScaleWidth > 0 Then
      cmdOK.Left = ScaleWidth - cmdOK.Width - 40
      cmdConfig.Left = ScaleWidth - cmdOK.Width - 40 - cmdConfig.Width - 40
      
      trvMain.Height = cmdExit.Top - SSPanel1.Height - pnlHeader.Height - 100
      SSFrame3.Height = trvMain.Height
      
      SSFrame3.Left = trvMain.Width
      pnlHeader.Left = SSFrame3.Left
      SSPanel1.Left = pnlHeader.Left
      pnlHeader.Width = ScaleWidth - trvMain.Width
      SSPanel1.Width = pnlHeader.Width
      
   End If
'   Exit Sub
'ErrorHandler:
'   glbErrorLog.LocalErrorMsg = "Eror"
'   glbErrorLog.ShowUserError
End Sub

Private Sub InitReport1_0_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width * 1, False, "", , "PAY", , "PAY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จ่าย Pay"), , , , "PAY_NAME")
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("บาท Baht"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "TYPE", "CHEQUE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("แบบเช็ค"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่ Date"))
   
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("เลขที่เช็ค"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ A/C PAYEE ONLY", , "SHOW_AC")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ ขีดผู้ถือ (or bearer)", , "SHOW_BEARER")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "TYPE", "CHEQUE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("แบบเช็ค"))
   
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("เลขที่เช็ค"))
   
   Call LoadControl("T", txtGeneric(0).Width * 1, True, "", , "PAY", , "PAY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จ่าย Pay"), , , , "PAY_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่ลงวันที่เช็ค", , "NO_DUE_DATE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ A/C PAYEE ONLY", , "SHOW_AC")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ ขีดผู้ถือ (or bearer)", , "SHOW_BEARER")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportRoot2_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width * 1, True, "", , "PAY", , "PAY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จ่าย Pay"), , , , "PAY_NAME")
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FROM_CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเลขที่เช็ค"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TO_CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเลขที่เช็ค"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่พิมพ์"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่พิมพ์"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BANK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่BANK"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BANK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่BANK"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_STOCK_MARKET_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เช็ค"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_STOCK_MARKET_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เช็ค"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "TYPE", "CHEQUE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบบเช็ค"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BBANK_ID", "BBANK_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สาขาธนาคาร"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PASS_FLAG", "PASS_FLAG_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ผ่านเช็ค"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "CANCEL_FLAG", "CANCEL_FLAG_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ยกเลิกเช็ค"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE", "ORDER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง LOG การพิมพ์เช็ค", , "SHOW_PRINT_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง LOG การผ่านเช็ค", , "SHOW_PASS_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง LOG การยกเลิกเช็ค", , "SHOW_CANCEL_DETAIL")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportRoot2_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_STOCK_MARKET_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่เช็ค"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_STOCK_MARKET_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("ถึงวันที่เช็ค"))
   
   Call LoadControl("T", txtGeneric(0).Width * 1, True, "", , "PAY", , "PAY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จ่าย Pay"), , , , "PAY_NAME")
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FROM_CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเลขที่เช็ค"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TO_CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเลขที่เช็ค"))
      
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "BBANK_ID", "BBANK_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สาขาธนาคาร"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PASS_FLAG", "PASS_FLAG_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ผ่านเช็ค"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE", "ORDER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   End If
End Sub
