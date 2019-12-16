VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmChequeConfig 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmChequeConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6855
         Left            =   180
         TabIndex        =   0
         Top             =   840
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   12091
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
         Column(1)       =   "frmChequeConfig.frx":27A2
         Column(2)       =   "frmChequeConfig.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmChequeConfig.frx":290E
         FormatStyle(2)  =   "frmChequeConfig.frx":2A6A
         FormatStyle(3)  =   "frmChequeConfig.frx":2B1A
         FormatStyle(4)  =   "frmChequeConfig.frx":2BCE
         FormatStyle(5)  =   "frmChequeConfig.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmChequeConfig.frx":2D5E
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   3
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmChequeConfig.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   1
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmChequeConfig.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   2
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
         TabIndex        =   5
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmChequeConfig.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmChequeConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_ChequeConfig As CChequeConfig
Private m_TempChequeConfig As CChequeConfig
Private m_Rs As ADODB.Recordset
Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditChequeConfig.HeaderText = MapText("เพิ่มข้อมูล")
   frmAddEditChequeConfig.ShowMode = SHOW_ADD
   Load frmAddEditChequeConfig
   frmAddEditChequeConfig.Show 1

   OKClick = frmAddEditChequeConfig.OKClick

   Unload frmAddEditChequeConfig
   Set frmAddEditChequeConfig = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim KEY As String

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   KEY = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(1)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   
   m_ChequeConfig.CHEQUE_NAME = KEY
   Call m_ChequeConfig.DeleteData
   
   Call QueryData(True)
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim KEY As String
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   KEY = GridEX1.Value(1)
   
   frmAddEditChequeConfig.KEY = KEY
   frmAddEditChequeConfig.HeaderText = MapText("เแก้ไขข้อมูลตั้งค่าเช็ค")
   frmAddEditChequeConfig.ShowMode = SHOW_EDIT
   Load frmAddEditChequeConfig
   frmAddEditChequeConfig.Show 1

   OKClick = frmAddEditChequeConfig.OKClick

   Unload frmAddEditChequeConfig
   Set frmAddEditChequeConfig = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
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
      
      m_ChequeConfig.CHEQUE_NAME = ""
      If Not glbDaily.QueryChequeConfig(m_ChequeConfig, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   
   GridEX1.ItemCount = m_Rs.RecordCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
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
      'Call cmdClear_Click
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
   Col.Width = 4000
   Col.Caption = MapText("ชื่อ")
      
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 4000
   Col.Caption = MapText("สาขาธนาคาร")
      
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลการตั้งค่าเช็ค")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   
   Set m_ChequeConfig = New CChequeConfig
   Set m_TempChequeConfig = New CChequeConfig
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(1)
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
   Call m_TempChequeConfig.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempChequeConfig.CHEQUE_NAME
   Values(2) = m_TempChequeConfig.BBANK_NAME
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = ScaleWidth - cmdExit.Width - 50 - cmdOK.Width - 50
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   End If
End Sub
Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempKey1 As String
Dim TempID2 As Long
Dim CC As CChequeConfig
Dim IsOK As Boolean
Dim OKClick As Boolean
   
   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempKey1 = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("COPY")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set CC = New CChequeConfig
      CC.CHEQUE_NAME = TempKey1
      Call glbDaily.CopyChequeConfig(CC, IsOK, True, glbErrorLog)
      Call QueryData(True)
      Set CC = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

