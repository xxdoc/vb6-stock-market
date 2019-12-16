VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMasterMain 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMasterMain.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         Top             =   7800
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10095
            TabIndex        =   4
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   525
            Left            =   1770
            TabIndex        =   2
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   525
            Left            =   150
            TabIndex        =   1
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   525
            Left            =   3420
            TabIndex        =   3
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2ABC
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1508
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":2DD6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":36B2
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2850
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":39CE
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6915
         Left            =   0
         TabIndex        =   0
         Top             =   900
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   12197
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
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
         Column(1)       =   "frmMasterMain.frx":3CE8
         Column(2)       =   "frmMasterMain.frx":3DB0
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMasterMain.frx":3E54
         FormatStyle(2)  =   "frmMasterMain.frx":3FB0
         FormatStyle(3)  =   "frmMasterMain.frx":4060
         FormatStyle(4)  =   "frmMasterMain.frx":4114
         FormatStyle(5)  =   "frmMasterMain.frx":41EC
         ImageCount      =   0
         PrinterProperties=   "frmMasterMain.frx":42A4
      End
   End
End
Attribute VB_Name = "frmMasterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_MasterRef As CMasterRef
Private m_MasterRef1 As CMasterRef
Public m_TempArea As MASTER_TYPE

Public HeaderText As String
Private m_FieldLists As Collection
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim MI As CMenuItem
   
   
   Set MI = m_FieldLists(Trim(Str(m_TempArea)))
   
   frmAddEditMaster1.KEY_CODE = MI.KEYWORD
   frmAddEditMaster1.KEY_NAME = MI.MENU_TEXT
   frmAddEditMaster1.MasterArea = m_TempArea
   frmAddEditMaster1.MasterKey = Trim(Str(m_TempArea))
   frmAddEditMaster1.ShowMode = SHOW_ADD
   frmAddEditMaster1.HeaderText = MapText("เพิ่มข้อมูล")
   Load frmAddEditMaster1
   frmAddEditMaster1.Show 1
   
   OKClick = frmAddEditMaster1.OKClick
   
   Unload frmAddEditMaster1
   Set frmAddEditMaster1 = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub AddMenuItem(KeyCode As String, KeyName As String, KEY As String)
Dim MI As CMenuItem

   Set MI = New CMenuItem
   MI.KEYWORD = KeyCode
   MI.MENU_TEXT = KeyName
   
   Call m_FieldLists.Add(MI, KEY)
   
   Set MI = Nothing
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
   
   Call EnableForm(Me, False)
   
   Call InitGrid(Trim(Str(m_TempArea)))
   
   If m_TempArea > 0 Then
      Dim Mr As CMasterRef
      Set Mr = New CMasterRef
      Mr.KEY_ID = -1
      Mr.MASTER_AREA = m_TempArea
      Call Mr.QueryData(1, m_Rs, ItemCount, True)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      Set Mr = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
Dim Status As Boolean
Dim IsOK As Boolean
Dim TempID As Long
   
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
      
   m_MasterRef.KEY_ID = TempID
   Status = glbDaily.DeleteMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   If Status Then
      Call QueryData(True)
   Else
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Exit Sub
   
ErrorHandler:
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim TempID As Long
Dim MI As CMenuItem
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   TempID = GridEX1.Value(1)
   
   Set MI = m_FieldLists(Trim(Str(m_TempArea)))
      
   frmAddEditMaster1.KEY_CODE = MI.KEYWORD
   frmAddEditMaster1.KEY_NAME = MI.MENU_TEXT
   frmAddEditMaster1.MasterArea = m_TempArea
   frmAddEditMaster1.ID = TempID
   frmAddEditMaster1.MasterKey = Trim(Str(m_TempArea))
   frmAddEditMaster1.ShowMode = SHOW_EDIT
   frmAddEditMaster1.HeaderText = MapText("แก้ไขข้อมูล")
   Load frmAddEditMaster1
   frmAddEditMaster1.Show 1
   
   OKClick = frmAddEditMaster1.OKClick
   
   Unload frmAddEditMaster1
   Set frmAddEditMaster1 = Nothing
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub Form_Activate()
Dim ItemCount As Long
   
   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
         
      If m_TempArea = MASTER_BBANK Then
         Call AddMenuItem(MapText("รหัสสาขาธนาคาร"), MapText("สาขาธนาคาร"), Trim(Str(m_TempArea)))
      End If
      
      Call QueryData(True)
      m_HasActivate = True
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
'      Call cmdOK_Click
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
   
   Set m_MasterRef = Nothing
   Set m_MasterRef1 = Nothing
   Set m_FieldLists = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid(KEY As String)
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle
Dim MI As CMenuItem

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   If (KEY <> "") And (KEY <> "Root") Then
      Set MI = m_FieldLists(KEY)
      
      Set Col = GridEX1.Columns.Add '1
      Col.Width = 0
      Col.Caption = "ID"
   
      Set Col = GridEX1.Columns.Add '2
      Col.Width = 2235
      Col.Caption = MI.KEYWORD
         
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 5100
      Col.Caption = MI.MENU_TEXT
      
   End If
   
   GridEX1.ItemCount = 0
   GridEX1.Rebind
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR

   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   Set m_MasterRef = New CMasterRef
   Set m_FieldLists = New Collection
   
   Call InitFormLayout
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   End If
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
   Call m_MasterRef.PopulateFromRS(1, m_Rs)
      
   Values(1) = m_MasterRef.KEY_ID
   Values(2) = m_MasterRef.KEY_CODE
   Values(3) = m_MasterRef.KEY_NAME
   
Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth
   GridEX1.Height = ScaleHeight - pnlHeader.Height - pnlFooter.Height
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.Height
   
   cmdExit.Left = ScaleWidth - cmdExit.Width - 20
   
End Sub

