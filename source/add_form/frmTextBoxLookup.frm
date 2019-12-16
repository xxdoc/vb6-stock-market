VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTextBoxLookup 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmTextBoxLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12091
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5445
         Left            =   45
         TabIndex        =   1
         Top             =   1320
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   9604
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
         Column(1)       =   "frmTextBoxLookup.frx":27A2
         Column(2)       =   "frmTextBoxLookup.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmTextBoxLookup.frx":290E
         FormatStyle(2)  =   "frmTextBoxLookup.frx":2A6A
         FormatStyle(3)  =   "frmTextBoxLookup.frx":2B1A
         FormatStyle(4)  =   "frmTextBoxLookup.frx":2BCE
         FormatStyle(5)  =   "frmTextBoxLookup.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmTextBoxLookup.frx":2D5E
      End
      Begin StockMarket.uctlTextBox txtSearchText 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   767
      End
      Begin Threed.SSOption SSOption2 
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin VB.Label lblSearchText 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   4
         Top             =   900
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmTextBoxLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Rs As ADODB.Recordset
Public KEYWORD As String
Public KeySearch As String

Public OKClick As Boolean
Public HeaderText As String
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = ScaleWidth / 2
   Col.Caption = MapText("รายละเอียด1")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = ScaleWidth / 2
   Col.Caption = MapText("รายละเอียด2")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   
   Call InitNormalLabel(lblSearchText, MapText("KEY SEARCH"))
   
   txtSearchText.Text = KEYWORD
   txtSearchText.Enabled = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitOptionEx(SSOption1, "เลือกคอลัมน์ 1")
   Call InitOptionEx(SSOption2, "เลือกคอลัมน์ 2")
   SSOption1.Value = True
End Sub
Private Sub Form_Load()
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call ReturnKeyWord
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      KeyCode = 0
      Unload Me
   ElseIf KeyCode = 13 Or KeyCode = 32 Then
      Call ReturnKeyWord
   End If
End Sub
Private Sub ReturnKeyWord()
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
      
   If SSOption1 Then
      KEYWORD = GridEX1.Value(1)
   Else
      KEYWORD = GridEX1.Value(2)
   End If
   Unload Me
End Sub
