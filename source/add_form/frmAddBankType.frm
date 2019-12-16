VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddBankType 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmAddBankType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   9763
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   -120
         TabIndex        =   2
         Top             =   0
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSOption OpNakon 
         Height          =   495
         Left            =   2880
         TabIndex        =   10
         Top             =   3720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption OpUOB 
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   3720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption OpKrungSri 
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   2880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption OpThanthai 
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption OpGsikornthai 
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption OpKrungThep 
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption OpThaiPanit 
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption Opkrungthai 
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2040
         TabIndex        =   0
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddBankType.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddBankType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HeaderText As String
Public OptionBankTypeID As SHOW_BANK_TYPE
Public m_HasModify  As Boolean
Public OKClick  As Boolean
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   SaveData = False
   
   SaveData = True
End Function
Private Sub Form_Activate()
   Me.Refresh
   DoEvents
   OKClick = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "เลือกธนาคาร"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitOptionEx(Opkrungthai, "ธนาคารกรุงไทย")
   Call InitOptionEx(OpKrungThep, "ธนาคารกรุงเทพ")
   Call InitOptionEx(OpGsikornthai, "ธนาคารกสิกรไทย")
   Call InitOptionEx(OpUOB, "ธนาคารยูโอบี")
   Call InitOptionEx(OpThaiPanit, "ธนาคารไทยพาณิชย์")
   Call InitOptionEx(OpKrungSri, "ธนาคารกรุงศรีอยุธยา")
   Call InitOptionEx(OpThanthai, "ธนาคารทหารไทย")
   Call InitOptionEx(OpNakon, "ธนาคารนครหลวงไทย")
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
   Opkrungthai.Value = True
End Sub
Private Sub Form_Load()
   Call InitFormLayout
End Sub
Private Sub SubMapOption2ID()
   If Opkrungthai.Value Then
      OptionBankTypeID = KRUNGTHAI
   ElseIf OpKrungThep.Value Then
      OptionBankTypeID = KRUNGTHEP
   ElseIf OpGsikornthai.Value Then
      OptionBankTypeID = GSIKORNTHAI
   ElseIf OpUOB.Value Then
      OptionBankTypeID = UOB
   ElseIf OpThaiPanit.Value Then
      OptionBankTypeID = THAIPANIT
   ElseIf OpKrungSri.Value Then
      OptionBankTypeID = KRUNGSRI
   ElseIf OpThanthai.Value Then
      OptionBankTypeID = TMB
   ElseIf OpNakon.Value Then
      OptionBankTypeID = NAKORN
   End If
End Sub

Private Sub OpGsikornthai_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub OpKrungSri_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub Opkrungthai_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub OpKrungThep_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub OpNakon_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub OpThaiPanit_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub OpThanthai_Click(Value As Integer)
   Call SubMapOption2ID
End Sub

Private Sub OpUOB_Click(Value As Integer)
   Call SubMapOption2ID
End Sub
