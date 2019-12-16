VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5A3A1-7FFE-11D4-A13A-004005FA6275}#1.0#0"; "ImagXpr6.dll"
Begin VB.Form frmAddEditChequeConfig 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmAddEditChequeConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11865
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBBank 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2955
      End
      Begin VB.ComboBox cboDateType 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1260
         Width           =   2955
      End
      Begin Threed.SSFrame SSFraLogo1 
         Height          =   1335
         Left            =   120
         TabIndex        =   49
         Top             =   3480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2355
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin StockMarket.uctlTextBox txtLogo1Path 
            Height          =   400
            Left            =   1200
            TabIndex        =   21
            Top             =   720
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo1Left 
            Height          =   400
            Left            =   1200
            TabIndex        =   17
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo1Top 
            Height          =   400
            Left            =   3120
            TabIndex        =   18
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo1Height 
            Height          =   405
            Left            =   5520
            TabIndex        =   19
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo1Weight 
            Height          =   405
            Left            =   7440
            TabIndex        =   20
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin VB.Label lblLogo1Weight 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   495
            Left            =   6120
            TabIndex        =   66
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLogo1Height 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   375
            Left            =   4440
            TabIndex        =   65
            Top             =   360
            Width           =   975
         End
         Begin Threed.SSCommand cmdFileName1 
            Height          =   400
            Left            =   8900
            TabIndex        =   22
            Top             =   740
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditChequeConfig.frx":27A2
            ButtonStyle     =   3
         End
         Begin VB.Label lblLogo1Path 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblLogo1Left 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   495
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLogo1Top 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   495
            Left            =   1800
            TabIndex        =   50
            Top             =   360
            Width           =   1215
         End
      End
      Begin IMAGXPR6LibCtl.ImagXpress ImagLogo1 
         Height          =   1815
         Left            =   9720
         TabIndex        =   53
         Top             =   2985
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         ErrStr          =   "F4NRO2IK2AP-ER3063PXEP"
         ErrCode         =   432366639
         ErrInfo         =   1515989874
         Persistence     =   -1  'True
         _cx             =   1
         _cy             =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ScrollBarLargeChangeH=   10
         ScrollBarSmallChangeH=   1
         SaveGIFInterlaced=   -1  'True
         SaveGIFTransparent=   -1  'True
         OLEDropMode     =   0
         ScrollBarLargeChangeV=   10
         ScrollBarSmallChangeV=   1
         DisplayProgressive=   -1  'True
         SaveTIFByteOrder=   0
         LoadRotated     =   0
         FTPUserName     =   ""
         FTPPassword     =   ""
         ProxyServer     =   ""
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtName 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   3975
         _ExtentX        =   14843
         _ExtentY        =   661
      End
      Begin StockMarket.uctlTextBox txtDateleft 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1260
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
      End
      Begin StockMarket.uctlTextBox txtDateTop 
         Height          =   405
         Left            =   3360
         TabIndex        =   4
         Top             =   1260
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtThaiLeft 
         Height          =   405
         Left            =   5400
         TabIndex        =   8
         Top             =   1740
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtThaiTop 
         Height          =   405
         Left            =   7440
         TabIndex        =   9
         Top             =   1740
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtPayLeft 
         Height          =   405
         Left            =   1440
         TabIndex        =   6
         Top             =   1740
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtPayTop 
         Height          =   405
         Left            =   3360
         TabIndex        =   7
         Top             =   1740
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtAmountLeft 
         Height          =   405
         Left            =   9360
         TabIndex        =   10
         Top             =   1740
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtAmountTop 
         Height          =   405
         Left            =   11160
         TabIndex        =   11
         Top             =   1740
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin Threed.SSFrame SSFraLogo2 
         Height          =   1335
         Left            =   120
         TabIndex        =   54
         Top             =   5400
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2355
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin StockMarket.uctlTextBox txtLogo2Path 
            Height          =   400
            Left            =   1200
            TabIndex        =   28
            Top             =   720
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo2Left 
            Height          =   400
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo2Top 
            Height          =   400
            Left            =   3120
            TabIndex        =   25
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo2Height 
            Height          =   405
            Left            =   5520
            TabIndex        =   26
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin StockMarket.uctlTextBox txtLogo2Weight 
            Height          =   405
            Left            =   7440
            TabIndex        =   27
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
         End
         Begin VB.Label lblLogo2Weight 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   495
            Left            =   6120
            TabIndex        =   68
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLogo2Height 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   375
            Left            =   4440
            TabIndex        =   67
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLogo2Top 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   495
            Left            =   1800
            TabIndex        =   57
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLogo2Left 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   495
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLogo2Path 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   840
            Width           =   975
         End
         Begin Threed.SSCommand cmdFileName2 
            Height          =   400
            Left            =   8900
            TabIndex        =   29
            Top             =   740
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditChequeConfig.frx":2ABC
            ButtonStyle     =   3
         End
      End
      Begin IMAGXPR6LibCtl.ImagXpress ImagLogo2 
         Height          =   1815
         Left            =   9720
         TabIndex        =   58
         Top             =   4905
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         ErrStr          =   "F4NRO2IK2AP-ER3063PXEP"
         ErrCode         =   432366639
         ErrInfo         =   1515989874
         Persistence     =   -1  'True
         _cx             =   1
         _cy             =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ScrollBarLargeChangeH=   10
         ScrollBarSmallChangeH=   1
         OLEDropMode     =   0
         ScrollBarLargeChangeV=   10
         ScrollBarSmallChangeV=   1
         DisplayProgressive=   -1  'True
         SaveTIFByteOrder=   0
         LoadRotated     =   0
         FTPUserName     =   ""
         FTPPassword     =   ""
         ProxyServer     =   ""
      End
      Begin StockMarket.uctlTextBox txtName1Desc 
         Height          =   405
         Left            =   1920
         TabIndex        =   34
         Top             =   7320
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtName1DescLeft 
         Height          =   405
         Left            =   1920
         TabIndex        =   30
         Top             =   6840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtName1DescTop 
         Height          =   405
         Left            =   4800
         TabIndex        =   31
         Top             =   6840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtName2Desc 
         Height          =   405
         Left            =   7320
         TabIndex        =   35
         Top             =   7320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtName2DescLeft 
         Height          =   405
         Left            =   7320
         TabIndex        =   32
         Top             =   6840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtName2DescTop 
         Height          =   405
         Left            =   10200
         TabIndex        =   33
         Top             =   6840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin StockMarket.uctlTextBox txtAcLeft 
         Height          =   405
         Left            =   1440
         TabIndex        =   12
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtACTop 
         Height          =   405
         Left            =   3360
         TabIndex        =   13
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtBearerLeft 
         Height          =   405
         Left            =   5400
         TabIndex        =   14
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin StockMarket.uctlTextBox txtBearerTop 
         Height          =   405
         Left            =   7440
         TabIndex        =   15
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin Threed.SSCommand cmdAddBankType 
         Height          =   525
         Left            =   9960
         TabIndex        =   2
         Top             =   720
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeConfig.frx":2DD6
         ButtonStyle     =   3
      End
      Begin VB.Label lblBearerLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   4080
         TabIndex        =   74
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label lblBearerTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   5880
         TabIndex        =   73
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label lblACLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   0
         TabIndex        =   72
         Top             =   2340
         Width           =   1335
      End
      Begin VB.Label lblACTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   2040
         TabIndex        =   71
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label lblBBank 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   5400
         TabIndex        =   70
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lblDateType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   5160
         TabIndex        =   69
         Top             =   1320
         Width           =   1695
      End
      Begin Threed.SSCheck ChkLogo2 
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   4920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck ChkLogo1 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblName2DescTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   8400
         TabIndex        =   64
         Top             =   6960
         Width           =   1695
      End
      Begin VB.Label lblName2DescLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   5760
         TabIndex        =   63
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label lblName2Desc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   5640
         TabIndex        =   62
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Label lblName1DescTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   3000
         TabIndex        =   61
         Top             =   6960
         Width           =   1695
      End
      Begin VB.Label lblName1DescLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   240
         TabIndex        =   60
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Label lblName1Desc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Label lblAmountTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   10080
         TabIndex        =   48
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblAmountLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   8040
         TabIndex        =   47
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPayTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   2040
         TabIndex        =   46
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPayLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   0
         TabIndex        =   45
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblThaiTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   5880
         TabIndex        =   44
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblThaiLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   4080
         TabIndex        =   43
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblDateTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   2040
         TabIndex        =   42
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDateleft 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   0
         TabIndex        =   41
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   -240
         TabIndex        =   40
         Top             =   960
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10035
         TabIndex        =   37
         Top             =   7830
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8385
         TabIndex        =   36
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeConfig.frx":30F0
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditChequeConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_ChequeConfig As CChequeConfig

Public KEY As String
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cboBBank_Click()
   m_HasModify = True
End Sub

Private Sub cboDateType_Click()
   m_HasModify = True
End Sub

Private Sub ChkLogo1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ChkLogo2_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAddBankType_Click()
Dim ID As SHOW_BANK_TYPE
   
      Load frmAddBankType
      frmAddBankType.Show 1
      
      Unload frmAddBankType
      If frmAddBankType.OKClick Then
         ID = frmAddBankType.OptionBankTypeID
         txtDateleft.Text = BankType2Value(ID, 1)
         txtDateTop.Text = BankType2Value(ID, 2)
         cboDateType.ListIndex = IDToListIndex(cboDateType, BankType2Value(ID, 3))
         txtPayLeft.Text = BankType2Value(ID, 4)
         txtPayTop.Text = BankType2Value(ID, 5)
         
         txtThaiLeft.Text = BankType2Value(ID, 6)
         txtThaiTop.Text = BankType2Value(ID, 7)
         txtAmountLeft.Text = BankType2Value(ID, 8)
         txtAmountTop.Text = BankType2Value(ID, 9)
         
         txtAcLeft.Text = BankType2Value(ID, 10)
         txtACTop.Text = BankType2Value(ID, 11)
         txtBearerLeft.Text = BankType2Value(ID, 12)
         txtBearerTop.Text = BankType2Value(ID, 13)
         
      End If
      Set frmAddBankType = Nothing
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_ChequeConfig.CHEQUE_NAME = KEY
      If Not glbDaily.QueryChequeConfig(m_ChequeConfig, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If m_Rs.RecordCount > 0 Then
      Call m_ChequeConfig.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_ChequeConfig.CHEQUE_NAME
      txtDateleft.Text = m_ChequeConfig.DATE_LEFT
      txtDateTop.Text = m_ChequeConfig.DATE_TOP
      cboDateType.ListIndex = IDToListIndex(cboDateType, m_ChequeConfig.DATE_TYPE)
      txtPayLeft.Text = m_ChequeConfig.PAY_LEFT
      txtPayTop.Text = m_ChequeConfig.PAY_TOP
      txtThaiLeft.Text = m_ChequeConfig.THAI_LEFT
      txtThaiTop.Text = m_ChequeConfig.THAI_TOP
      txtAmountLeft.Text = m_ChequeConfig.AMOUNT_LEFT
      txtAmountTop.Text = m_ChequeConfig.AMOUNT_TOP
      ChkLogo1.Value = FlagToCheck(m_ChequeConfig.LOGO1_FLAG)
      txtLogo1Left.Text = m_ChequeConfig.LOGO1_LEFT
      txtLogo1Top.Text = m_ChequeConfig.LOGO1_TOP
      txtLogo1Path.Text = m_ChequeConfig.LOGO1_PATH
      
      ChkLogo2.Value = FlagToCheck(m_ChequeConfig.LOGO2_FLAG)
      txtLogo2Left.Text = m_ChequeConfig.LOGO2_LEFT
      txtLogo2Top.Text = m_ChequeConfig.LOGO2_TOP
      txtLogo2Path.Text = m_ChequeConfig.LOGO2_PATH
      
      txtName1Desc.Text = m_ChequeConfig.NAME1_DESC
      txtName1DescLeft.Text = m_ChequeConfig.NAME1_DESC_LEFT
      txtName1DescTop.Text = m_ChequeConfig.NAME1_DESC_TOP
      
      txtName2Desc.Text = m_ChequeConfig.NAME2_DESC
      txtName2DescLeft.Text = m_ChequeConfig.NAME2_DESC_LEFT
      txtName2DescTop.Text = m_ChequeConfig.NAME2_DESC_TOP
      
      ImagLogo1.ZoomToFit ZOOMFIT_BEST
      ImagLogo2.ZoomToFit ZOOMFIT_BEST
      
      txtLogo1Height.Text = m_ChequeConfig.LOGO1_HEIGHT
      txtLogo1Weight.Text = m_ChequeConfig.LOGO1_WEIGHT
      txtLogo2Height.Text = m_ChequeConfig.LOGO2_HEIGHT
      txtLogo2Weight.Text = m_ChequeConfig.LOGO2_WEIGHT
      
      txtAcLeft.Text = m_ChequeConfig.AC_LEFT
      txtACTop.Text = m_ChequeConfig.AC_TOP
      txtBearerLeft.Text = m_ChequeConfig.BEARER_LEFT
      txtBearerTop.Text = m_ChequeConfig.BEARER_TOP
      
      cboBBank.ListIndex = IDToListIndex(cboBBank, m_ChequeConfig.BBANK_ID)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   'เช็คว่ามีชื่อ หรือยังเพราะห้ามซ้ำ
   If m_ChequeConfig.PREV_NAME <> txtName.Text Then
      If Not CheckUniqueNsKey(CONFIG_CHEQUE, txtName.Text) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   m_ChequeConfig.AddEditMode = ShowMode
   
   m_ChequeConfig.CHEQUE_NAME = txtName.Text
   m_ChequeConfig.DATE_LEFT = Val(txtDateleft.Text)
   m_ChequeConfig.DATE_TOP = Val(txtDateTop.Text)
   m_ChequeConfig.DATE_TYPE = cboDateType.ItemData(Minus2Zero(cboDateType.ListIndex))
   m_ChequeConfig.PAY_LEFT = Val(txtPayLeft.Text)
   m_ChequeConfig.PAY_TOP = Val(txtPayTop.Text)
   m_ChequeConfig.THAI_LEFT = Val(txtThaiLeft.Text)
   m_ChequeConfig.THAI_TOP = Val(txtThaiTop.Text)
   m_ChequeConfig.AMOUNT_LEFT = Val(txtAmountLeft.Text)
   m_ChequeConfig.AMOUNT_TOP = Val(txtAmountTop.Text)
   m_ChequeConfig.LOGO1_FLAG = Check2Flag(ChkLogo1.Value)
   m_ChequeConfig.LOGO1_LEFT = Val(txtLogo1Left.Text)
   m_ChequeConfig.LOGO1_TOP = Val(txtLogo1Top.Text)
   m_ChequeConfig.LOGO1_PATH = txtLogo1Path.Text
   
   m_ChequeConfig.LOGO2_FLAG = Check2Flag(ChkLogo2.Value)
   m_ChequeConfig.LOGO2_LEFT = Val(txtLogo2Left.Text)
   m_ChequeConfig.LOGO2_TOP = Val(txtLogo2Top.Text)
   m_ChequeConfig.LOGO2_PATH = txtLogo2Path.Text
   
   m_ChequeConfig.NAME1_DESC = txtName1Desc.Text
   m_ChequeConfig.NAME1_DESC_LEFT = Val(txtName1DescLeft.Text)
   m_ChequeConfig.NAME1_DESC_TOP = Val(txtName1DescTop.Text)
   
   m_ChequeConfig.NAME2_DESC = txtName2Desc.Text
   m_ChequeConfig.NAME2_DESC_LEFT = Val(txtName2DescLeft.Text)
   m_ChequeConfig.NAME2_DESC_TOP = Val(txtName2DescTop.Text)
   
   m_ChequeConfig.LOGO1_HEIGHT = Val(txtLogo1Height.Text)
   m_ChequeConfig.LOGO1_WEIGHT = Val(txtLogo1Weight.Text)
   m_ChequeConfig.LOGO2_HEIGHT = Val(txtLogo2Height.Text)
   m_ChequeConfig.LOGO2_WEIGHT = Val(txtLogo2Weight.Text)
   
   m_ChequeConfig.AC_LEFT = txtAcLeft.Text
   m_ChequeConfig.AC_TOP = txtACTop.Text
   m_ChequeConfig.BEARER_LEFT = txtBearerLeft.Text
   m_ChequeConfig.BEARER_TOP = txtBearerTop.Text
   
   m_ChequeConfig.BBANK_ID = cboBBank.ItemData(Minus2Zero(cboBBank.ListIndex))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditChequeConfig(m_ChequeConfig, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitChequeDateType(cboDateType)
      Call LoadMasterRef(cboBBank, Nothing, MASTER_BBANK)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         KEY = ""
      End If
      
      m_HasModify = False
   End If
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
   SSFraLogo1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFraLogo2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption & "                      (L แทนกั้นหน้าซ้าย)   (T แทนกั้นบน)"
   
   Call InitNormalLabel(lblName, MapText("ชื่อแบบเช็ค"))
   
   Call InitNormalLabel(lblDateleft, MapText("วันที่ L"))
   Call InitNormalLabel(lblDateTop, MapText("วันที่ T"))
   Call InitNormalLabel(lblDateType, MapText("แบบวันที่"))
   
   Call InitNormalLabel(lblPayLeft, MapText("ชื่อบัญชี L"))
   Call InitNormalLabel(lblPayTop, MapText("ชื่อบัญชี T"))
   Call InitNormalLabel(lblThaiLeft, MapText("เงินสดไทย L"))
   Call InitNormalLabel(lblThaiTop, MapText("เงินสดไทย T"))
   Call InitNormalLabel(lblAmountLeft, MapText("ยอดเงิน L"))
   Call InitNormalLabel(lblAmountTop, MapText("ยอดเงิน T"))
   
   Call InitNormalLabel(lblLogo1Left, MapText("Logo1 L"))
   Call InitNormalLabel(lblLogo1Top, MapText("Logo1 T"))
   Call InitNormalLabel(lblLogo1Path, MapText("ที่อยู่ Logo1"))
   
   Call InitNormalLabel(lblLogo2Left, MapText("Logo2 L"))
   Call InitNormalLabel(lblLogo2Top, MapText("Logo2 T"))
   Call InitNormalLabel(lblLogo2Path, MapText("ที่อยู่ Logo2"))
      
   Call InitNormalLabel(lblLogo1Height, MapText("Logo1 H"))
   Call InitNormalLabel(lblLogo1Weight, MapText("Logo1 W"))
   Call InitNormalLabel(lblLogo2Height, MapText("Logo2 H"))
   Call InitNormalLabel(lblLogo2Weight, MapText("Logo2 W"))
   
   Call InitNormalLabel(lblName1DescLeft, MapText("รายละเอียด 1 L"))
   Call InitNormalLabel(lblName1DescTop, MapText("รายละเอียด 1 T"))
   Call InitNormalLabel(lblName1Desc, MapText("รายละเอียด 1"))
   
   Call InitNormalLabel(lblName2DescLeft, MapText("รายละเอียด 2 L"))
   Call InitNormalLabel(lblName2DescTop, MapText("รายละเอียด 2 T"))
   Call InitNormalLabel(lblName2Desc, MapText("รายละเอียด 2"))
   
   Call InitNormalLabel(lblBBank, MapText("สาขาธนาคาร"))
   
   Call InitNormalLabel(lblACLeft, MapText("AC L"))
   Call InitNormalLabel(lblACTop, MapText("AC T"))
   Call InitNormalLabel(lblBearerLeft, MapText("ผู้ถือ L"))
   Call InitNormalLabel(lblBearerTop, MapText("ผู้ถือ T"))
   
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Call InitCheckBox(ChkLogo1, "แสดง LOGO 1")
   Call InitCheckBox(ChkLogo2, "แสดง LOGO 2")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitCombo(cboDateType)
   Call InitCombo(cboBBank)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAddBankType.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdFileName1, MapText(".B."))
   Call InitMainButton(cmdFileName2, MapText(".B."))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAddBankType, MapText("เพิ่มธนาคาร"))
   
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_ChequeConfig = New CChequeConfig
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub cmdFileName1_Click()
On Error Resume Next
Dim strDescription As String
Dim ID As Long
Dim MyName As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.JPG)|*.JPG;"
   'dlgAdd.Filter = "Access Files (*.GIF,*.JPG)|*.GIF;*.JPG;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
      
   txtLogo1Path.Text = dlgAdd.FileName
   ImagLogo1.ZoomToFit ZOOMFIT_BEST
   m_HasModify = True

End Sub
Private Sub cmdFileName2_Click()
On Error Resume Next
Dim strDescription As String
Dim ID As Long
Dim MyName As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.JPG)|*.JPG;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtLogo2Path.Text = dlgAdd.FileName
   ImagLogo2.ZoomToFit ZOOMFIT_BEST
   m_HasModify = True

End Sub
Private Sub Imaglogo1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ขนาดดีที่สุด", "-", "ตามสูง", "-", "ตามกว้าง", "-", "25%", "-", "50%", "-", "75%", "-", "100%", "-", "125", "-", "150", "-", "175", "-", "200")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
      
      If lMenuChosen = 1 Then
         ' Note: This will be the equivalent of ZOOMFIT_HEIGHT ot ZOOMFIT_WIDTH
         ' depending on which one fits the entire image within the control.
         ImagLogo1.ZoomToFit ZOOMFIT_BEST
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & Imaglogo1.IPZoomF
      ElseIf lMenuChosen = 3 Then
         ImagLogo1.ZoomToFit ZOOMFIT_HEIGHT
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & Imaglogo1.IPZoomF
      ElseIf lMenuChosen = 5 Then
         ImagLogo1.ZoomToFit ZOOMFIT_WIDTH
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & Imaglogo1.IPZoomF
      ElseIf lMenuChosen = 7 Then
         ImagLogo1.Zoom 0.25
      ElseIf lMenuChosen = 9 Then
         ImagLogo1.Zoom 0.5
      ElseIf lMenuChosen = 11 Then
         ImagLogo1.Zoom 0.75
      ElseIf lMenuChosen = 13 Then
         ImagLogo1.Zoom 1
      ElseIf lMenuChosen = 15 Then
         ImagLogo1.Zoom 1.25
      ElseIf lMenuChosen = 17 Then
         ImagLogo1.Zoom 1.5
      ElseIf lMenuChosen = 19 Then
         ImagLogo1.Zoom 1.75
      ElseIf lMenuChosen = 21 Then
         ImagLogo1.Zoom 2
      End If
      
      Call txtlogo1path_Change
   End If
End Sub
Private Sub Imaglogo2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ขนาดดีที่สุด", "-", "ตามสูง", "-", "ตามกว้าง", "-", "25%", "-", "50%", "-", "75%", "-", "100%", "-", "125", "-", "150", "-", "175", "-", "200")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
      
      If lMenuChosen = 1 Then
         ' Note: This will be the equivalent of ZOOMFIT_HEIGHT ot ZOOMFIT_WIDTH
         ' depending on which one fits the entire image within the control.
         ImagLogo2.ZoomToFit ZOOMFIT_BEST
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & ImagLogo2.IPZoomF
      ElseIf lMenuChosen = 3 Then
         ImagLogo2.ZoomToFit ZOOMFIT_HEIGHT
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & ImagLogo2.IPZoomF
      ElseIf lMenuChosen = 5 Then
         ImagLogo2.ZoomToFit ZOOMFIT_WIDTH
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & ImagLogo2.IPZoomF
      ElseIf lMenuChosen = 7 Then
         ImagLogo2.Zoom 0.25
      ElseIf lMenuChosen = 9 Then
         ImagLogo2.Zoom 0.5
      ElseIf lMenuChosen = 11 Then
         ImagLogo2.Zoom 0.75
      ElseIf lMenuChosen = 13 Then
         ImagLogo2.Zoom 1
      ElseIf lMenuChosen = 15 Then
         ImagLogo2.Zoom 1.25
      ElseIf lMenuChosen = 17 Then
         ImagLogo2.Zoom 1.5
      ElseIf lMenuChosen = 19 Then
         ImagLogo2.Zoom 1.75
      ElseIf lMenuChosen = 21 Then
         ImagLogo2.Zoom 2
      End If
      
      Call txtlogo2path_Change
   End If
End Sub

Private Sub txtAcLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtACTop_Change()
   m_HasModify = True
End Sub

Private Sub txtAmountLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtAmountTop_Change()
   m_HasModify = True
End Sub

Private Sub txtBearerLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtBearerTop_Change()
   m_HasModify = True
End Sub

Private Sub txtDateleft_Change()
   m_HasModify = True
End Sub

Private Sub txtDateTop_Change()
   m_HasModify = True
End Sub

Private Sub txtLogo1Height_Change()
   m_HasModify = True
End Sub

Private Sub txtLogo1Left_Change()
   m_HasModify = True
End Sub

Private Sub txtlogo1path_Change()
On Error GoTo Errorhanderor
   m_HasModify = True
   
   
   'If ShowMode = SHOW_ADD Then
      'ImagLogo1.CancelLoad = True
      'ImagLogo1.DeleteSaveBuffer
      
      ImagLogo1.FileName = txtLogo1Path.Text
      
'   Else
'      Imaglogo1.FileName = glbParameterObj.MapDrivePicture & txtPath.Text
'   End If
   
   Exit Sub
Errorhanderor:
   glbErrorLog.LocalErrorMsg = "Error หารูปไม่พบ"
   glbErrorLog.ShowUserError
End Sub

Private Sub txtLogo1Top_Change()
   m_HasModify = True
End Sub

Private Sub txtLogo1Weight_Change()
   m_HasModify = True
End Sub

Private Sub txtLogo2Height_Change()
   m_HasModify = True
End Sub

Private Sub txtLogo2Left_Change()
   m_HasModify = True
End Sub

Private Sub txtlogo2path_Change()
On Error GoTo Errorhanderor
   m_HasModify = True
   
   'If ShowMode = SHOW_ADD Then
      'ImagLogo2.CancelLoad = True
      'ImagLogo2.DeleteSaveBuffer
      
      ImagLogo2.FileName = txtLogo2Path.Text
      
'   Else
'      Imaglogo1.FileName = glbParameterObj.MapDrivePicture & txtPath.Text
'   End If
   
   Exit Sub
Errorhanderor:
   glbErrorLog.LocalErrorMsg = "Error หารูปไม่พบ"
   glbErrorLog.ShowUserError
End Sub

Private Sub txtLogo2Top_Change()
   m_HasModify = True
End Sub

Private Sub txtLogo2Weight_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtName1Desc_Change()
   m_HasModify = True
End Sub

Private Sub txtName1DescLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtName1DescTop_Change()
   m_HasModify = True
End Sub

Private Sub txtName2Desc_Change()
   m_HasModify = True
End Sub

Private Sub txtName2DescLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtName2DescTop_Change()
   m_HasModify = True
End Sub

Private Sub txtPayLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtPayTop_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiTop_Change()
   m_HasModify = True
End Sub
