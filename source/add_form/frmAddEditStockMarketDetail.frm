VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditStockMarketDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditStockMarketDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   6800
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin StockMarket.uctlTextBox txtAmount 
         Height          =   435
         Left            =   2760
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtOrderNo 
         Height          =   435
         Left            =   2760
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtPayName 
         Height          =   435
         Left            =   2760
         TabIndex        =   1
         Top             =   660
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtPrice 
         Height          =   435
         Left            =   6360
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtChequeNo 
         Height          =   435
         Left            =   2760
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtChequeName 
         Height          =   435
         Left            =   2760
         TabIndex        =   5
         Top             =   1980
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   767
      End
      Begin VB.Label lblChequeNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   1845
      End
      Begin VB.Label lblChequeName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   1980
         Width           =   1845
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label lblPayName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblOrderNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2325
         TabIndex        =   6
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3975
         TabIndex        =   7
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5625
         TabIndex        =   8
         Top             =   2670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   1200
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditStockMarketDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_StockMarketDetail As CStockMarketDetail

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_Products As Collection
Private m_Locations As Collection

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblOrderNo, MapText("ลำดับ"))
   Call InitNormalLabel(lblPayName, MapText("ชื่อผู้ถือหุ้น"))
   Call InitNormalLabel(lblAmount, MapText("จำนวนหุ้น"))
   Call InitNormalLabel(lblPrice, MapText("ผลตอบแทน"))
   Call InitNormalLabel(lblChequeNo, MapText("เลขที่เช็ค"))
   Call InitNormalLabel(lblChequeName, MapText("ชื่อจ่ายแทน"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CStockMarketDetail
         
         Set BD = TempCollection.Item(ID)
         
         txtOrderNo.Text = BD.ORDER_NO
         txtPayName.Text = BD.PAY_NAME
         txtAmount.Text = BD.AMOUNT
         txtPrice.Text = BD.PRICE
         txtChequeNo.Text = BD.CHEQUE_NO
         txtChequeName.Text = BD.CHEQUE_NAME
         
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
   End If
   Call QueryData(True)
   
   Call txtChequeNo.SetFocus
   
   Call ParentForm.RefreshGrid
   
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim I As Long
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim BD As CStockMarketDetail
   If ShowMode = SHOW_ADD Then
      Set BD = New CStockMarketDetail
      BD.Flag = "A"
      Call TempCollection.Add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.ORDER_NO = Val(txtOrderNo.Text)
   BD.PAY_NAME = txtPayName.Text
   BD.AMOUNT = Val(txtAmount.Text)
   BD.PRICE = Val(txtPrice.Text)
   BD.CHEQUE_NO = txtChequeNo.Text
   BD.CHEQUE_NAME = txtChequeName.Text
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      
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
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   Set m_Products = New Collection
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_StockMarketDetail = New CStockMarketDetail
   Set m_Locations = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_StockMarketDetail = Nothing
   Set m_Products = Nothing
   Set m_Locations = Nothing
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtChequeName_Change()
   m_HasModify = True
End Sub

Private Sub txtChequeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtOrderNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPayName_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub
