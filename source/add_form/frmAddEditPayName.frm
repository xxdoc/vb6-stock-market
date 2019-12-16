VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPayName 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmAddEditPayName.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   5106
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin StockMarket.uctlTextBox txtPayName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   8205
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtPayNameChange 
         Height          =   435
         Left            =   1890
         TabIndex        =   1
         Top             =   1560
         Width           =   8205
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblPayNameChange 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label lblPayName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   3
         Top             =   2220
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3585
         TabIndex        =   2
         Top             =   2220
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPayName.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPayName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PayName As CPayName

Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public KEY As String
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
      
      m_PayName.PAY_NAME = KEY
      Call m_PayName.QueryData(m_Rs, ItemCount)
   End If
      
   If m_Rs.RecordCount > 0 Then
      Call m_PayName.PopulateFromRS(1, m_Rs)
      
      txtPayName.Text = m_PayName.PAY_NAME
      txtPayNameChange.Text = m_PayName.PAY_NAME_CHANGE
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblPayName, txtPayName, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   'เช็คว่ามีชื่อ หรือยังเพราะห้ามซ้ำ
   If Not CheckUniqueNsKey(PAY_NAME, txtPayName.Text) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPayName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   m_PayName.AddEditMode = ShowMode
   m_PayName.PAY_NAME = txtPayName.Text
   m_PayName.PAY_NAME_CHANGE = txtPayNameChange.Text
   
   Call EnableForm(Me, False)
   
   Call m_PayName.AddEditData
   
   Call glbPayeeName.Add(m_PayName, Trim(m_PayName.PAY_NAME))
   
   Call EnableForm(Me, True)
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblPayName, MapText("จ่ายให้"))
   Call InitNormalLabel(lblPayNameChange, MapText("ชื่อแทน"))
   
   Call txtPayName.SetTextLenType(TEXT_STRING, glbSetting.THAI_COMPANY_LONG_NAME_LEN)
   
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
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
   
   Set m_PayName = New CPayName
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_PayName = Nothing
End Sub

Private Sub txtPayName_Change()
   m_HasModify = True
End Sub

Private Sub txtPayNameChange_Change()
   m_HasModify = True
End Sub
