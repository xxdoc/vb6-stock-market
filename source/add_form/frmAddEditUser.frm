VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditUser 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmAddEditUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6588
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUserGroup 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2370
         Width           =   2955
      End
      Begin StockMarket.uctlTextBox txtUserName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   8
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtUserDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1920
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   767
      End
      Begin StockMarket.uctlTextBox txtPassword 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1470
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblPasswordLimit 
         Caption         =   "Label1"
         Height          =   435
         Left            =   6480
         TabIndex        =   13
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1530
         Width           =   1575
      End
      Begin Threed.SSCheck chkEnable 
         Height          =   345
         Left            =   6420
         TabIndex        =   1
         Top             =   1050
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblUserGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   2490
         Width           =   1575
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblUserDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2010
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   6
         Top             =   2940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3585
         TabIndex        =   5
         Top             =   2940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditUser.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_User As CUserAccount

Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public KEY As String

Private m_UserGroup As CUserGroup
Private Sub cboUserGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_User.USER_NAME = KEY
      m_User.QueryFlag = 1
      If Not glbDaily.QueryUserAccount(m_User, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If m_Rs.RecordCount > 0 Then
      Call m_User.PopulateFromRS(1, m_Rs)
      
      txtUserName.Text = m_User.USER_NAME
      txtUserDesc.Text = m_User.USER_DESC
      txtPassword.Text = DecryptText(m_User.USER_PASSWORD)
      cboUserGroup.ListIndex = IDToListIndex(cboUserGroup, m_User.GROUP_ID)
      chkEnable.Value = FlagToCheck(m_User.USER_STATUS)
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
   
   If Not VerifyTextControl(lblUsername, txtUserName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUserGroup, cboUserGroup, False) Then
      Exit Function
   End If
   If Len(txtPassword.Text) < 4 Then
      glbErrorLog.LocalErrorMsg = "กรุณาใส่รหัสผ่านอย่างน้อย 4 ตัว"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_User.AddEditMode = ShowMode
   m_User.USER_NAME = txtUserName.Text
   m_User.USER_DESC = txtUserDesc.Text
   m_User.USER_STATUS = Check2Flag(chkEnable.Value)
   m_User.GROUP_ID = cboUserGroup.ItemData(Minus2Zero(cboUserGroup.ListIndex))
   m_User.USER_PASSWORD = EncryptText(txtPassword.Text)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditUserAccount(m_User, IsOK, True, glbErrorLog) Then
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
      
      Call LoadUserGroup(m_UserGroup, cboUserGroup)
      
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
   
   Call InitNormalLabel(lblUsername, MapText("ชื่อผู้ใช้"))
   Call InitNormalLabel(lblUserDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้"))
   Call InitNormalLabel(lblPassword, MapText("รหัสผ่าน"))
   Call InitNormalLabel(lblPasswordLimit, MapText("(รหัสผ่านอย่างน้อย 4 ตัว)"), RGB(255, 0, 0))
   
   Call InitCheckBox(chkEnable, "ใช้งานได้")
   
   Call txtUserName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtUserName.SetTextType(1)
   Call txtUserDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPassword.SetTextLenType(TEXT_STRING, glbSetting.PASSWORD_TYPE)
   txtPassword.PasswordChar = "*"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboUserGroup)
   
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
   
   Set m_User = New CUserAccount
   Set m_Rs = New ADODB.Recordset
   Set m_UserGroup = New CUserGroup
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Set m_UserGroup = Nothing
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtUserDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtUsername_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub
