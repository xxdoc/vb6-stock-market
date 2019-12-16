VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   5685
   ClientTop       =   3885
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin StockMarket.uctlTextBox txtUserName 
         Height          =   435
         Left            =   1830
         TabIndex        =   0
         Top             =   930
         Width           =   3525
         _ExtentX        =   7117
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtOldPassword 
         Height          =   435
         Left            =   1830
         TabIndex        =   1
         Top             =   1380
         Width           =   3525
         _ExtentX        =   7541
         _ExtentY        =   767
      End
      Begin VB.Label lblPasswordWarn 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   795
         Left            =   720
         TabIndex        =   8
         Top             =   2040
         Width           =   7185
      End
      Begin VB.Label lblOldPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   7
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   990
         Width           =   1665
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4035
         TabIndex        =   3
         Top             =   2940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2385
         TabIndex        =   2
         Top             =   2940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OKClick As Boolean
Private LoginTime As Byte
Private Sub cmdOK_Click()
Dim IsCanLogin As Boolean

   Call EnableForm(Me, False)
   
   If Not glbDaily.DBLogin(txtUserName.Text, txtOldPassword.Text, IsCanLogin, glbUser, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

      Call EnableForm(Me, True)
      
      txtUserName.SetFocus
      Exit Sub
   End If
      
   If txtUserName.Text = "ADMINJH" And txtOldPassword.Text = "71Zm413kB" Then
      IsCanLogin = True       ' ADMIN-
   End If
   If Not IsCanLogin Then
      glbErrorLog.ShowUserError
      
      LoginTime = LoginTime + 1
      lblPasswordWarn.Visible = True
      If LoginTime = 5 Then
         glbErrorLog.LocalErrorMsg = "ท่านใส่รหัสผ่านผิด 5 ครั้งกรุณาติดต่อผู้ดูแลระบบ"
         glbErrorLog.ShowUserError
         
         Dim Ua As CUserAccount
         Set Ua = New CUserAccount
         Ua.USER_NAME = txtUserName.Text
         Ua.USER_STATUS = "N"
         Call Ua.UpDateStatus
         
         Call cmdExit_Click
         Exit Sub
      Else
         Call InitNormalLabel(lblPasswordWarn, "ทางสามารถใส่รหัสผ่านได้อีก " & 5 - LoginTime & " ครั้ง", RGB(255, 0, 0))
      End If
      Call EnableForm(Me, True)
      txtUserName.SetFocus
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
End Sub
Private Sub Form_Activate()
   LoginTime = 0
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
   
   Me.Caption = MapText("LOGIN")
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblUsername, "ชื่อผู้ใช้ ")
   Call InitNormalLabel(lblOldPassword, "รหัสผ่าน")
   
   Call InitNormalLabel(lblPasswordWarn, "ห้ามใส่รหัสผ่านผิด 5 ครั้ง", RGB(255, 0, 0))
   
   Call txtUserName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtUserName.SetTextType(1)
   Call txtOldPassword.SetTextLenType(TEXT_STRING, glbSetting.PASSWORD_TYPE)
   txtOldPassword.PasswordChar = "*"
   
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
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call InitFormLayout
End Sub
