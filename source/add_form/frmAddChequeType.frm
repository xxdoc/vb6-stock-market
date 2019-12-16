VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddChequeType 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAddChequeType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   5741
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboChequeType 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   795
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   11535
      End
      Begin VB.Label lblChequeType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3720
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5640
         TabIndex        =   1
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeType.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddChequeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HeaderText As String
Public ChequeType As String
Public AccNo As String
Public AccName As String
Public m_HasModify  As Boolean
Public OKClick  As Boolean
Private Sub cboChequeType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   SaveData = False
   
   If Not VerifyCombo(lblChequeType, cboChequeType, False) Then
      Exit Function
   End If
   
   SaveData = True
   ChequeType = cboChequeType.Text
End Function
Private Sub Form_Activate()
   Me.Refresh
   DoEvents
   Call LoadChequeConfig(cboChequeType)
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
   pnlHeader.Caption = "��������������� ( ��¡�úѭ�� " & AccNo & " ��� �����Ţ�ѭ�� " & AccName & " )"
   
   Call InitNormalLabel(lblChequeType, MapText("�����������"), RGB(255, 0, 0))
   Call InitNormalLabel(lblNote, MapText("��س��ʻ���������� �����͡�ҡ��¡�úѭ�� ��� �����Ţ�ѭ�� �ͧ����� EXPRESS"))
   
   Call InitCombo(cboChequeType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 17
   
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call InitFormLayout
End Sub
