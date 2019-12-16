VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmDBSetting 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   2055
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3625
      _Version        =   131073
      PictureBackgroundStyle=   1
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtServerName 
         Height          =   435
         Left            =   1890
         TabIndex        =   0
         Top             =   720
         Width           =   3555
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin VB.Label lblServerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   5
         Top             =   810
         Width           =   1575
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   525
         Left            =   2820
         TabIndex        =   2
         Top             =   1320
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1200
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDBSetting.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmDBSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public Header As String

Public FileDb As String
Public UserName As String
Public Password As String
Public IP As String
Public Port As String

Private Sub cmdCancel_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If Len(txtServerName.Text) > 0 Then
      FileDb = Trim(txtServerName.Text) & ":" & "C:\GDB\DATABASE.GDB"
      glbErrorLog.LocalErrorMsg = "เชื่อมต่อไปยังเครื่อง " & Trim(txtServerName.Text) & " แล้ว"
   Else
      FileDb = "C:\GDB\DATABASE.GDB"
      glbErrorLog.LocalErrorMsg = "เชื่อมต่อไปยัง LOCALHOST แล้ว"
   End If
   
   UserName = "SYSDBA"
   Password = "masterkey"
   
   Call EnableForm(Me, False)
   
   Call glbDatabaseMngr.DisConnectDatabase
   
   If Not glbDatabaseMngr.ConnectDatabase(FileDb, UserName, Password, glbErrorLog) Then
'      glbErrorObj.LocalErrorMsg = "ไม่สามารถเชื่อมต่าดาตาเบสได้ กรุณาลองใหม่ "
'      glbErrorObj.ShowUserError
      
      Call EnableForm(Me, True)
      txtServerName.SetFocus
      
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
   
   glbErrorLog.ShowUserError
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = DUMMY_KEY Then
      Call cmdCancel_Click
      KeyCode = 0
   End If
End Sub


Private Sub Form_Load()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = "SETUP DATABASE"
   pnlHeader.Caption = Me.Caption
   
   OKClick = False
'   Call InitDialogHeader(lblHeader, Header)
   
   Call InitNormalLabel(lblServerName, "ชื่อ SERVER")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
   'D:\Database\GDB\CHEQUE-LOGO.GDB
End Sub
