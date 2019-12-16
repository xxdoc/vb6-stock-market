VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAboutUs 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "frmAboutUs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5741
      _Version        =   131073
      BackColor       =   16777215
      PictureBackgroundStyle=   2
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   2955
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   7455
      End
   End
End
Attribute VB_Name = "frmAboutUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
   Me.Refresh
   DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = DUMMY_KEY Then
      Unload Me
   End If
End Sub

Private Sub InitFormLayout()
Dim TempStr As String
   
   'SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = "เกี่ยวกับ XIVESS-CHEQUE MANAGEMENT"
   
   TempStr = "-----------------------------------------------------------------------"
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "XIVESS-CHEQUE MANAGEMENT V 1 " & VersionToString
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "-----------------------------------------------------------------------"
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "พัฒนาโดย XIVESS TEAM"
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "-----------------------------------------------------------------------"
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "Copyright 2010 All Right Reserve"
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "-----------------------------------------------------------------------"
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "CONNECTED TO " & GetServerName
   TempStr = TempStr & vbCrLf
   TempStr = TempStr & "-----------------------------------------------------------------------"
   
   Call InitNormalLabel(lblNote, TempStr)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   'Call A
End Sub
Private Sub Form_Load()
   Call InitFormLayout
End Sub
Private Sub A()
Dim I As Long
   For I = 1 To 255
      Debug.Print (Chr$(I))
   Next I
End Sub
