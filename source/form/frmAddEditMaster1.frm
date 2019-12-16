VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMaster1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMaster1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame Frame1 
      Height          =   1245
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   2196
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin StockMarket.uctlTextBox txtCode 
         Height          =   495
         Left            =   2520
         TabIndex        =   0
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
      End
      Begin StockMarket.uctlTextBox txtName 
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         TabIndex        =   9
         Top             =   690
         Width           =   2295
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2325
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   825
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1455
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   3360
         TabIndex        =   2
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5040
         TabIndex        =   3
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   13230
         TabIndex        =   7
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMaster1"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public MasterKey As String
Public MasterArea As MASTER_TYPE

Private m_MasterRef As CMasterRef

Public KEY_CODE As String
Public KEY_NAME As String

Public m_TempMr As CMasterRef
Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCode, "")
   Call InitNormalLabel(lblName, "")

   Call InitNormalLabel(lblCode, MapText(KEY_CODE))
   Call InitNormalLabel(lblName, MapText(KEY_NAME))
      
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)

   Call InitMainButton(cmdSave, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      m_MasterRef.KEY_ID = ID
      Call m_MasterRef.QueryData(1, m_Rs, ItemCount, True)
      If ItemCount > 0 Then
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         txtCode.Text = m_MasterRef.KEY_CODE
         txtName.Text = m_MasterRef.KEY_NAME
      End If
      Call EnableForm(Me, True)
   End If
   
   IsOK = True
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdSave_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
On Error GoTo ErrorHandler
Dim IsOK As Boolean

   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, Not txtName.Visible) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(MASTER_CODE, txtCode.Text, ID, Val(MasterArea)) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   
   m_MasterRef.ShowMode = ShowMode
   m_MasterRef.KEY_NAME = txtName.Text
   m_MasterRef.KEY_CODE = txtCode.Text
   m_MasterRef.MASTER_AREA = MasterArea
   
   Call glbDaily.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   
   Call EnableForm(Me, True)
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Call EnableForm(Me, True)
   SaveData = False
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
      Call cmdSave_Click
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
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_TempMr = New CMasterRef
   Set m_MasterRef = New CMasterRef
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_TempMr = Nothing
   Set m_MasterRef = Nothing
End Sub
Private Sub txtCode_Change()
   m_HasModify = True
End Sub
Private Sub txtName_Change()
   m_HasModify = True
End Sub
