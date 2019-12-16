VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDatabaseSelect 
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmDatabaseSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5700
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   10054
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   540
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4035
         Left            =   180
         TabIndex        =   0
         Top             =   870
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   7117
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2625
         TabIndex        =   1
         Top             =   5010
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
         Left            =   4275
         TabIndex        =   2
         Top             =   5010
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblMotherNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1620
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDatabaseSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DBPath As String
Public CompanyName As String
Public CompanyFullName  As String

Private FileName As String
Private m_Companys As Collection
Private Sub LoadTreeView(Col As Collection)
Dim N As Node
Dim Np As Node
Dim C As CSCComp
Dim I As Long

   I = 0
'Debug.Print "====="
   For Each C In Col
      I = I + 1
      Set N = TreeView1.Nodes.Add(, tvwFirst, C.PATH & "-" & I, C.COMPNAM & " (" & C.PATH & ")")
      N.Tag = C.PATH
'Debug.Print C.PATH
      N.Expanded = False
   Next C
'Debug.Print "====="
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim L As Long
   If TreeView1.SelectedItem Is Nothing Then
      Exit Function
   End If
   DBPath = TreeView1.SelectedItem.Tag
   If Len(DBPath) > 8 Then
      DBPath = Left(DBPath, 8)
   End If
   glbCompanyCode = DBPath
   L = InStr(1, glbParameterObj.DBFileExpress, "Secure", vbTextCompare)
   DBPath = Mid(glbParameterObj.DBFileExpress, 1, L - 1) & DBPath
   CompanyName = Mid(TreeView1.SelectedItem.Text, InStr(1, TreeView1.SelectedItem.Text, ".") + 1)
   glbCompanyName = CompanyName
   
'   CompanyFullName = Mid(TreeView1.SelectedItem.Text, 6)
'   CompanyFullName = Left(CompanyFullName, InStr(1, CompanyFullName, "(") - 1)
'   glbCompanyFullName = CompanyFullName
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
'      Call LoadDBPath(cboDBPath)

      Call LoadCompany(Nothing, m_Companys)
      Call LoadTreeView(m_Companys)
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Companys = Nothing
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   pnlHeader.Caption = "กรุณาเลือกบริษัทที่ต้องการออกรายงาน"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
'   Call InitCombo(cboDBPath)

   Call InitTreeView
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub InitTreeView()
   TreeView1.Font.Name = GLB_FONT
   TreeView1.Font.Size = 14
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Companys = New Collection
End Sub

Private Sub TreeView1_DblClick()
   Call cmdOK_Click
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
End Sub
