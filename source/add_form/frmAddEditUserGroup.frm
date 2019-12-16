VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddEditUserGroup 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAddEditUserGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TreeView trvMain 
         Height          =   5235
         Left            =   150
         TabIndex        =   14
         Top             =   2490
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9234
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin StockMarket.uctlTextBox txtGroupName 
         Height          =   435
         Left            =   2100
         TabIndex        =   0
         Top             =   990
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin StockMarket.uctlTextBox txtGroupDesc 
         Height          =   435
         Left            =   2100
         TabIndex        =   2
         Top             =   1440
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   767
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6000
         Top             =   7800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":27A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":307C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":3956
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":4230
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":438A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":4C64
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":553E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":5858
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":6132
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":6A0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":72E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditUserGroup.frx":7FC0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Threed.SSCheck chkEnable 
         Height          =   345
         Left            =   6660
         TabIndex        =   1
         Top             =   1020
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblGroupName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label lblGroupDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   12
         Top             =   1530
         Width           =   1995
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditUserGroup.frx":8C9A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   10
         Top             =   1650
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditUserGroup.frx":8FB4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditUserGroup.frx":92CE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   7
         Top             =   7830
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   4
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   3
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditUserGroup.frx":95E8
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditUserGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_UserGroup As CUserGroup

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub InitTreeView()
Dim Node As Node
   trvMain.Font.Size = 14
   trvMain.Font.Name = GLB_FONT_EX
   
   ClearTreeView (trvMain.hwnd)
    Call GenerateTree1(m_UserGroup.RightItems, Nothing, "ROOT", 0, 0)
End Sub
Private Sub GenerateTree1(RightItems As Collection, N As Node, NodeID As String, PID As Long, Level As Long)
Dim O As CGroupRight
Dim Node As Node
Dim NewNodeID As String
Dim L As Long

   For Each O In RightItems
      If O.PARENT_ID = PID Then
         If Level = 0 Then
            Set Node = trvMain.Nodes.Add(, tvwFirst, NodeID & O.RIGHT_ID, O.RIGHT_ITEM_DESC & " (" & O.RIGHT_ITEM_NAME & ")", 4, 4)
            Node.Tag = O.RIGHT_ID
            Node.Checked = FlagToCheck(O.RIGHT_STATUS)
            Call GenerateTree1(RightItems, Node, NodeID & O.RIGHT_ID, O.RIGHT_ID, Level + 1)
            Node.Expanded = False
         Else
            NewNodeID = NodeID & "-" & O.RIGHT_ID
            Set Node = trvMain.Nodes.Add(N, tvwChild, NewNodeID, O.RIGHT_ITEM_DESC & " (" & O.RIGHT_ITEM_NAME & ")", 3, 3)
            Node.Tag = O.RIGHT_ID
            Node.Checked = FlagToCheck(O.RIGHT_STATUS)
            Call GenerateTree1(RightItems, Node, NewNodeID, O.RIGHT_ID, Level + 1)
            Node.Expanded = False
         End If
      End If
   Next O
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
      
      m_UserGroup.GROUP_ID = ID
      m_UserGroup.QueryFlag = 1
      If Not glbDaily.QueryUserGroup(m_UserGroup, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If m_Rs.RecordCount > 0 Then
      Call m_UserGroup.PopulateFromRS(1, m_Rs)
      
      txtGroupName.Text = m_UserGroup.GROUP_NAME
      txtGroupDesc.Text = m_UserGroup.GROUP_DESC
      chkEnable.Value = FlagToCheck(m_UserGroup.GROUP_STATUS)
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
   
   If Not VerifyTextControl(lblGroupName, txtGroupName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblGroupDesc, txtGroupDesc, True) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_UserGroup.AddEditMode = ShowMode
   m_UserGroup.GROUP_ID = ID
   m_UserGroup.GROUP_NAME = txtGroupName.Text
   m_UserGroup.GROUP_DESC = txtGroupDesc.Text
   m_UserGroup.GROUP_STATUS = Check2Flag(chkEnable.Value)

   Call EnableForm(Me, False)
   If Not glbDaily.AddEditUserGroup(m_UserGroup, IsOK, True, glbErrorLog) Then
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
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         Call InitTreeView
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
         Call InitTreeView
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
   
   Call InitNormalLabel(lblGroupName, MapText("ชื่อกลุ่ม"))
   Call InitNormalLabel(lblGroupDesc, MapText("รายละเอียด"))
   
   Call InitCheckBox(chkEnable, "ใช้งานได้")
   
   Call txtGroupName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtGroupName.SetTextType(1)
   Call txtGroupDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
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
   
   Set m_UserGroup = New CUserGroup
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_UserGroup = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub UpdateChild(ByVal Node As MSComctlLib.Node, Flag As Boolean)
Dim N As Node

   If Node Is Nothing Then
      Exit Sub
   End If
   
   Node.Checked = Flag
   Set N = Node
   While Not (N Is Nothing)
      N.Checked = Flag
      Call UpdateChild(N.Child, Flag)
      Call UpdateCollection(N.Tag, Flag)
      Set N = N.Next
   Wend
End Sub

Private Sub UpdateCollection(ID As Long, Flag As Boolean)
Dim D As CGroupRight
   For Each D In m_UserGroup.RightItems
      If D.RIGHT_ID = ID Then
         If D.Flag <> "A" Then
            D.Flag = "E"
         End If
         
         If Flag Then
            D.RIGHT_STATUS = "Y"
         Else
            D.RIGHT_STATUS = "N"
         End If
      End If
   Next D
End Sub

Private Sub trvMain_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
   Call UpdateChild(Node.Child, Node.Checked)
   Call UpdateCollection(Node.Tag, Node.Checked)
End Sub

Private Sub txtGroupDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtGroupName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   trvMain.Width = ScaleWidth - 2 * trvMain.Left
   trvMain.Height = ScaleHeight - trvMain.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   
End Sub

