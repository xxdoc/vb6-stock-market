VERSION 5.00
Begin VB.UserControl uctlTextBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   LockControls    =   -1  'True
   ScaleHeight     =   585
   ScaleWidth      =   3105
   Begin VB.TextBox txtTextBox 
      Height          =   435
      Left            =   -30
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -30
      Width           =   3105
   End
End
Attribute VB_Name = "uctlTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_NullAllow As Boolean
Private m_PasswordChar As String
Private m_TextType As Long
Private KeySearch As String

Public Event Change()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Property Get IsNullAllow() As Boolean
   IsNullAllow = m_NullAllow
End Property
Public Sub SetSelectText(Start As Long, L As Long)
   txtTextBox.SelStart = Start
   txtTextBox.SelLength = L
End Sub
Public Property Let NullAllow(B As String)
   If B = "T" Then
      m_NullAllow = True
   Else
      m_NullAllow = False
   End If
End Property
Public Property Get PasswordChar() As String
   PasswordChar = txtTextBox.PasswordChar
End Property

Public Property Let PasswordChar(S As String)
   txtTextBox.PasswordChar = S
End Property

Public Property Get Text() As String
   Text = txtTextBox.Text
End Property

Public Property Let Text(S As String)
   txtTextBox.Text = S
End Property

Public Property Get Tag() As String
   Tag = UserControl.Tag
End Property

Public Property Let Tag(S As String)
   UserControl.Tag = S
End Property

Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(S As Boolean)
   UserControl.Enabled = S
   Call SetEnableDisableTextBox(txtTextBox, S)
End Property

Private Sub txtTextBox_Change()
   RaiseEvent Change
End Sub
Private Sub txtTextBox_GotFocus()
   Call SetSelect(txtTextBox)
End Sub
Private Sub txtTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 38 Or KeyCode = 40 Then
      If Len(KeySearch) > 0 Then
         frmTextBoxLookup.KeySearch = KeySearch
         frmTextBoxLookup.KEYWORD = txtTextBox.Text
         frmTextBoxLookup.HeaderText = "ค้นหา"
         Load frmTextBoxLookup
         frmTextBoxLookup.Show 1
         
         txtTextBox.Text = frmTextBoxLookup.KEYWORD
         
         Unload frmTextBoxLookup
         Set frmTextBoxLookup = Nothing
         SendKeys ("{TAB}")
      End If
   Else
      If KeyCode = 13 Then
         SendKeys ("{TAB}")
      End If
   End If
End Sub
Private Sub txtTextBox_KeyPress(KeyAscii As Integer)
   If m_TextType = 1 Then
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      RaiseEvent KeyPress(KeyAscii)
   End If
End Sub
Private Sub UserControl_Initialize()
   Call InitTextBox(txtTextBox, "")
   m_TextType = 0
End Sub
Public Sub SetFocus()
   If txtTextBox.Visible Then
      txtTextBox.SetFocus
   End If
End Sub
Public Sub SetTextLenType(TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      txtTextBox.Alignment = 1
   End If
   
   UserControl.Tag = TT
   txtTextBox.MaxLength = L
End Sub

Public Sub SetTextType(TextType As Long)
   m_TextType = TextType
End Sub
Public Sub SetKeySearch(Key As String)
    KeySearch = Key
End Sub
Public Sub Refresh()
   txtTextBox.Refresh
End Sub
Private Sub UserControl_Resize()
   txtTextBox.Width = UserControl.Width
   txtTextBox.Height = UserControl.Height
End Sub

