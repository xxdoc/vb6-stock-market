VERSION 5.00
Begin VB.UserControl uctlTextLookup 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LockControls    =   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   5385
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   3915
   End
   Begin StockMarket.uctlTextBox txtCode 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   767
   End
End
Attribute VB_Name = "uctlTextLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MyCombo As ComboBox
Public MyTextBox As uctlTextBox
Public MyCollection As Collection

Public Event Change()
Public TextEventRaise As Boolean

Private m_ClearText As Boolean

Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(S As Boolean)
   UserControl.Enabled = S
   txtCode.Enabled = S
   Call SetEnableDisableComboBox(cboName, S)
End Property

Private Sub cboName_Click()
Dim O As Object
Dim TempID As Long

   RaiseEvent Change
   
   If cboName.ListIndex <= 0 Then
      If m_ClearText Then
         txtCode.Text = ""
      End If
      Exit Sub
   End If
   
   TempID = cboName.ItemData(Minus2Zero(cboName.ListIndex))
   Set O = MyCollection.Item(Trim(Str(TempID)))
      
   txtCode.Text = O.KEY_LOOKUP
End Sub

Private Sub cboName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   ElseIf KeyAscii = 8 Then
      txtCode.SetFocus
   End If
End Sub

Private Sub txtCode_Change()
Static InUsed As Long
Static OldLen As Long
Dim p As Object
Dim Lenght As Long
Dim Found As Boolean

   If InUsed = 1 Then
      Exit Sub
   End If
   
   InUsed = 1
   
   Lenght = Len(txtCode.Text)
   If OldLen >= Lenght Then
      OldLen = Lenght
      
      InUsed = 0
      Exit Sub
   End If
   
   OldLen = Lenght
   Found = False
   For Each p In MyCollection
      If Mid(p.KEY_LOOKUP, 1, Lenght) = txtCode.Text Then
         txtCode.Text = p.KEY_LOOKUP
         Call txtCode.SetSelectText(Lenght, Len(p.KEY_LOOKUP) - Lenght)
         
         Found = True
         Exit For
      End If
   Next p
   
   If Found Then
      cboName.ListIndex = IDToListIndex(MyCombo, p.KEY_ID)
   Else
      m_ClearText = False
      cboName.ListIndex = -1
      m_ClearText = True
   End If
   
   InUsed = 0

   If TextEventRaise Then
      RaiseEvent Change
   End If
End Sub

Private Sub UserControl_Initialize()
   m_ClearText = True
   TextEventRaise = True
   
   Set MyCollection = New Collection
   Call txtCode.SetTextLenType(TEXT_STRING, 30)
   Call InitCombo(cboName)
   
   Set MyTextBox = txtCode
   Set MyCombo = cboName
End Sub

Private Sub UserControl_Terminate()
'   Set MyCollection = Nothing
End Sub
