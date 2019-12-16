VERSION 5.00
Begin VB.UserControl uctlTime 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   LockControls    =   -1  'True
   ScaleHeight     =   405
   ScaleWidth      =   1170
   Begin VB.TextBox txtMM 
      Height          =   375
      Left            =   630
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtHH 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   285
      Left            =   540
      TabIndex        =   2
      Top             =   60
      Width           =   45
   End
End
Attribute VB_Name = "uctlTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event HasChange()

Public Property Let Enable(E As Boolean)
   Call SetEnableDisableTextBox(txtHH, E)
   Call SetEnableDisableTextBox(txtMM, E)
End Property

Public Property Let HR(HH As Long)
   txtHH.Text = Format(HH, "00")
End Property

Public Property Get HR() As Long
   HR = Val(txtHH.Text)
End Property

Public Property Let MI(MM As Long)
   txtMM.Text = Format(MM, "00")
End Property

Public Property Get MI() As Long
   MI = Val(txtMM.Text)
End Property

Private Sub txtHH_Change()
   RaiseEvent HasChange
End Sub

Private Sub txtHH_GotFocus()
   Call SetSelect(txtHH)
End Sub

Private Sub txtMM_Change()
   RaiseEvent HasChange
End Sub

Private Sub txtMM_GotFocus()
   Call SetSelect(txtMM)
End Sub

Private Sub txtHH_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
      (KeyAscii <> Asc(vbBack)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtMM_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
      (KeyAscii <> Asc(vbBack)) Then
      KeyAscii = 0
   End If
End Sub

Public Function VerifyTime(AllowNull As Boolean) As Boolean
   If AllowNull Then
      If (txtHH.Text = "") Or (txtMM.Text = "") Then
         If (txtHH.Text = "") And (txtMM.Text = "") Then
            VerifyTime = True
         Else
            VerifyTime = False
         End If
      Else
         If (Val(txtHH.Text) >= 0) And (Val(txtHH.Text) <= 23) And (Val(txtMM.Text) >= 0) And (Val(txtMM.Text) <= 59) Then
            VerifyTime = True
         Else
            VerifyTime = False
         End If
      End If
   Else
      If (txtHH.Text = "") Or (txtMM.Text = "") Then
         VerifyTime = False
      Else
         If (Val(txtHH.Text) >= 0) And (Val(txtHH.Text) <= 23) And (Val(txtMM.Text) >= 0) And (Val(txtMM.Text) <= 59) Then
            VerifyTime = True
         Else
            VerifyTime = False
         End If
      End If
   End If
End Function

Public Sub SetFocus()
   txtHH.SetFocus
End Sub

Private Sub UserControl_Initialize()
   UserControl.BackColor = GLB_FORM_COLOR
   Call InitNormalLabel(Label1, ":")
   
   Call InitTextBox(txtHH, "")
   Call InitTextBox(txtMM, "")
   
   Call SetTextLenType(txtHH, TEXT_INTEGER, 2)
   Call SetTextLenType(txtMM, TEXT_INTEGER, 2)
End Sub
