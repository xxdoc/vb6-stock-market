VERSION 5.00
Begin VB.UserControl uctlDate 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ScaleHeight     =   405
   ScaleWidth      =   2805
   Begin VB.TextBox txtYear 
      Height          =   405
      Left            =   2120
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1635
   End
   Begin VB.TextBox txtDay 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "uctlDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event HasChange()
Private m_ShowDate As Date
Private m_HH As String
Private m_MM As String
Private m_SS As String
Private m_SetTimeFlag As Boolean
Private DisableFlag As Boolean

Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
Dim LastDay As Byte
   LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", ValidDate) + 1, ValidDate))))
   LastDayOfMonth = LastDay
End Function
Public Property Let Enable(E As Boolean)
   UserControl.Enabled = E
   Call SetEnableDisableTextBox(txtDay, E)
   Call SetEnableDisableComboBox(cboMonth, E)
   Call SetEnableDisableTextBox(txtYear, E)
End Property

Public Property Let ShowFirstDate(D As Date)
'   m_ShowDate = D
   If D > 0 Then
      txtDay.Text = 1
      txtYear.Text = Format(Year(D) + 543, "0000")
      cboMonth.ListIndex = Month(D)
   Else
      txtDay.Text = ""
      If txtYear.Enabled Then
         txtYear.Text = ""
      Else
         txtYear.Text = "" 'Year(Now) + 543
      End If
      cboMonth.ListIndex = 0
   End If
End Property

Public Property Let ShowLastDate(D As Date)
'   m_ShowDate = D
   If D > 0 Then
      txtDay.Text = LastDayOfMonth(D)
      txtYear.Text = Format(Year(D) + 543, "0000")
      cboMonth.ListIndex = Month(D)
   Else
      txtDay.Text = ""
      If txtYear.Enabled Then
         txtYear.Text = ""
      Else
         txtYear.Text = "" 'Year(Now) + 543
      End If
      cboMonth.ListIndex = 0
   End If
End Property

Public Property Let ShowTime(HH As Long, MM As Long, SS As Long)
   m_SetTimeFlag = True
   
   m_HH = Format(HH, "00")
   m_MM = Format(MM, "00")
   m_SS = Format(SS, "00")
End Property

Public Property Let ShowDate(D As Date)
   m_ShowDate = D
   If D > 0 Then
      txtDay.Text = Day(D)
      txtYear.Text = Format(Year(D) + 543, "0000")
      cboMonth.ListIndex = Month(D)
   Else
      txtDay.Text = ""
      If txtYear.Enabled Then
         txtYear.Text = ""
      Else
         txtYear.Text = "" 'Year(Now) + 543
      End If
      cboMonth.ListIndex = 0
   End If
End Property

Public Property Get ShowDate() As Date
Dim InternalDate As String
   If (txtYear.Text = "") Or (txtDay.Text = "") Or (cboMonth.Text = "") Then
      m_ShowDate = -3
      ShowDate = m_ShowDate
   Else
      If txtYear.Enabled Then
         If m_SetTimeFlag Then
            InternalDate = Format(Val(txtYear) - 543, "0000") & "-" & Format(cboMonth.ListIndex, "00") & "-" & Format(txtDay, "00") & " " & m_HH & ":" & m_MM & ":" & m_SS
         Else
            InternalDate = Format(Val(txtYear) - 543, "0000") & "-" & Format(cboMonth.ListIndex, "00") & "-" & Format(txtDay, "00") & " " & "00:00:00"
         End If
      Else
         If m_SetTimeFlag Then
            InternalDate = Format(Year(Now), "0000") & "-" & Format(cboMonth.ListIndex, "00") & "-" & Format(txtDay, "00") & " " & m_HH & ":" & m_MM & ":" & m_SS
         Else
            InternalDate = Format(Year(Now), "0000") & "-" & Format(cboMonth.ListIndex, "00") & "-" & Format(txtDay, "00") & " " & "00:00:00"
         End If
      End If
      m_ShowDate = InternalDateToDate(InternalDate)
      ShowDate = m_ShowDate
   End If
End Property

Public Sub SetFocus()
   If (txtDay.Visible) And (txtDay.Enabled) Then
      txtDay.SetFocus
   End If
End Sub

Public Function VerifyDate(AllowNull As Boolean) As Boolean
Dim TempYear As String
Dim FirstDate As Date
Dim LastDate As Date
   
'   If txtYear.Enabled = False Then
'      TempYear = "2003"
'   Else
      TempYear = txtYear.Text
'   End If
   If AllowNull Then
      If (txtDay.Text = "") Or (TempYear = "") Or (cboMonth.Text = "") Then
         If (txtDay.Text = "") And (TempYear = "") And (cboMonth.Text = "") Then
            VerifyDate = True
         Else
            VerifyDate = False
         End If
      Else
         VerifyDate = True
      End If
   Else
      If (txtDay.Text = "") Or (TempYear = "") Or (cboMonth.Text = "") Then
         VerifyDate = False
      Else
         
         VerifyDate = True
      End If
   End If
   'เช็คจำนวนวันที่ใส่ถูกต้องหรือไม่
   
      
   If Val(txtDay.Text) > LastDayOfMonth(DateSerial(Val(TempYear) - 543, cboMonth.ListIndex, 1)) Then
      VerifyDate = False
   End If
End Function

Private Sub cboMonth_Change()
   RaiseEvent HasChange
End Sub

Private Sub cboMonth_DropDown()
   RaiseEvent HasChange
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
Static DigitCount As Integer
Static Buf(2) As String
Dim Temp As Long
   
   If DigitCount = 0 Then
      If (Chr(KeyAscii) = "0") Or (Chr(KeyAscii) = "1") Then
         Buf(1) = Chr(KeyAscii)
      Else
         Buf(1) = "0"
         Buf(2) = "0"
      End If
      DigitCount = 1
   ElseIf DigitCount = 1 Then
      If (Chr(KeyAscii) >= "0") And (Chr(KeyAscii) <= "9") Then
         Buf(2) = Chr(KeyAscii)
      End If
      DigitCount = 0
      
      Temp = Val(Buf(1)) * 10 + Val(Buf(2))
      If Temp <= 12 Then
         cboMonth.ListIndex = Temp
      End If
   Else
      DigitCount = 0
      Buf(1) = "0"
      Buf(2) = "0"
   End If
      
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtDay_Change()
   RaiseEvent HasChange
End Sub

Private Sub txtDay_GotFocus()
   Call SetSelect(txtDay)
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
      (KeyAscii <> Asc(vbBack)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtYear_Change()
   RaiseEvent HasChange
End Sub

Private Sub txtYear_GotFocus()
   Call SetSelect(txtYear)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
      (KeyAscii <> Asc(vbBack)) Then
      KeyAscii = 0
   End If
   RaiseEvent HasChange
End Sub

Private Sub UserControl_Initialize()
Dim I As Long

   m_SetTimeFlag = False
   m_HH = ""
   m_MM = ""
   m_SS = ""
   
   Call InitTextBox(txtDay, "")
   Call InitTextBox(txtYear, "")
   Call InitCombo(cboMonth)
   txtDay.MaxLength = 2
   txtYear.MaxLength = 4
   
   m_ShowDate = -3
   For I = 0 To 12
      cboMonth.AddItem (IntToThaiMonth(I))
   Next
End Sub

Public Function EnableDisableYear(Flag As Boolean) As Boolean
   Call SetEnableDisableTextBox(txtYear, Flag)
   If Len(txtYear.Text) <= 0 And (Not Flag) Then
      txtYear.Text = Year(Now) + 543
   End If
End Function


