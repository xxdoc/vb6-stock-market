Attribute VB_Name = "modMain"
Option Explicit

Public Const ROOT_TREE = "Root"

Public Const DUMMY_KEY = 27
Public GLB_GRID_COLOR As Long
Public GLB_NORMAL_COLOR As Long
Public GLB_ALERT_COLOR As Long
Public GLB_SHOW_COLOR As Long
Public GLB_FORM_COLOR As Long
Public GLB_HEAD_COLOR As Long
Public GLB_GRIDHD_COLOR As Long
Public GLB_MANDATORY_COLOR As Long

Public Enum SHOW_MODE_TYPE
   SHOW_ADD = 1
   SHOW_EDIT = 2
   SHOW_VIEW = 3
   SHOW_VIEW_ONLY = 4
End Enum

Public Enum SHOW_BANK_TYPE
   KRUNGTHAI = 1
   KRUNGTHEP = 2
   GSIKORNTHAI = 3
   UOB = 4
   THAIPANIT = 5
   KRUNGSRI = 6
   TMB = 7
   NAKORN = 8
End Enum

Public Enum TEXT_BOX_TYPE
   TEXT_STRING = 1
   TEXT_INTEGER = 2
   TEXT_FLOAT = 3
   TEXT_FLOAT_MONEY = 4
   TEXT_INTEGER_MONEY = 5
End Enum

Public Enum MASTER_TYPE
   MASTER_BBANK = 1
End Enum

Public Enum UNIQUE_TYPE
   MASTER_CODE = 1
   MASTER_NAME = 2
   PAY_NAME = 3
   ENTERPRISE_CODE = 4
   ENTERPRISE_NAME = 5
   CONFIG_CHEQUE = 6
End Enum
'===================== For clear treeview =========================
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd _
    As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const TV_FIRST As Long = &H1100
Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Const TVGN_ROOT As Long = &H0
Const WM_SETREDRAW As Long = &HB
'===================== For clear treeview =========================

Public Const PROJECT_NAME = "CHEQUE MANAGEMENT"
Public Const GLB_FONT = "JasmineUPC"
Public Const GLB_FONT_EX = "Cordia New"
Private Const MODULE_NAME = "modMain"

Public glbErrorLog As clsErrorLog
Public glbDatabaseMngr As clsDatabaseMngr
Public glbSetting As clsGlobalSetting
Public glbParameterObj As clsParameter
Public glbDaily As clsDaily
Public glbUser As clsUser

Public glbAccessRight As Collection
Public glbPayeeName As Collection
Public glbCompanyName As String
Public glbCompanyCode As String
Public GenIDRs As ADODB.Recordset
Public glbGuiConfigs As CGuiConfigs
Public CounterTime As Long

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function VerifyDate(L As Label, D As uctlDate, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = "วันที่"
   Else
      S = L.Caption
   End If

   If Not D.VerifyDate(NullAllow) Then
      VerifyDate = False
      D.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyDate = True
   End If
End Function
Public Function VerifyDateToDay(L As Label, D As uctlDate) As Boolean
Dim S As String
Dim ServerDateTime As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   'CompareDate
   If D.ShowDate < InternalDateToDateEx(ServerDateTime) Then
      Call MsgBox("หมายเหตุ มีการใส่วันที่เช็คย้อนหลัง", vbOKOnly, PROJECT_NAME)
   End If
   
   'CompareDate
'   If Not D.VerifyDate(NullAllow) Then
'      VerifyDate = False
'      D.SetFocus
'      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
'   Else
'      VerifyDate = True
'   End If
End Function

Public Function VerifyTime(L As Label, T As uctlTime, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not T.VerifyTime(NullAllow) Then
      VerifyTime = False
      T.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyTime = True
   End If
End Function

Public Function VerifyTextData(L As Label, T As TextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextData = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextData = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextData = False
            Else
               VerifyTextData = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextData = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextData = False
               Else
                  VerifyTextData = True
               End If
            End If
            Exit Function
         End If
      End If
      
      VerifyTextData = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextData = False
         Exit Function
      End If
      
      VerifyTextData = True
   End If
End Function

Public Function VerifyTextControl(L As Label, T As uctlTextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextControl = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextControl = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextControl = True 'false
               Exit Function 'remove this if false
            Else
               VerifyTextControl = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextControl = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextControl = True 'false
                  Exit Function
               Else
                  VerifyTextControl = True
                  Exit Function
               End If
            End If
'            Exit Function
         End If
      End If
      
      VerifyTextControl = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextControl = False
         Exit Function
      End If
      
      VerifyTextControl = True
   End If
End Function

Public Function VerifyCombo(L As Label, C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyCombo = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled And C.Visible Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyCombo = True
End Function

Public Function VerifyComboEx(C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyComboEx = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyComboEx = True
End Function

Public Function VerifyItem(C As Collection, T As Object, Idx As Long) As Boolean
Dim I As Long
Dim Count As Long

   If C.Count <= 0 Then
      VerifyItem = True
      Exit Function
   End If
   
   For I = 1 To C.Count
      If C.Item(I).CURRENT_FLAG = "Y" Then
         Count = Count + 1
      End If
   Next I
   
   If Count <> 1 Then
      Call MsgBox("กรุณาเลือกข้อมูลให้มีค่าปัจจุบัน 1 รายการ", vbOKOnly, PROJECT_NAME)
   
      T.Tabs.Item(Idx).Selected = True
      VerifyItem = False
      Exit Function
   End If
   
   VerifyItem = True
End Function

Public Sub SetTextLenType(T As TextBox, TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      T.Alignment = 1
   End If
   
   T.Tag = TT
   T.MaxLength = L
End Sub

Public Function ChangeQuote(StrQ As String) As String
   ChangeQuote = Replace(StrQ, "'", "''")
End Function

Public Function NVLI(Value As Variant, I As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLI = I
   Else
      NVLI = Value
   End If
End Function
Public Function NVLB(Value As Variant, I As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLB = I
   Else
      NVLB = Value
   End If
End Function

Public Function NVLD(Value As Variant, I As Double) As Double
On Error Resume Next

   If IsNull(Value) Then
      NVLD = I
   Else
      NVLD = Value
   End If
End Function

Public Function NVLS(Value As Variant, S As String) As String
On Error Resume Next

   If IsNull(Value) Then
      NVLS = S
'   ElseIf IsEmpty(Value) Then
'      NVLS = S
   Else
      NVLS = Trim(Replace(Value, vbCrLf, ""))
   End If
End Function


Public Function EmptyToString(Value As String, S As String) As String
On Error Resume Next

   If Value = "" Then
      EmptyToString = S
   Else
      EmptyToString = Value
   End If
End Function

Public Function CryptString(strInput As String, strKey As String, action As Boolean)
Dim I As Integer, C As Integer
Dim strOutput As String

If Len(strKey) Then
    For I = 1 To Len(strInput)
        C = Asc(Mid$(strInput, I, 1))
        If action Then
            C = C + Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        Else: C = C - Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        End If
        strOutput = strOutput & Chr$(C And &HFF)
    Next I
Else
    strOutput = strInput
End If
CryptString = strOutput
End Function

Public Function EncryptText(PText As String) As String
   EncryptText = CryptString(PText, "GENETICOTHELLO", True)
End Function

Public Function DecryptText(CText As String) As String
   DecryptText = CryptString(CText, "GENETICOTHELLO", False)
End Function

Public Function EnableForm(Frm As Form, En As Boolean)
   If Frm Is Nothing Then
      Exit Function
   End If
   
   Frm.Enabled = En
   If En Then
      Screen.MousePointer = vbArrow
   Else
      Frm.Refresh
      DoEvents
      Screen.MousePointer = 11
   End If
End Function

Public Function IntToThaiMonth(M As Long) As String
   If glbParameterObj Is Nothing Then
      Exit Function
   End If
   
   If M = 1 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มกราคม"
      Else
         IntToThaiMonth = "January"
      End If
   ElseIf M = 2 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กุมภาพันธ์"
      Else
         IntToThaiMonth = "February"
      End If
      
   ElseIf M = 3 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มีนาคม"
      Else
         IntToThaiMonth = "March"
      End If
      
   ElseIf M = 4 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "เมษายน"
      Else
         IntToThaiMonth = "April"
      End If
      
   ElseIf M = 5 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "พฤษภาคม"
      Else
         IntToThaiMonth = "May"
      End If
      
   ElseIf M = 6 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มิถุนายน"
      Else
         IntToThaiMonth = "June"
      End If
      
   ElseIf M = 7 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กรกฎาคม"
      Else
         IntToThaiMonth = "July"
      End If
      
   ElseIf M = 8 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "สิงหาคม"
      Else
         IntToThaiMonth = "August"
      End If
      
   ElseIf M = 9 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กันยายน"
      Else
         IntToThaiMonth = "September"
      End If
      
   ElseIf M = 10 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "ตุลาคม"
      Else
         IntToThaiMonth = "October"
      End If
      
   ElseIf M = 11 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "พฤศจิกายน"
      Else
         IntToThaiMonth = "November"
      End If
      
   ElseIf M = 12 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "ธันวาคม"
      Else
         IntToThaiMonth = "December"
      End If
   Else
      IntToThaiMonth = ""
   End If
End Function

Public Function DateToStringMonthYearExt(D As Date) As String
   If D < 0 Then
      DateToStringMonthYearExt = ""
      Exit Function
   End If
   
   DateToStringMonthYearExt = " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
End Function

Public Function DateToStringExt(D As Date) As String
   If D < 0 Then
      DateToStringExt = "-"
      Exit Function
   Else
      DateToStringExt = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
   End If
End Function

Public Function DateToStringExtEx(D As Date) As String
   If D < 0 Then
      DateToStringExtEx = ""
      Exit Function
   End If
   
   DateToStringExtEx = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
End Function

Public Function DateToStringIntEx2(D As Date, Minute As Long, Second As Long) As String
   DateToStringIntEx2 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & " " & _
   Format(Minute, "00") & ":" & Format(Second, "00") & ":00"
End Function

Public Function DateToStringExtEx2(D As Date) As String
   If D > 0 Then
      DateToStringExtEx2 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      DateToStringExtEx2 = ""
   End If
End Function
Public Function DateToStringExtEx22(D As Date) As String
   If D > 0 Then
      DateToStringExtEx22 = Left(Format(Day(D), "00"), 1) & "     " & Right(Format(Day(D), "00"), 1) & "     " & Left(Format(Month(D), "00"), 1) & "     " & Right(Format(Month(D), "00"), 1) & "     " & Left(Format(Year(D) + 543, "0000"), 1) & "     " & Mid(Format(Year(D) + 543, "0000"), 2, 1) & "     " & Mid(Format(Year(D) + 543, "0000"), 3, 1) & "     " & Right(Format(Year(D) + 543, "0000"), 1)
   Else
      DateToStringExtEx22 = ""
   End If
End Function
Public Function DateToStringExtEx222(D As Date) As String
   If D > 0 Then
      DateToStringExtEx222 = Left(Format(Day(D), "00"), 1) & "     " & Right(Format(Day(D), "00"), 1) & "     " & Left(Format(Month(D), "00"), 1) & "     " & Right(Format(Month(D), "00"), 1) & "     " & Mid(Format(Year(D) + 543, "0000"), 3, 1) & "     " & Right(Format(Year(D) + 543, "0000"), 1)
   Else
      DateToStringExtEx222 = ""
   End If
End Function

Public Function DateToStringExtEx3(D As Date) As String
   If D > 0 Then
      DateToStringExtEx3 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
      DateToStringExtEx3 = DateToStringExtEx3 & " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   Else
      DateToStringExtEx3 = ""
   End If
End Function

Public Function DateToStringIntEx3(D As Date) As String
   DateToStringIntEx3 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00")
End Function

Public Function InternalDateToStringEx4(D As String) As String
Dim T As Date
   T = InternalDateToDate(D)
   If T > 0 Then
      InternalDateToStringEx4 = Format(Day(T), "00") & "/" & Format(Month(T), "00") & "/" & Format(Year(T) + 543, "0000")
   Else
      InternalDateToStringEx4 = ""
   End If
End Function

Public Function DateToStringInt(D As Date) As String
   If D = -1 Then
      DateToStringInt = "9999-99-99 99:99:99"
   ElseIf D = -2 Then
      DateToStringInt = "0000-00-00 00:00:00"
   Else
      DateToStringInt = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   End If
End Function
Public Function DateToStringIntEndMonth(D As Date) As String
   DateToStringIntEndMonth = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-31" & _
                     " 00:00:00"
End Function

Public Function DateToStringIntEx(D As Date) As String
   DateToStringIntEx = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
End Function

Public Function DateToStringIntHi(D As Date) As String
   If D > 0 Then
      DateToStringIntHi = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
   Else
      DateToStringIntHi = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   End If
End Function

Public Function DateToStringIntLow(D As Date) As String
   If D = -1 Then
      DateToStringIntLow = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   ElseIf D = -2 Then
      DateToStringIntLow = "0000" & "-" & "00" & "-" & "00" & _
                     " 00:00:00"
   Else
      DateToStringIntLow = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                        " 00:00:00"
   End If
End Function
Public Function InternalDateToDate(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDate = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDate = -2
      Exit Function
   End If
   
   If Len(IntDate) < 19 Then
      InternalDateToDate = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
   
   HHStr = Mid(IntDate, 12, 2)
   MMStr = Mid(IntDate, 15, 2)
   SSStr = Mid(IntDate, 18, 2)
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDate = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If Len(IntDate) < 19 Then
      InternalDateToDateEx = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
      
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx2(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If Len(IntDate) < 10 Then
      InternalDateToDateEx2 = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
      
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
      
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx2 = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function
Public Function InternalDateToDateExpress(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "00000000") Then
      InternalDateToDateExpress = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "99999999") Then
      InternalDateToDateExpress = -2
      Exit Function
   End If
   
   If Len(IntDate) < 8 Then
      InternalDateToDateExpress = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 5, 2)
   DStr = Mid(IntDate, 7, 2)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateExpress = DateSerial(Y, M, D)
End Function

Public Function ReFormatDate(DStr As String) As String
Dim YYYY As String
Dim MM As String
Dim DD As String

   YYYY = Mid(DStr, 5, 4)
   MM = Mid(DStr, 3, 2)
   DD = Mid(DStr, 1, 2)
   
   ReFormatDate = YYYY & MM & DD
End Function

Public Sub InitTextBox(T As TextBox, Msg As String, Optional Password As String = "")
   T.PasswordChar = Password
   T.FontSize = 12
   T.FontName = "MS Sans Serif"
   T.Text = Msg
   T.BackColor = GLB_GRID_COLOR
   'T.FontBold = True
End Sub
Public Sub ClearTreeView(ByVal tvHwnd As Long)
Dim lNodeHandle As Long

    'Turn off redrawing on the Treeview for more speed improvements
    SendMessageLong tvHwnd, WM_SETREDRAW, False, 0

    Do
        lNodeHandle = SendMessageLong(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
         If lNodeHandle > 0 Then
            SendMessageLong tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
         Else
            Exit Do
         End If
    Loop

    'Turn on redrawing on the Treeview
    SendMessageLong tvHwnd, WM_SETREDRAW, True, 0
End Sub

Public Sub InitCombo(C As ComboBox)
   C.FontSize = 12
   C.FontName = "MS Sans Serif"
   C.BackColor = GLB_GRID_COLOR
End Sub

Public Function VerifyGrid(S As String) As Boolean
   If S = "" Then
      VerifyGrid = False
      glbErrorLog.LocalErrorMsg = "กรุณาเลือกข้อมูลที่ต้องการก่อน"
      glbErrorLog.ShowUserError
   Else
      VerifyGrid = True
   End If
End Function

Public Function ConfirmDelete(S As String) As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการจะลบข้อมูล " & S & "' ใช่หรือไม่"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmDelete = False
      Exit Function
   Else
      ConfirmDelete = True
   End If
End Function

Public Sub InitFormHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 20
   L.FontName = GLB_FONT
   L.Alignment = 2
   L.ForeColor = RGB(0, 10, 0)
End Sub

Public Sub InitDialogHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 16
   L.FontName = GLB_FONT
   L.Alignment = 2
End Sub

Public Sub InitNormalLabel(L As Label, Caption As String, Optional Color As Long = 0)
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub

Public Sub InitOption(O As OptionButton, Caption As String)
   O.Caption = Caption
   O.FontSize = 14
   O.FontBold = True
   O.FontName = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
End Sub

Public Sub InitOptionEx(O As SSOption, Caption As String)
   O.Caption = Caption
   O.Font.Size = 14
   O.Font.Bold = True
   O.Font.Name = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
   O.BackStyle = ssTransparent
End Sub

Public Sub InitCheckBox(C As SSCheck, Caption As String)
   C.Caption = Caption
   C.FontSize = 14
   C.FontBold = True
   C.FontName = GLB_FONT
   C.BackColor = GLB_FORM_COLOR
   C.BackStyle = ssTransparent
   C.TripleState = True
End Sub

Public Sub InitMainButton(B As SSCommand, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.Name = GLB_FONT
   B.Font3D = ssInsetLight
   B.BackColor = RGB(255, 255, 255)
   B.ButtonStyle = ssWin95 '= ssActiveBorders
   B.MousePointer = ssCustom
   B.MouseIcon = LoadPicture(glbParameterObj.ButtonCursor)
End Sub
Public Sub SetSelect(T As TextBox)
   T.SelStart = 0
   T.SelLength = Len(T.Text)
End Sub
Public Sub InitDialogButton(B As CommandButton, Caption As String)
   B.Caption = Caption
   B.FontBold = True
   B.FontSize = 14
   B.FontName = GLB_FONT
   
   B.BackColor = &HFFFFFF
End Sub
Public Sub ReleaseAll()
   Set glbErrorLog = Nothing
   Set glbDatabaseMngr = Nothing
   Set glbParameterObj = Nothing
   Set glbUser = Nothing
   Set glbDaily = Nothing
   Set glbPayeeName = Nothing
End Sub
Public Sub SetEnableDisableTextBox(T As TextBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableComboBox(T As ComboBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableButton(B As SSCommand, En As Boolean)
   If En Then
      B.Enabled = True
      B.BackColor = GLB_GRID_COLOR
   Else
      B.Enabled = False
      B.BackColor = &H8000000F
   End If
End Sub

Public Function ConfirmExit(HasEdit As Boolean) As Boolean
   If Not HasEdit Then
      ConfirmExit = True
   Else
      glbErrorLog.LocalErrorMsg = "ท่านต้องการจะออกจากโปรแกรมโดยไม่มีการบันทึกข้อมูลใช่หรือไม่"
      If glbErrorLog.AskMessage = vbYes Then
         ConfirmExit = True
      Else
         ConfirmExit = False
      End If
   End If
End Function
Public Function ThaiBaht(ByVal pamt As Double) As String
Dim valstr As String, vLen As Integer, vno As Integer, syslge As String
Dim I As Integer, j As Integer, v As Integer
Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
Dim vword(20) As String
   
 If pamt <= 0# Then
   ThaiBaht = ""
   Exit Function
 End If
 valstr = Trim(Format$(pamt, "##########0.00"))
 vLen = Len(valstr) - 3
 For I = 1 To 20
     vword(I) = ""
 Next I
wnumber(1) = "หนึ่ง": wnumber(2) = "สอง": wnumber(3) = "สาม": wnumber(4) = "สี่"
wnumber(5) = "ห้า": wnumber(6) = "หก": wnumber(7) = "เจ็ด": wnumber(8) = "แปด"
wnumber(9) = "เก้า": wdigit(1) = "บาท": wdigit(2) = "สิบ": wdigit(3) = "ร้อย": wdigit(4) = "พัน"
wdigit(5) = "หมื่น": wdigit(6) = "แสน": wdigit(7) = "ล้าน": spcdg(1) = "สตางค์": spcdg(2) = "เอ็ด"
spcdg(3) = "ยี่": spcdg(4) = "ถ้วน"
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
        vword(I) = ""
        If (vLen - I + 1) = 7 Then
            vword(I) = wdigit(7)             '--ล้าน
        End If
    Else
        If (vLen - I + 1) > 7 Then
            j = vLen - I - 5               '--เกินหลักล้าน
        Else
            j = vLen - I + 1               '--หลักแสน
        End If
        vword(I) = wnumber(vno) + wdigit(j) '-30ถึง90
        If vno = 1 And j = 2 Then
            vword(I) = wdigit(2)             '--สิบ
        End If
        If vno = 2 And j = 2 Then
            vword(I) = spcdg(3) + wdigit(j)  '--ยี่สิบ
        End If
        If j = 1 Then                       ' สิยเอ็ค -->เก้าสิบเอ็ด
            vword(I) = wnumber(vno)
            If vno = 1 And vLen > 1 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2)
                End If
            End If
        End If
        If j = 7 Then         '-แก้บักกรณี 11,111,111.00 สิบเอ็ด
            vword(I) = wnumber(vno) + wdigit(j)   '-ล้าน
            If vno = 1 And vLen > 7 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2) + wdigit(j)
                End If
            End If
        End If
    End If
Next I
    
If Int(pamt) > 0 Then
       vword(vLen) = vword(vLen) + wdigit(1)
End If
 '--------------ทศนิยม --------------
valstr = Mid$(valstr, vLen + 2, 2)
vLen = Len(valstr)
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
           vword(I + 10) = ""
    Else
           j = vLen - I + 1
           vword(I + 10) = wnumber(vno) + wdigit(j)
        If vno = 1 And j = 2 Then
              vword(I + 10) = wdigit(2)
        End If
        If vno = 2 And j = 2 Then
              vword(I + 10) = spcdg(3) + wdigit(j)
        End If
        If j = 1 Then
            If vno = 1 And Int(Val(Mid$(valstr, I - 1, 1))) <> 0 Then
                 vword(I + 10) = spcdg(2)
            Else
                 vword(I + 10) = wnumber(vno)
            End If
        End If
    End If
Next I
If pamt <> 0 Then
    If Val(valstr) = 0 Then
        vword(13) = spcdg(4)
    Else
        vword(13) = spcdg(1)
    End If
End If

 '*** เผื่อใช้กรณียาวมาก และต้องการตัดประโยค
 valstr = ""
 For I = 1 To 20
    'IF LEN(valstr) < 70 AND LEN(valstr + vword(i)) > 70 Then
    '   valstr = valstr + REPLICATE(" ",70 - LEN(valstr))
    'END IF
    valstr = valstr + vword(I)
 Next I
 'valstr='('+valstr+')'
 ThaiBaht = (valstr)
End Function

Public Function WildCard(WStr As String, SubLen As Long, NewStr As String) As Boolean
Dim Tmp As String
   Tmp = Trim(WStr)
   If Tmp = "" Then
      WildCard = False
      Exit Function
   End If
   
   If Mid(Tmp, Len(Tmp)) = "%" Then
      SubLen = Len(Tmp) - 1
      NewStr = Mid(Tmp, 1, SubLen)
      
      WildCard = True
   Else
      WildCard = False
   End If
End Function

Public Function FormatString(S As String, Patch As String, L As Long) As String
Dim Temp As String
Dim Start As Long
Dim I As Long
Dim j As Long

   Temp = Space(L)
   Call Replace(Temp, " ", Patch)
   j = 0
   Start = (L - Len(S)) \ 2
   
   For I = 1 To L
      If I < Start Then
         Mid(Temp, I) = Patch
      Else
         If I > Start + Len(S) Then
            Mid(Temp, I) = Patch
         Else
            j = j + 1
            Mid(Temp, I) = Mid(S, j)
         End If
      End If
   Next I
   
   FormatString = Temp
End Function

Public Function FormatNumber(N As Variant, Optional ZeroString As String = "0.00") As String
Dim T As Double

   If IsNull(N) Then
      T = 0
   Else
      T = Val(N)
   End If
   
   If T = 0 Then
      FormatNumber = ZeroString
   ElseIf T > 0 Then
      FormatNumber = Format(T, "#,##0.00")
   ElseIf T < 0 Then
      FormatNumber = "(" & Format(-1 * T, "#,##0.00") & ")"
   End If
End Function

Public Function FormatNumberInt(N As Variant, Optional ZeroString As String = "0") As String
Dim T As Double

   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      FormatNumberInt = ZeroString
   ElseIf T > 0 Then
      FormatNumberInt = Format(T, "#,##0")
   ElseIf T < 0 Then
      FormatNumberInt = "(" & Format(-1 * T, "#,##0") & ")"
   End If
End Function
Public Function FormatNumberToNull(N As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "") As String
Dim T As Double
Dim TempStr As String
Dim I As Long

   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumberToNull = ZeroString & TempStr
      Else
         FormatNumberToNull = ZeroString
      End If
   ElseIf Quat Then
      FormatNumberToNull = Format(T, "#,##0" & TempStr)
   Else
      FormatNumberToNull = Format(T, "0" & TempStr)
   End If
End Function

Public Function ReverseFormatNumber(N As String) As Double
   ReverseFormatNumber = Val(Replace(N, ",", ""))
End Function

Public Function IDToListIndex(Cbo As ComboBox, ID As Long) As Long
Dim I As Long
Dim Temp As String

   IDToListIndex = -1
   For I = 0 To Cbo.ListCount - 1
      If InStr(Cbo.ItemData(I), ":") <= 0 Then
         Temp = Cbo.ItemData(I)
      Else
         Temp = Mid(Cbo.ItemData(I), 1, InStr(Cbo.ItemData(I), ":") - 1)
      End If
      If Temp = ID Then
         IDToListIndex = I
      End If
   Next I
End Function

Public Sub Main()
Dim I As Long
Dim AmountKey As Integer
   
   GLB_GRID_COLOR = RGB(255, 255, 250)
   GLB_NORMAL_COLOR = RGB(0, 0, 0)
   GLB_ALERT_COLOR = RGB(255, 0, 0)
   GLB_FORM_COLOR = RGB(180, 200, 200)
   GLB_HEAD_COLOR = GLB_FORM_COLOR
   GLB_GRIDHD_COLOR = RGB(149, 194, 240)
   GLB_SHOW_COLOR = RGB(0, 0, 240)
   GLB_MANDATORY_COLOR = RGB(0, 0, 255)

   Set glbSetting = New clsGlobalSetting
   Set glbParameterObj = New clsParameter
   Set glbUser = New clsUser
   Set glbErrorLog = New clsErrorLog
   Set glbAccessRight = New Collection
   glbErrorLog.DayKeepLog = 10
   glbErrorLog.LogFileMode = LOG_CURRENT_DATE
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Main"
   glbErrorLog.MsgBoxTitle = PROJECT_NAME
   
   If App.PrevInstance = True Then
      glbErrorLog.LocalErrorMsg = "โปรแกรมเดิมได้ถูกรันก่อนหน้านี้แล้ว"
      glbErrorLog.ShowUserError

      Set glbErrorLog = Nothing
      Exit Sub
   End If
   
   Load frmSplash
   frmSplash.Show 0
   frmSplash.Refresh
   
   Set glbDatabaseMngr = New clsDatabaseMngr
   
   If Not glbDatabaseMngr.ConnectDatabase(glbParameterObj.DBFile, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
      frmDBSetting.UserName = glbParameterObj.UserName
      frmDBSetting.Password = glbParameterObj.Password
      frmDBSetting.FileDb = glbParameterObj.DBFile
      frmDBSetting.Header = " ไม่สามารถเชื่อต่อฐานข้อมูลได้ "
      
      Load frmDBSetting
      frmDBSetting.Show 1
      If frmDBSetting.OKClick Then
         glbParameterObj.UserName = frmDBSetting.UserName
         glbParameterObj.Password = frmDBSetting.Password
         glbParameterObj.DBFile = frmDBSetting.FileDb
         
         glbErrorLog.LocalErrorMsg = "ระบบจะทำการบันทึกค่าที่ต้องใช้ในระบบ กรุณาเปิดโปรแกรมใหม่เพื่อใช้ค่าที่ตั้งขึ้น"
         glbErrorLog.ShowUserError
         
         Unload frmDBSetting
         Set frmDBSetting = Nothing

         Unload frmSplash
         Set frmSplash = Nothing
         
         Call ReleaseAll
         Exit Sub
      Else
         Unload frmDBSetting
         Set frmDBSetting = Nothing

         Unload frmSplash
         Set frmSplash = Nothing

         Call ReleaseAll
         End
      End If
      Unload frmDBSetting
      Set frmDBSetting = Nothing
   End If
   
   glbParameterObj.ComputerName = glbDatabaseMngr.GetComputerName
   
   Unload frmSplash
   Set frmSplash = Nothing
   
   Set glbDaily = New clsDaily
   Set glbPayeeName = New Collection
   
   Set glbGuiConfigs = New CGuiConfigs
   Call glbGuiConfigs.CreateGuiConfig("")
   
   Load frmMain
   frmMain.Show
   
End Sub

Public Sub InitOrderType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("น้อยไปมาก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("มากไปน้อย"))
   C.ItemData(2) = 2
End Sub

Public Function GetItem(Col As Collection, Idx As Long, RealIndex As Long) As Object
Dim I As Long
Dim Count As Long

   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
      If Count = Idx Then
         RealIndex = I
         Set GetItem = Col.Item(I)
         Exit Function
      End If
   Next I
   
   Set GetItem = Nothing
End Function

Public Function CountItem(Col As Collection) As Long
Dim I As Long
Dim Count As Long

   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
   Next I
   
   CountItem = Count
End Function

Public Function VSP_CalTable(ByVal pRaw As String, ByVal pWidth As Long, ByRef pPer() As Long) As String
On Error GoTo ErrorHandler
Dim strTemp As String
Dim I As Long
Dim Count As Long
Dim iPer As Long
Dim tPer As Long
Dim TOTAL As Long
Dim Prefix() As String
Dim Value() As Long
Dim iTemp As Long
   
   pRaw = Trim$(pRaw)
   If Len(pRaw) <= 0 Then
      VSP_CalTable = ""
      Exit Function
   End If
   Count = 0
   iPer = 1
   TOTAL = 0
   strTemp = ""
   While iPer <= Len(pRaw)
      If Val(Mid$(pRaw, iPer, 1)) <= 0 Then
         strTemp = strTemp & Mid$(pRaw, iPer, 1)
      Else
         Count = Count + 1
         ReDim Preserve Prefix(Count)
         ReDim Preserve Value(Count)
         Prefix(Count) = strTemp
         tPer = InStr(iPer, pRaw, "|")
         If tPer <= 0 Then tPer = InStr(iPer, pRaw, ";")

         Value(Count) = Val(Mid$(pRaw, iPer, tPer - iPer))
         TOTAL = TOTAL + Value(Count)
         iPer = tPer
         strTemp = ""
      End If
      iPer = iPer + 1
   Wend
   strTemp = ""
   ReDim pPer(Count)
   For I = 1 To Count - 1
      iTemp = CLng((Value(I) * pWidth) / TOTAL)
      strTemp = strTemp & Trim$(Prefix(I)) & Trim$(Str$(iTemp)) & "|"
      If I = 1 Then
         pPer(I - 1) = iTemp
      Else
         pPer(I - 1) = pPer(I - 2) + iTemp
      End If
   Next I
   strTemp = strTemp & Trim$(Prefix(I)) & CLng(((Value(I) * pWidth) / TOTAL)) & ";"
   If I > 1 Then
      iTemp = CLng((Value(I) * pWidth) / TOTAL)
      pPer(I - 1) = pPer(I - 2) + iTemp
   End If
   VSP_CalTable = strTemp

   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function

Public Function Check2Flag(A As Long) As String
   If A = ssCBChecked Then
      Check2Flag = "Y"
   Else
      Check2Flag = "N"
   End If
End Function

Public Function Check2FlagConvert(A As Long) As String
   If A = 1 Then
      Check2FlagConvert = "N"
   Else
      Check2FlagConvert = "Y"
   End If
End Function

Public Function FlagToCheck(F As String) As Long
   If F = "Y" Then
      FlagToCheck = 1
   Else
      FlagToCheck = 0
   End If
End Function

Public Function Minus2Zero(A As Double) As Long
   If A < 0 Then
      Minus2Zero = 0
   Else
      Minus2Zero = A
   End If
End Function

Public Function Zero2One(A As Double) As Long
   If A = 0 Then
      Zero2One = 1
   Else
      Zero2One = A
   End If
End Function

Public Function Minus2Flag(A As Double) As String
   If A < 0 Then
      Minus2Flag = "Y"
   Else
      Minus2Flag = "N"
   End If
End Function

Public Function AdjustPage(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional TestFlag As Boolean = False, Optional SpaceCount As Long) As Boolean
Dim TempStr As String

   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   If (Vsp.Y1 + Offset - SpaceCount) > (Vsp.PageHeight - Vsp.MarginBottom) Then
      If Not TestFlag Then
         Vsp.NewPage
      End If
      AdjustPage = True
   Else
      AdjustPage = False
   End If
End Function

Public Function PatchTable(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional EnableFlag As Boolean = True, Optional SpaceCount As Long = 0) As Boolean
Dim TempStr As String
   
   If Not EnableFlag Then
      PatchTable = True
      Exit Function
   End If
   
   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   While Not AdjustPage(Vsp, Header, Body, Offset, True, SpaceCount)
      Call Vsp.AddTable(Header, "", Body)
   Wend
End Function
Public Function MyDiffEx(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiffEx = 0
   Else
      MyDiffEx = D1 / D2
   End If
End Function

Public Function MyDiff(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiff = 0
   Else
      MyDiff = CDbl(Format(D1 / D2, "0.00"))
   End If
End Function
Public Function MapText(Msg As String) As String
   MapText = Msg
End Function

Public Function SetReportConfig(Vsp As VSPrinter, ReportClassName As String) As Boolean
Dim I As Long
Dim Count As Long
Dim Rp As CReportConfig
Dim TempRs As ADODB.Recordset
Dim Rps As Collection
Dim iCount As Long

   If Rps Is Nothing Then
      Set TempRs = New ADODB.Recordset
      
      Set Rps = New Collection
      Set Rp = New CReportConfig
      
      Rp.REPORT_CONFIG_ID = -1
      Call Rp.QueryData(TempRs, iCount)
      Set Rp = Nothing
      
      While Not TempRs.EOF
         Set Rp = New CReportConfig
         
         Call Rp.PopulateFromRS(1, TempRs)
         Call Rps.Add(Rp)
         
         Set Rp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Rp = Nothing
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   SetReportConfig = False
   For Each Rp In Rps
      If (Rp.REPORT_KEY = ReportClassName) And (LCase(Rp.COMPUTER_NAME) = LCase(glbParameterObj.ComputerName)) Then
         Vsp.PaperSize = Rp.PAPER_SIZE
         Vsp.ORIENTATION = Rp.ORIENTATION
         Vsp.MarginBottom = Rp.MARGIN_BOTTOM * 567
         Vsp.MarginFooter = Rp.MARGIN_FOOTER * 567
         Vsp.MarginHeader = Rp.MARGIN_HEADER * 567
         Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
         Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
         Vsp.MarginTop = Rp.MARGIN_TOP * 567
'         Vsp.FontName = Rp.FONT_NAME
'         Vsp.FontSize = Rp.FONT_SIZE
         Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
         Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
         If Rp.PAPER_HEIGHT > 0 Then
            Vsp.PaperWidth = Rp.PAPER_HEIGHT * 567
         End If
         If Rp.PAPER_WIDTH > 0 Then
            Vsp.PaperHeight = Rp.PAPER_HEIGHT * 567
         End If
         
         SetReportConfig = True
         Exit Function
      End If
   Next Rp
   Set Rps = Nothing
End Function


Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
Dim LastDay As Byte
   LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", ValidDate) + 1, ValidDate))))
   LastDayOfMonth = LastDay
End Function

Public Sub GetFirstLastDate(D As Date, FD As Date, Ld As Date)
Dim MM As Long
Dim DD1 As Long
Dim DD2 As Long
Dim YYYY As Long

   MM = Month(D)
   DD1 = 1
   DD2 = LastDayOfMonth(D)
   YYYY = Year(D)
   
   FD = DateSerial(YYYY, MM, DD1)
   Ld = DateSerial(YYYY, MM, DD2)
End Sub

Public Sub StartExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
   Vsp.ExportFile = glbParameterObj.ReportFile
   Vsp.ExportFormat = vpxPlainHTML
End Sub

Public Sub CloseExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
End Sub

Public Sub InitOrientation(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2Orientation(orLandscape))
   C.ItemData(1) = orLandscape

   C.AddItem (ID2Orientation(orPortrait))
   C.ItemData(2) = orPortrait
End Sub

Public Sub InitPaperSize(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2PaperSize(pprA4))
   C.ItemData(1) = pprA4

   C.AddItem (ID2PaperSize(pprLetter))
   C.ItemData(2) = pprLetter

   C.AddItem (ID2PaperSize(pprFanfoldUS))
   C.ItemData(3) = pprFanfoldUS
End Sub

Public Sub InitFontName(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("AngsanaUPC")
   C.ItemData(1) = 1
End Sub

Public Function ID2Orientation(TempID As OrientationSettings) As String
   If TempID = orLandscape Then
      ID2Orientation = "แนวนอน"
   Else
      ID2Orientation = "แนวตั้ง"
   End If
End Function

Public Function ID2PaperSize(TempID As PaperSizeSettings) As String
   If TempID = pprA4 Then
      ID2PaperSize = "A4"
   ElseIf TempID = pprLetter Then
      ID2PaperSize = "Letter"
   ElseIf TempID = pprFanfoldUS Then
      ID2PaperSize = "Us standard"
   Else
      ID2PaperSize = "A4"
   End If
End Function

Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function
Public Function GetObject(ClassName As String, m_TempCol As Collection, TempKey As String, Optional SetNew As Boolean = True) As Object
On Error Resume Next
Dim Ei As Object
   
   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If SetNew Then
         Set GetObject = GetNewClass(ClassName)
      End If
   Else
      Set GetObject = Ei
   End If
End Function
Public Function GetNewClass(ClassName As String) As Object
   If ClassName = "CStockMarket" Then
      Static m_CStockMarket As CStockMarket
      If m_CStockMarket Is Nothing Then
         Set m_CStockMarket = New CStockMarket
      End If
      Set GetNewClass = m_CStockMarket
   End If
End Function
Public Sub InitThaiMonth(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มกราคม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("กุมภาพันธ์"))
   C.ItemData(2) = 2

   C.AddItem (MapText("มีนาคม"))
   C.ItemData(3) = 3

   C.AddItem (MapText("เมษายน"))
   C.ItemData(4) = 4

   C.AddItem (MapText("พฤษภาคม"))
   C.ItemData(5) = 5
   
   C.AddItem (MapText("มิถุนายน"))
   C.ItemData(6) = 6

   C.AddItem (MapText("กรกฎาคม"))
   C.ItemData(7) = 7
   
   C.AddItem (MapText("สิงหาคม"))
   C.ItemData(8) = 8
   
   C.AddItem (MapText("กันยายน"))
   C.ItemData(9) = 9
   
   C.AddItem (MapText("ตุลาคม"))
   C.ItemData(10) = 10
   
   C.AddItem (MapText("พฤษศจิกายน"))
   C.ItemData(11) = 11

   C.AddItem (MapText(" ธันวาคม"))
   C.ItemData(12) = 12
End Sub
Public Function PatchWildCard(T As String) As String
   If Len(Trim(T)) <> 0 Then
      PatchWildCard = T & "%"
   Else
      PatchWildCard = T
   End If
End Function
Public Function CheckUniqueNs(UnqType As UNIQUE_TYPE, KEY As String, ID As Long, Optional FieldNameExTendValue As String, Optional FieldNameExTendValueEX As String, Optional NullFlag As Boolean = False) As Boolean
On Error GoTo ErrorHandler
Dim TableName As String
Dim FieldName1 As String
Dim FieldName2 As String
Dim FieldNameExTend As String
Dim FieldNameExTendEX As String
Dim Flag As Boolean
Dim Count As Long

   CheckUniqueNs = False
   
   Flag = False
   
   If UnqType = MASTER_CODE Then
      TableName = "MASTER_REF"
      FieldName1 = "KEY_CODE"
      FieldName2 = "KEY_ID"
      FieldNameExTend = "MASTER_AREA"
      Flag = True
   ElseIf UnqType = MASTER_NAME Then
      TableName = "MASTER_REF"
      FieldName1 = "KEY_NAME"
      FieldName2 = "KEY_ID"
      FieldNameExTend = "MASTER_AREA"
      Flag = True
   ElseIf UnqType = ENTERPRISE_CODE Then
      TableName = "ENTERPRISE"
      FieldName1 = "ENTERPRISE_CODE"
      FieldName2 = "ENTERPRISE_ID"
      Flag = True
   ElseIf UnqType = ENTERPRISE_NAME Then
      TableName = "ENTERPRISE"
      FieldName1 = "ENTERPRISE_NAME"
      FieldName2 = "ENTERPRISE_ID"
      Flag = True
   End If
   
   If Flag Then
      Count = glbDatabaseMngr.CountRecord(TableName, FieldName1, FieldName2, KEY, ID, glbErrorLog, FieldNameExTend, FieldNameExTendValue, FieldNameExTendEX, FieldNameExTendValueEX, NullFlag)
      If Count <> 0 Then
         CheckUniqueNs = False
      Else
         CheckUniqueNs = True
      End If
   End If
      
   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   CheckUniqueNs = False
End Function
Public Function CheckUniqueNsKey(UnqType As UNIQUE_TYPE, KEY As String) As Boolean
On Error GoTo ErrorHandler
Dim TableName As String
Dim FieldName1 As String
Dim Flag As Boolean
Dim Count As Long

   CheckUniqueNsKey = False
   
   Flag = False
   
   If UnqType = PAY_NAME Then
      TableName = "PAY_NAME"
      FieldName1 = "PAY_NAME"
      Flag = True
   ElseIf UnqType = CONFIG_CHEQUE Then
      TableName = "CHEQUE_CONFIG"
      FieldName1 = "CHEQUE_NAME"
      Flag = True
   End If
   
   If Flag Then
      Count = glbDatabaseMngr.CountRecordKey(TableName, FieldName1, KEY, glbErrorLog)
      If Count <> 0 Then
         CheckUniqueNsKey = False
      Else
         CheckUniqueNsKey = True
      End If
   End If
      
   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   CheckUniqueNsKey = False
End Function
Public Function VerifyAccessRight(Acc As String, Optional ReportName As String = "", Optional ShowError As Boolean = True) As Boolean
Dim R As CGroupRight
Dim iCount As Long
Dim TempParentID As Long
Dim FoundFlag As Boolean
   
   If glbUser.USER_NAME = "" Then
      VerifyAccessRight = True
      Exit Function
   End If
   
   Call glbDaily.StartTransaction
   Call CreatePermissionNode(Acc, TempParentID, ReportName)
   Call glbDaily.CommitTransaction
   
   FoundFlag = False
   If Not (glbAccessRight Is Nothing) Then
      For Each R In glbAccessRight
         If R.RIGHT_ITEM_NAME = Acc Then
            FoundFlag = True
            If R.RIGHT_STATUS = "Y" Then
               VerifyAccessRight = True
               Set R = Nothing
               Exit For
            Else
               VerifyAccessRight = False
               Set R = Nothing
               Exit For
            End If
         End If
      Next R
   Else
      VerifyAccessRight = True
   End If
   If FoundFlag And (Not VerifyAccessRight) Then
      VerifyAccessRight = False
      If ShowError Then
         glbErrorLog.LocalErrorMsg = "ไม่สามารถใช้งานโปรแกรมส่วนนี้ได้เนื่องจากมีสิทธ์ไม่พอเพียง -> " & Acc
         glbErrorLog.ShowUserError
      End If
   Else
      VerifyAccessRight = True
   End If
   Set R = Nothing
End Function
Private Function CreatePermissionNode(Acc As String, ParentID As Long, ReportName As String) As Boolean
Dim ParentKey As String
Dim TopFlag As Boolean
Dim TempParentID As Long
Dim CreateFlag As Boolean
Dim Ri As CRightItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long
   
   'Create node here
   Set Ri = New CRightItem
   Set TempRs = New ADODB.Recordset
   TempParentID = 0
   
   Ri.RIGHT_ID = -1
   Ri.RIGHT_ITEM_NAME = Acc
   Call Ri.QueryData(1, TempRs, iCount)
   If TempRs.EOF Then
      ParentKey = GetParentKey(Acc, TopFlag)
      If Not TopFlag Then
         Call CreatePermissionNode(ParentKey, TempParentID, ReportName)
         Ri.PARENT_ID = TempParentID
      End If
      
      Ri.AddEditMode = SHOW_ADD
      Call GetParentItemDesc(Acc, Ri, ReportName)
      Call Ri.AddEditData
      ParentID = Ri.RIGHT_ID
   Else
      Call Ri.PopulateFromRS(1, TempRs)
      ParentID = Ri.RIGHT_ID
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set Ri = Nothing
End Function
Private Function GetParentKey(Acc As String, TopFlag As Boolean) As String
Dim I As Long
Dim j As Long

   For I = 1 To Len(Acc)
      If Mid(Acc, I, 1) = "_" Then
         j = I
      End If
   Next I
   
   If j > 1 Then
      GetParentKey = Mid(Acc, 1, j - 1)
      TopFlag = False
   Else
      GetParentKey = ""
      TopFlag = True
   End If
End Function
Private Sub GetParentItemDesc(Acc As String, Ri As CRightItem, ReportName As String)
   Ri.DEFAULT_VALUE = "N"
   
   If Acc = "ADMIN" Then
      Ri.RIGHT_ITEM_DESC = "ระบบผู้ใช้งาน"
   ElseIf Acc = "ADMIN_GROUP" Then
      Ri.RIGHT_ITEM_DESC = "กลุ่มข้อมูลผู้ใช้งาน"
   ElseIf Acc = "ADMIN_USER" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลผู้ใช้งาน"
      
   ElseIf Acc = "MASTER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลัก"
   ElseIf Acc = "MASTER_MAIN" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลหลักส่วนกลาง"
   ElseIf Acc = "MASTER_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานข้อมูลหลัก"

   ElseIf Acc = "MAIN" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลส่วนกลาง"
   ElseIf Acc = "MAIN_ENTERPRISE" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลบริษัท"
   ElseIf Acc = "MAIN_CUSTOMER" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลลูกค้า"
      ElseIf Acc = "MAIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานข้อมูลส่วนกลาง"
      
   ElseIf Acc = "CHEQUE" Then
      Ri.RIGHT_ITEM_DESC = "ระบบเช็ค"
   ElseIf Acc = "CHEQUE_PRINT" Then
      Ri.RIGHT_ITEM_DESC = "พิมพ์เช็ค"
      ElseIf Acc = "CHEQUE_CONFIG" Then
      Ri.RIGHT_ITEM_DESC = "ตั้งค่าเช็ค"
   ElseIf Acc = "CHEQUE_LOG" Then
      Ri.RIGHT_ITEM_DESC = "ประวัติเช็ค"
      
         
   ElseIf Acc = "REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงาน"
   ElseIf Acc = "REPORT_CHEQUE1" Then
      Ri.RIGHT_ITEM_DESC = "รายงานประวัติเช็ค"
   
      
   ElseIf Acc = "OPTION" Then
      Ri.RIGHT_ITEM_DESC = "OPTION"
   Else
      Ri.RIGHT_ITEM_DESC = ReportName
   End If
   
End Sub
Public Function VersionToString() As String
   #If Version = 1 Then
      VersionToString = "STANDARD EDITION"
   #ElseIf Version = 2 Then
      VersionToString = "PROFESSIONAL EDITION"
   #ElseIf Version = 3 Then
      VersionToString = "PREMIUM EDITION"
   #End If
End Function
Public Function CryptStr(ByVal strInput As String, ByVal strKey As String, ByVal action As Boolean) As String
Dim I As Integer, C As Integer
Dim strOutput As String

If Len(strKey) > 0 Then
    For I = 1 To Len(strInput)
        C = CryptTableInt(Mid(strInput, I, 1))
        If action Then
            C = C + CryptTableInt(Mid(strKey, (I Mod Len(strKey)) + 1, 1))
            C = C Mod 62
        Else
            C = C - CryptTableInt(Mid(strKey, (I Mod Len(strKey)) + 1, 1))
            If C < 0 Then
                C = C + 62
            End If
        End If

        strOutput = strOutput & CryptTableChar(C)
    Next I
Else        ' ?????????????????? Encrypt
    strOutput = strInput
End If
CryptStr = strOutput
'Debug.Print strOutput
End Function
Public Function CryptTableChar(ByVal Ind As Integer) As String
Dim KeyStr As String
Select Case Ind
    
    Case 1
        KeyStr = "1"
    Case 2
        KeyStr = "2"
    Case 3
        KeyStr = "3"
    Case 4
        KeyStr = "4"
    Case 5
        KeyStr = "5"
    Case 6
        KeyStr = "6"
    Case 7
        KeyStr = "7"
    Case 8
        KeyStr = "8"
    Case 9
        KeyStr = "9"
    Case 10
        KeyStr = "0"
    Case 11
        KeyStr = "a"
    Case 12
        KeyStr = "b"
    Case 13
        KeyStr = "c"
    Case 14
        KeyStr = "d"
    Case 15
        KeyStr = "e"
    Case 16
        KeyStr = "f"
    Case 17
        KeyStr = "g"
    Case 18
        KeyStr = "h"
    Case 19
        KeyStr = "i"
    Case 20
        KeyStr = "j"
    Case 21
        KeyStr = "k"
    Case 22
        KeyStr = "l"
    Case 23
        KeyStr = "m"
    Case 24
        KeyStr = "n"
    Case 25
        KeyStr = "o"
    Case 26
        KeyStr = "p"
    Case 27
        KeyStr = "q"
    Case 28
        KeyStr = "r"
    Case 29
        KeyStr = "s"
    Case 30
        KeyStr = "t"
    Case 31
        KeyStr = "u"
    Case 32
        KeyStr = "v"
    Case 33
        KeyStr = "w"
    Case 34
        KeyStr = "x"
    Case 35
        KeyStr = "y"
    Case 36
        KeyStr = "z"
    Case 37
        KeyStr = "A"
    Case 38
        KeyStr = "B"
    Case 39
        KeyStr = "C"
    Case 40
        KeyStr = "D"
    Case 41
        KeyStr = "E"
    Case 42
        KeyStr = "F"
    Case 43
        KeyStr = "G"
    Case 44
        KeyStr = "H"
    Case 45
        KeyStr = "I"
    Case 46
        KeyStr = "J"
    Case 47
        KeyStr = "K"
    Case 48
        KeyStr = "L"
    Case 49
        KeyStr = "M"
    Case 50
        KeyStr = "N"
    Case 51
        KeyStr = "O"
    Case 52
        KeyStr = "P"
    Case 53
        KeyStr = "Q"
    Case 54
        KeyStr = "R"
    Case 55
        KeyStr = "S"
    Case 56
        KeyStr = "T"
    Case 57
        KeyStr = "U"
    Case 58
        KeyStr = "V"
    Case 59
        KeyStr = "W"
    Case 60
        KeyStr = "X"
    Case 61
        KeyStr = "Y"
    Case 0
        KeyStr = "Z"
    Case Else
        KeyStr = "|"   'It Show For Error
End Select

CryptTableChar = KeyStr
End Function
Public Function CryptTableInt(ByVal KeyStr As String) As Integer
        Dim I As Integer
        Select Case KeyStr
            
            Case "1"
                I = 1
            Case "2"
                I = 2
            Case "3"
                I = 3
            Case "4"
                I = 4
            Case "5"
                I = 5
            Case "6"
                I = 6
            Case "7"
                I = 7
            Case "8"
                I = 8
            Case "9"
                I = 9
            Case "0"
                I = 10
            Case "a"
                I = 11
            Case "b"
                I = 12
            Case "c"
                I = 13
            Case "d"
                I = 14
            Case "e"
                I = 15
            Case "f"
                I = 16
            Case "g"
                I = 17
            Case "h"
                I = 18
            Case "i"
                I = 19
            Case "j"
                I = 20
            Case "k"
                I = 21
            Case "l"
                I = 22
            Case "m"
                I = 23
            Case "n"
                I = 24
            Case "o"
                I = 25
            Case "p"
                I = 26
            Case "q"
                I = 27
            Case "r"
                I = 28
            Case "s"
                I = 29
            Case "t"
                I = 30
            Case "u"
                I = 31
            Case "v"
                I = 32
            Case "w"
                I = 33
            Case "x"
                I = 34
            Case "y"
                I = 34
            Case "z"
                I = 36
            Case "A"
                I = 37
            Case "B"
                I = 38
            Case "C"
                I = 39
            Case "D"
                I = 40
            Case "E"
                I = 41
            Case "F"
                I = 42
            Case "G"
                I = 43
            Case "H"
                I = 44
            Case "I"
                I = 45
            Case "J"
                I = 46
            Case "K"
                I = 47
            Case "L"
                I = 48
            Case "M"
                I = 49
            Case "N"
                I = 50
            Case "O"
                I = 51
            Case "P"
                I = 52
            Case "Q"
                I = 53
            Case "R"
                I = 54
            Case "S"
                I = 55
            Case "T"
                I = 56
            Case "U"
                I = 57
            Case "V"
                I = 58
            Case "W"
                I = 59
            Case "X"
                I = 60
            Case "Y"
                I = 61
            Case "Z"
                I = 0
            Case Else
                I = -1   'It Show For Error
        End Select
         
        CryptTableInt = I
    End Function
Public Function GenerateRegisterKey() As String
Dim I As Integer
Dim TempStr As String
Dim RanNo As Double
Dim AmountKey As Integer
   Randomize
   TempStr = ""
   
   #If Version = 1 Then                         'STANDARD  KEY จะมี 7 ตัว
      AmountKey = 7
   #Else
      AmountKey = 10
   #End If
   
   For I = 1 To AmountKey
      RanNo = 0
      While Not ((RanNo >= 48 And RanNo <= 57) Or (RanNo >= 65 And RanNo <= 90) Or (RanNo >= 97 And RanNo <= 122))
         RanNo = Int(Math.Rnd() * 123)
      Wend
      TempStr = TempStr & "" & Chr$(RanNo)
   Next I
      
   GenerateRegisterKey = TempStr
End Function
Public Function GetServerName() As String
Dim Fc As Long
   Fc = InStr(1, glbParameterObj.DBFile, "C:\")
   If Fc > 1 Then
      GetServerName = Left(glbParameterObj.DBFile, Fc - 2)
   Else
      GetServerName = "LOCALHOST"
   End If
End Function
Public Function BankType2Value(BankType As SHOW_BANK_TYPE, ConTrolID As Long) As Double
   If BankType = KRUNGTHAI Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 3
      End If
   ElseIf BankType = KRUNGTHEP Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 3
      End If
   ElseIf BankType = GSIKORNTHAI Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 2.5
      End If
   ElseIf BankType = UOB Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.5
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.3
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.4
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17
      ElseIf ConTrolID = 13 Then
         BankType2Value = 2.3
      End If
   ElseIf BankType = THAIPANIT Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 2.5
      End If
   ElseIf BankType = KRUNGSRI Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 4.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 2.5
      End If
   ElseIf BankType = TMB Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 2.5
      End If
   ElseIf BankType = NAKORN Then
      If ConTrolID = 1 Then
         BankType2Value = 12.5
      ElseIf ConTrolID = 2 Then
         BankType2Value = 0.2
      ElseIf ConTrolID = 3 Then
         BankType2Value = 2
      ElseIf ConTrolID = 4 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 5 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 6 Then
         BankType2Value = 5
      ElseIf ConTrolID = 7 Then
         BankType2Value = 2.5
      ElseIf ConTrolID = 8 Then
         BankType2Value = 12
      ElseIf ConTrolID = 9 Then
         BankType2Value = 3.5
      ElseIf ConTrolID = 10 Then
         BankType2Value = 1.7
      ElseIf ConTrolID = 11 Then
         BankType2Value = 0.5
      ElseIf ConTrolID = 12 Then
         BankType2Value = 17.7
      ElseIf ConTrolID = 13 Then
         BankType2Value = 2.5
      End If
   End If
End Function
