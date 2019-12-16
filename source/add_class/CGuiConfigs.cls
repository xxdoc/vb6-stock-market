VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuiConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GuiConfigs As Collection                                                                    '1

Public OptionMenuItems As Collection                                                                                 '1

Private Sub Class_Initialize()
   Set m_GuiConfigs = New Collection
   Set OptionMenuItems = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_GuiConfigs = Nothing
   
   Set OptionMenuItems = Nothing
End Sub

Private Sub CreateGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = New CGuiConfig
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
   
   Call m_GuiConfigs.Add(TempC, Key)
   Set TempC = Nothing
End Sub

Private Sub SetGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(Key)
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
End Sub

Private Sub GenerateOptionMenuItems(MenuMask As String, Col As Collection)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem
   
   TempKey(1) = MapText("ดึงข้อมูลยกมาจากระบบ Express")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("เกี่ยวกับโปรแกรม XIVESS-CHEQUE MANAGEMENT")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("ลงทะเบียนโปรแกรม")
   TempID(5) = 5
   TempKey(6) = MapText("-")
   TempID(6) = 6
   TempKey(7) = MapText("เชื่อมต่อฐานข้อมูลใหม่")
   TempID(7) = 7
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.Add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'lMenuChosen = oMenu.Popup("พิมพ์บนกระดาษเปล่า", "ปรับค่าหน้ากระดาษ", "-", "พิมพ์บนฟอร์มเปล่า (เหมือนจริง)", "พิมพ์บนฟอร์มเปล่า (ข้อมูล)", "ปรับค่าหน้ากระดาษ")
Private Sub GeneratePrintMenuItems(MenuMask As String, Col As Collection)
Dim Mask As String
Dim I As Long
Dim TempKey(25) As String
Dim TempID(25) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("พิมพ์บนกระดาษเปล่า")
   TempID(1) = 1
   TempKey(2) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("พิมพ์ PREVIEW (F10)")
   TempID(4) = 4
   TempKey(5) = MapText("พิมพ์ ทันที (F11)")
   TempID(5) = 5
   TempKey(6) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(6) = 6
   TempKey(7) = MapText("-")
   TempID(7) = 7
   TempKey(8) = MapText("พิมพ์  ลาเบล ส่งของ")
   TempID(8) = 8
   TempKey(9) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(9) = 9
   TempKey(10) = MapText("-")
   TempID(10) = 10
   TempKey(11) = MapText("รายงานใบแยกสถานที่จัดส่ง")
   TempID(11) = 11
   TempKey(12) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(12) = 12
   TempKey(13) = MapText("รายงานใบรับของ")
   TempID(13) = 13
   TempKey(14) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(14) = 14
   TempKey(15) = MapText("-")
   TempID(15) = 15
   TempKey(16) = MapText("ใบสำคัญจ่าย (PREVIEW)")
   TempID(16) = 16
   TempKey(17) = MapText("ใบสำคัญจ่าย")
   TempID(17) = 17
   TempKey(18) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(18) = 18
   TempKey(19) = MapText("-")
   TempID(19) = 19
   TempKey(20) = MapText("ใบเสร็จทั้งหมดจากเอกสารเป็นชุด (PREVIEW)")
   TempID(20) = 20
   TempKey(21) = MapText("ใบเสร็จทั้งหมดจากเอกสารเป็นชุด")
   TempID(21) = 21
   TempKey(22) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(22) = 22
   TempKey(23) = MapText("-")
   TempID(23) = 23
   TempKey(24) = MapText("ใบวางบิลสำหรับ MK")
   TempID(24) = 24
   TempKey(25) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(25) = 25
   
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.Add(D)
         Set D = Nothing
      End If
   Next I
End Sub
'oMenu.Popup("เพิ่มรายการใหม่ (ทีละรายการ)", "เพิ่มรายการใหม่ (เป็นชุด)")
Public Sub CreateGuiConfig(Programowner As String)
    
   '====
    Call CreateGuiConfigObj("AGENT_SERVER_CONNECT", False)
   
    Call SetGuiConfigObj("AGENT_SERVER_CONNECT", True)
   
   '====
   Call CreateGuiConfigObj("GL_VIEW", True)
    
    Call SetGuiConfigObj("GL_VIEW", False)
   
   #If Version = 3 Then
      Call GenerateOptionMenuItems("YYYYYYY", OptionMenuItems)
   #Else
      Call GenerateOptionMenuItems("NNYYYYY", OptionMenuItems)
   #End If
End Sub
Public Function VerifyGuiConfig(KeyCode As String, Optional PrintError As Boolean = False) As Boolean
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(KeyCode)
   VerifyGuiConfig = TempC.ENABLE_FLAG
   
   If Not TempC.ENABLE_FLAG Then
      If PrintError Then
         glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
         glbErrorLog.ShowUserError
      End If
   End If
End Function
