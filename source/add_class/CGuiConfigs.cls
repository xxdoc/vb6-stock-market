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
   
   TempKey(1) = MapText("�֧������¡�Ҩҡ�к� Express")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("����ǡѺ����� XIVESS-CHEQUE MANAGEMENT")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("ŧ����¹�����")
   TempID(5) = 5
   TempKey(6) = MapText("-")
   TempID(6) = 6
   TempKey(7) = MapText("�������Ͱҹ����������")
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

'lMenuChosen = oMenu.Popup("����캹��д������", "��Ѻ���˹�ҡ�д��", "-", "����캹��������� (����͹��ԧ)", "����캹��������� (������)", "��Ѻ���˹�ҡ�д��")
Private Sub GeneratePrintMenuItems(MenuMask As String, Col As Collection)
Dim Mask As String
Dim I As Long
Dim TempKey(25) As String
Dim TempID(25) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("����캹��д������")
   TempID(1) = 1
   TempKey(2) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("����� PREVIEW (F10)")
   TempID(4) = 4
   TempKey(5) = MapText("����� �ѹ�� (F11)")
   TempID(5) = 5
   TempKey(6) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(6) = 6
   TempKey(7) = MapText("-")
   TempID(7) = 7
   TempKey(8) = MapText("�����  ���� �觢ͧ")
   TempID(8) = 8
   TempKey(9) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(9) = 9
   TempKey(10) = MapText("-")
   TempID(10) = 10
   TempKey(11) = MapText("��§ҹ��¡ʶҹ���Ѵ��")
   TempID(11) = 11
   TempKey(12) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(12) = 12
   TempKey(13) = MapText("��§ҹ��Ѻ�ͧ")
   TempID(13) = 13
   TempKey(14) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(14) = 14
   TempKey(15) = MapText("-")
   TempID(15) = 15
   TempKey(16) = MapText("��Ӥѭ���� (PREVIEW)")
   TempID(16) = 16
   TempKey(17) = MapText("��Ӥѭ����")
   TempID(17) = 17
   TempKey(18) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(18) = 18
   TempKey(19) = MapText("-")
   TempID(19) = 19
   TempKey(20) = MapText("����稷������ҡ�͡����繪ش (PREVIEW)")
   TempID(20) = 20
   TempKey(21) = MapText("����稷������ҡ�͡����繪ش")
   TempID(21) = 21
   TempKey(22) = MapText("��Ѻ���˹�ҡ�д��")
   TempID(22) = 22
   TempKey(23) = MapText("-")
   TempID(23) = 23
   TempKey(24) = MapText("��ҧ�������Ѻ MK")
   TempID(24) = 24
   TempKey(25) = MapText("��Ѻ���˹�ҡ�д��")
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
'oMenu.Popup("������¡������ (������¡��)", "������¡������ (�繪ش)")
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
         glbErrorLog.LocalErrorMsg = "��������ʹѺʹع�ѧ��ѹ���������ѹ���"
         glbErrorLog.ShowUserError
      End If
   End If
End Function
