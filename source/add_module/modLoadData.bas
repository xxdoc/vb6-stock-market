Attribute VB_Name = "modLoadData"
Option Explicit
Public Sub LoadMasterRef(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterArea As MASTER_TYPE)
On Error GoTo ErrorHandler
Dim D As CMasterRef
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long

   Set D = New CMasterRef
   Set Rs = New ADODB.Recordset
   
   D.KEY_ID = -1
   D.MASTER_AREA = MasterArea
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.KEY_NAME)
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitUserGroupOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(1) = 1
End Sub

Public Sub InitUserStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใช้งานได้")
   C.ItemData(1) = 1

   C.AddItem ("ถูกระงับ")
   C.ItemData(2) = 2
End Sub
Public Sub InitUserOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ชื่อผู้ใช้"))
   C.ItemData(1) = 1
End Sub
Public Sub LoadUserGroup(Ug As CUserGroup, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserGroup
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Ug.GROUP_ID = -1
   Call Ug.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_NAME)
         C.ItemData(I) = TempData.GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccessRight(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGroupRight
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupRight
Dim I As Long

   Set D = New CGroupRight
   Set Rs = New ADODB.Recordset
   
   D.GROUP_RIGHT_ID = -1
   D.GROUP_ID = GroupID
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGroupRight
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RIGHT_ITEM_NAME)
         C.ItemData(I) = TempData.GROUP_RIGHT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitChequeDateType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("DD/MM/YYYY")
   C.ItemData(1) = 1
   
   C.AddItem ("D   D  M   M   Y   Y    Y    Y")
   C.ItemData(2) = 2
   
   C.AddItem ("D   D  M   M   Y   Y ")
   C.ItemData(3) = 3
End Sub
Public Sub InitPassFlag(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ผ่านเช็ค")
   C.ItemData(1) = 1

   C.AddItem ("ไม่ผ่านเช็ค")
   C.ItemData(2) = 2
End Sub
Public Function PassChequeToString(ID As Long) As String
   If ID = 1 Then
      PassChequeToString = "Y"
   ElseIf ID = 2 Then
      PassChequeToString = "N"
   Else
      PassChequeToString = ""
   End If
End Function
Public Sub InitCancelFlag(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ยกเลิก")
   C.ItemData(1) = 1

   C.AddItem ("ไม่ยกเลิก")
   C.ItemData(2) = 2
   
End Sub
Public Sub InitChequeOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อผู้จ่าย")
   C.ItemData(1) = 11
   
   C.AddItem ("วันที่พิมพ์เช็ค")
   C.ItemData(2) = 12
   
   C.AddItem ("วันที่เช็ค")
   C.ItemData(3) = 13
   
   C.AddItem ("เลขที่เช็ค")
   C.ItemData(4) = 14
   
   C.AddItem ("ยอดเงิน")
   C.ItemData(5) = 15
   
   C.AddItem ("สาขาธนาคาร")
   C.ItemData(6) = 16
End Sub
Public Sub LoadChequeConfig(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CChequeConfig
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CChequeConfig
Dim I As Long

   Set D = New CChequeConfig
   Set Rs = New ADODB.Recordset
   
   D.CHEQUE_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CChequeConfig
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CHEQUE_NAME)
         C.ItemData(I) = I
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.CHEQUE_NAME))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAllPrintCheque(Cl As Collection, Optional STOCK_MARKET_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CStockMarketDetail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockMarketDetail
Dim I As Long
   
   Set D = New CStockMarketDetail
   Set Rs = New ADODB.Recordset
   
   D.STOCK_MARKET_ID = STOCK_MARKET_ID
   Call D.QueryData(2, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockMarketDetail
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPrintChequeHeader(Cl As Collection, Optional STOCK_MARKET_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CStockMarket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockMarket
Dim I As Long
   
   Set D = New CStockMarket
   Set Rs = New ADODB.Recordset
   
   D.STOCK_MARKET_ID = STOCK_MARKET_ID
   Call D.QueryData(1, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockMarket
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

