VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WIN_MIS_REVISION = "3"
Private Const APP_REVISION = "12"

Private Const RegKey = "Software\WIN\StockMarket"
Private Reg As clsRegistry

Public DBFile As String
Public UserName As String
Public Password As String

Public LicenseIP As String
Public LicensePort As Long
Public Version As String
Public Language As Long

Public AppName As String
Public AppVersion As String

Public MainPicture As String
Public MainButton As String
Public NormalButton1 As String
Public NormalForm1 As String
Public ButtonCursor As String
Public ApplicationPath As String
Public ReportFile As String
Public LogoMain As String
Public ComputerName As String
Private Sub Class_Initialize()
Dim ErrCode As Integer
Dim No As Long
   
   Set Reg = New clsRegistry
   Reg.RgtRootKey = &H80000002 'HKEY_LOCAL_MACHINE
   Reg.RgtSubKey = RegKey

   ApplicationPath = App.Path
   ReportFile = ApplicationPath & "\" & "temp_report.html"
   
   AppName = "<MITTRAPHAP CHEQUE>"
   AppVersion = App.Major & "." & App.Minor & "." & WIN_MIS_REVISION & "." & App.Revision & "   (Firebird 2.0)"
   DBFile = ApplicationPath & "\STOCK_MARKET.GDB"
   
   UserName = "SYSDBA"
   
   LicenseIP = "127.0.0.1"
   LicensePort = 8888
   
   Call Reg.RgtGetStringValue("DBFile", DBFile, ErrCode)
   Call Reg.RgtGetStringValue("UserName", UserName, ErrCode)
   Call Reg.RgtGetStringValue("Password", Password, ErrCode)
   If Len(Password) <= 0 Then
      Password = "masterkey"
   Else
      Password = DecryptText(Password)
   End If
   Call Reg.RgtGetStringValue("LicenseIP", LicenseIP, ErrCode)
   Call Reg.RgtGetLongValue("LicensePort", LicensePort, ErrCode)
   
   Language = 1
   
   Version = App.Major & "." & App.Minor & "." & APP_REVISION & "." & App.Revision
   
   MainPicture = App.Path & "\Pictures\main1.jpg"
   MainButton = App.Path & "\Pictures\retangle_button1.jpg"
   NormalButton1 = App.Path & "\Pictures\normal_button2.jpg"
   NormalForm1 = App.Path & "\Pictures\dialog_form1.jpg"
   ButtonCursor = App.Path & "\Pictures\H_POINT.CUR"
   
   'LogoMain = App.PATH & "\Pictures\logo1.jpg"
   ComputerName = ""
End Sub

Private Sub Class_Terminate()
Dim ErrCode As Integer
   
   Call Reg.RgtSetStringValue("DBFile", DBFile, ErrCode)
   Call Reg.RgtSetStringValue("UserName", UserName, ErrCode)
   Password = EncryptText(Password)
   Call Reg.RgtSetStringValue("Password", Password, ErrCode)
   Call Reg.RgtSetStringValue("LicenseIP", LicenseIP, ErrCode)
   Call Reg.RgtSetLongValue("LicensePort", LicensePort, ErrCode)
   
   Set Reg = Nothing
End Sub
