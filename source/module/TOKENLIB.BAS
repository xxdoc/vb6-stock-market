Attribute VB_Name = "modTokenLib"
Option Explicit

Type TOKENLIB
   initstate As Long
   errstr As String * 120
   errorno As Long
   TokenNo As Long
   tokenlen As Long
   bufindex As Long
   tokenstr As String * 542
   buffer As String * 512
End Type

Public Declare Sub TOKLastErrorStr Lib "tokenlib.dll" (tok As TOKENLIB, ByVal Buf As String, Length As Long)
Public Declare Function TOKAboutStr Lib "tokenlib.dll" (tok As TOKENLIB, ByVal Buf As String, Length As Long) As Long
Public Declare Function TOKInitial Lib "tokenlib.dll" (tok As TOKENLIB) As Long
Public Declare Function TOKTerminate Lib "tokenlib.dll" (tok As TOKENLIB) As Long
Public Declare Function TOKGetVersion Lib "tokenlib.dll" (tok As TOKENLIB, Major As Long, Minor As Long, Revision As Long, Build As Long) As Long
Public Declare Function TOKValidateCommand Lib "tokenlib.dll" (tok As TOKENLIB, ByVal Instr As String, ByVal Length As Long) As Long
Public Declare Function TOKEncodeString Lib "tokenlib.dll" (tok As TOKENLIB, ByVal Instr As String, ByVal Length As Long, ByVal Buf As String, BufLen As Long) As Long
Public Declare Function TOKGetTokenStr Lib "tokenlib.dll" (tok As TOKENLIB, ByVal Instr As String, ByVal Length As Long, ByVal TokenNo As Long, ByVal Buf As String, BufLen As Long) As Long

