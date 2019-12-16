Attribute VB_Name = "modRegistry"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_DWORD = 4
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_DWORD_LITTLE_ENDIAN = 4

Public Const ERROR_SUCCESS = 0&

Public Const LONG_SIZE = 4

Public Const REG_CREATE_ERROR = 1
Public Const REG_DELETE_ERROR = 2
Public Const REG_OPEN_ERROR = 3
Public Const REG_DELVALUE_ERROR = 4
Public Const REG_SETVALUE_ERROR = 5
Public Const REG_QUERY_ERROR = 6
Public Const REG_TYPE_ERROR = 7
Public Const REG_KEYNONE_ERROR = 8

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function CreateKey(ByVal HKey As Long, ByVal SubKey As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim RegResult As Long

   RegResult = RegCreateKey(HKey, SubKey, HCurKey)
   If RegResult <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      CreateKey = False
      Exit Function
   End If
   RegResult = RegCloseKey(HCurKey)
   CreateKey = True
End Function

Public Function DeleteKey(ByVal HKey As Long, ByVal SubKey As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegDeleteKey(HKey, SubKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_DELETE_ERROR
      DeleteKey = False
      Exit Function
   End If
   DeleteKey = True
End Function

Public Function DeleteValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      DeleteValue = False
      Exit Function
   End If
   
   Result = RegDeleteValue(HCurKey, ValueName)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_DELVALUE_ERROR
      DeleteValue = False
      Exit Function
   End If
   
   Result = RegCloseKey(HCurKey)
   DeleteValue = True
End Function

Public Function SetStringValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, ByVal Value As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegCreateKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      SetStringValue = False
      Exit Function
   End If
   
   Result = RegSetValueEx(HCurKey, ValueName, CLng(0), REG_SZ, ByVal Value, Len(Value))
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_SETVALUE_ERROR
      SetStringValue = False
      Exit Function
   End If

   Result = RegCloseKey(HCurKey)
   SetStringValue = True
End Function

Public Function GetStringValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, RetString As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim ValueType As Long
Dim Buffer As String
Dim BufferSize As Long
Dim Position As Integer
Dim Result As Long
   
   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      GetStringValue = False
      Exit Function
   End If
   
   Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), ValueType, ByVal CLng(0), BufferSize)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_QUERY_ERROR
      GetStringValue = False
      Exit Function
   End If

   If ValueType = REG_SZ Then
      Buffer = String(BufferSize, " ")
      Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), CLng(0), ByVal Buffer, BufferSize)
      If Result <> ERROR_SUCCESS Then
         ErrorCode = REG_QUERY_ERROR
         GetStringValue = False
         Exit Function
      End If
      
      Position = InStr(Buffer, vbNullChar)
      If Position > 0 Then
         RetString = Left(Buffer, Position - 1)
      Else
         RetString = ""
      End If
      
   Else
      ErrorCode = REG_TYPE_ERROR
      GetStringValue = False
      Exit Function
   End If
   GetStringValue = True
End Function

Public Function SetLongValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, Value As Long, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegCreateKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      SetLongValue = False
      Exit Function
   End If
   
   Result = RegSetValueEx(HCurKey, ValueName, CLng(0), REG_DWORD, Value, LONG_SIZE)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_SETVALUE_ERROR
      SetLongValue = False
      Exit Function
   End If

   Result = RegCloseKey(HCurKey)
   SetLongValue = True
End Function

Public Function GetLongValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, RetLong As Long, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim ValueType As Long
Dim Buffer As Long
Dim BufferSize As Long
Dim Position As Integer
Dim Result As Long
   
   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      GetLongValue = False
      Exit Function
   End If
   
   BufferSize = LONG_SIZE
   Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), ValueType, Buffer, BufferSize)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_QUERY_ERROR
      GetLongValue = False
      Exit Function
   End If

   If ValueType = REG_DWORD Then
      RetLong = Buffer
   Else
      ErrorCode = REG_TYPE_ERROR
      GetLongValue = False
      Exit Function
   End If
   GetLongValue = True
End Function

Public Function SetBinaryValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, Value() As Byte, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegCreateKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      SetBinaryValue = False
      Exit Function
   End If
   
   Result = RegSetValueEx(HCurKey, ValueName, CLng(0), REG_BINARY, Value(0), UBound(Value()) + 1)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_SETVALUE_ERROR
      SetBinaryValue = False
      Exit Function
   End If

   Result = RegCloseKey(HCurKey)
   SetBinaryValue = True
End Function

Public Function GetBinaryValue(ByVal HKey As Long, ByVal SubKey As String, ByVal ValueName As String, RetByte() As Byte, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim ValueType As Long
Dim Buffer() As Byte
Dim BufferSize As Long
Dim Position As Integer
Dim Result As Long
   
   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      GetBinaryValue = False
      Exit Function
   End If
   
   Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), ValueType, ByVal CLng(0), BufferSize)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_QUERY_ERROR
      GetBinaryValue = False
      Exit Function
   End If

   If ValueType = REG_BINARY Then
      ReDim Buffer(BufferSize - 1) As Byte
      
      Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), CLng(0), Buffer(0), BufferSize)
      If Result <> ERROR_SUCCESS Then
         ErrorCode = REG_QUERY_ERROR
         GetBinaryValue = False
         Exit Function
      End If
      RetByte = Buffer
   Else
      ErrorCode = REG_TYPE_ERROR
      GetBinaryValue = False
      Exit Function
   End If
   
   GetBinaryValue = True
End Function

