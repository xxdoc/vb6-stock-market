VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private HKey As Long
Private SubKey As String
Private LogP As String

Private Type KeyNameType
   Name As String
   ValType As Long
End Type

Public Property Get DllVersion() As String
   DllVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Property

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

Public Property Let RgtRootKey(Key As Long)
   HKey = Key
End Property

Public Property Get RgtRootKey() As Long
   RgtRootKey = HKey
End Property

Public Property Let RgtSubKey(Key As String)
   SubKey = Key
End Property

Public Property Get RgtSubKey() As String
   RgtSubKey = SubKey
End Property

'Log path of log file
Public Property Get RgtLogPath() As String
   RgtLogPath = LogP
End Property

Public Property Let RgtLogPath(Path As String)
   LogP = Path
End Property

Public Function RgtCreateKey(ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim RegResult As Long

   RegResult = RegCreateKey(HKey, SubKey, HCurKey)
   If RegResult <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      RgtCreateKey = False
      Exit Function
   End If
   RegResult = RegCloseKey(HCurKey)
   RgtCreateKey = True
End Function

Public Function RgtDeleteKey(ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegDeleteKey(HKey, SubKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_DELETE_ERROR
      RgtDeleteKey = False
      Exit Function
   End If
   RgtDeleteKey = True
End Function

Public Function RgtDeleteValue(ByVal ValueName As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      RgtDeleteValue = False
      Exit Function
   End If
   
   Result = RegDeleteValue(HCurKey, ValueName)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_DELVALUE_ERROR
      RgtDeleteValue = False
      Exit Function
   End If
   
   Result = RegCloseKey(HCurKey)
   RgtDeleteValue = True
End Function

Public Function RgtSetStringValue(ByVal ValueName As String, ByVal Value As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegCreateKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      RgtSetStringValue = False
      Exit Function
   End If
   
   Result = RegSetValueEx(HCurKey, ValueName, CLng(0), REG_SZ, ByVal Value, Len(Value))
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_SETVALUE_ERROR
      RgtSetStringValue = False
      Exit Function
   End If

   Result = RegCloseKey(HCurKey)
   RgtSetStringValue = True
End Function

Public Function RgtGetStringValue(ByVal ValueName As String, RetString As String, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim ValueType As Long
Dim Buffer As String
Dim BufferSize As Long
Dim Position As Integer
Dim Result As Long
   
   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      RgtGetStringValue = False
      Exit Function
   End If
   
   Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), ValueType, ByVal CLng(0), BufferSize)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_QUERY_ERROR
      RgtGetStringValue = False
      Exit Function
   End If

   If ValueType = REG_SZ Then
      Buffer = String(BufferSize, " ")
      Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), CLng(0), ByVal Buffer, BufferSize)
      If Result <> ERROR_SUCCESS Then
         ErrorCode = REG_QUERY_ERROR
         RgtGetStringValue = False
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
      RgtGetStringValue = False
      Exit Function
   End If
   RgtGetStringValue = True
End Function

Public Function RgtSetLongValue(ByVal ValueName As String, Value As Long, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegCreateKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      RgtSetLongValue = False
      Exit Function
   End If
   
   Result = RegSetValueEx(HCurKey, ValueName, CLng(0), REG_DWORD, Value, LONG_SIZE)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_SETVALUE_ERROR
      RgtSetLongValue = False
      Exit Function
   End If

   Result = RegCloseKey(HCurKey)
   RgtSetLongValue = True
End Function

Public Function RgtGetLongValue(ByVal ValueName As String, RetLong As Long, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim ValueType As Long
Dim Buffer As Long
Dim BufferSize As Long
Dim Position As Integer
Dim Result As Long
   
   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      RgtGetLongValue = False
      Exit Function
   End If
   
   BufferSize = LONG_SIZE
   Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), ValueType, Buffer, BufferSize)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_QUERY_ERROR
      RgtGetLongValue = False
      Exit Function
   End If

   If ValueType = REG_DWORD Then
      RetLong = Buffer
   Else
      ErrorCode = REG_TYPE_ERROR
      RgtGetLongValue = False
      Exit Function
   End If
   RgtGetLongValue = True
End Function

Public Function RgtSetBinaryValue(ByVal ValueName As String, Value() As Byte, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim Result As Long

   Result = RegCreateKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_CREATE_ERROR
      RgtSetBinaryValue = False
      Exit Function
   End If
   
   Result = RegSetValueEx(HCurKey, ValueName, CLng(0), REG_BINARY, Value(0), UBound(Value()) + 1)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_SETVALUE_ERROR
      RgtSetBinaryValue = False
      Exit Function
   End If

   Result = RegCloseKey(HCurKey)
   RgtSetBinaryValue = True
End Function

Public Function RgtGetBinaryValue(ByVal ValueName As String, RetByte() As Byte, ErrorCode As Integer) As Boolean
Dim HCurKey As Long
Dim ValueType As Long
Dim Buffer() As Byte
Dim BufferSize As Long
Dim Position As Integer
Dim Result As Long
   
   Result = RegOpenKey(HKey, SubKey, HCurKey)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_OPEN_ERROR
      RgtGetBinaryValue = False
      Exit Function
   End If
   
   Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), ValueType, ByVal CLng(0), BufferSize)
   If Result <> ERROR_SUCCESS Then
      ErrorCode = REG_QUERY_ERROR
      RgtGetBinaryValue = False
      Exit Function
   End If

   If ValueType = REG_BINARY Then
      ReDim Buffer(BufferSize - 1) As Byte
      
      Result = RegQueryValueEx(HCurKey, ValueName, CLng(0), CLng(0), Buffer(0), BufferSize)
      If Result <> ERROR_SUCCESS Then
         ErrorCode = REG_QUERY_ERROR
         RgtGetBinaryValue = False
         Exit Function
      End If
      RetByte = Buffer
   Else
      ErrorCode = REG_TYPE_ERROR
      RgtGetBinaryValue = False
      Exit Function
   End If
   
   RgtGetBinaryValue = True
End Function

