VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Set VsPrint(vpPrint As VSPrinter)

End Property

Public Property Get ErrorMsg() As String

End Property

Public Property Get ReadParam() As Collection

End Property
Public Function AddParam(varItem As Variant, strIndex As String) As Boolean

End Function

Public Sub ClearParam()

End Sub

Public Function PrintDoc() As Boolean

End Function

Public Function Preview() As Boolean

End Function

