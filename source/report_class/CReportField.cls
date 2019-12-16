VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Width As Double
Public Align As String
Public FieldText As String
Public BodyAlign As String
Public TempKey As String

Public Function GetFormatStr(Ind As Long) As String
   If Ind = 1 Then
      GetFormatStr = Align & Width
   Else
      GetFormatStr = BodyAlign & Width
   End If
End Function

Public Function GetTextStr() As String
   GetTextStr = FieldText
End Function

Public Sub CopyField(D As CReportField)
   Width = D.Width
   Align = D.Align
   FieldText = D.FieldText
   BodyAlign = D.BodyAlign
   TempKey = D.TempKey
End Sub

Public Sub SetFieldValue(W As Double, Al As String, T As String, Optional BodyAl As String = "^", Optional Key As String = "")
   Width = W
   Align = Al
   FieldText = T
   BodyAlign = BodyAl
   TempKey = Key
End Sub

