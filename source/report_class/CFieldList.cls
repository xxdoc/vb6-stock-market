VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFieldList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Fields As Collection

Private Sub Class_Initialize()
   Set Fields = New Collection
End Sub

Private Sub Class_Terminate()
   Set Fields = Nothing
   
End Sub

Public Sub AddField(F As CReportField)
Dim Temp As CReportField

   Set Temp = New CReportField
   Call Temp.CopyField(F)
   Call Fields.Add(Temp)
   Set Temp = Nothing
End Sub

Public Sub ClearField()
   Set Fields = Nothing
   Set Fields = New Collection
End Sub

Public Function GetString(Ind As Long, FormatString As String, ValueString As String) As String
Dim D As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim I As Long

   TempStr1 = ""
   TempStr2 = ""
   
   I = 0
   For Each D In Fields
      I = I + 1
      If I < Fields.Count Then
         TempStr1 = TempStr1 & D.GetFormatStr(Ind) & "|"
         TempStr2 = TempStr2 & D.GetTextStr & "|"
      Else
         TempStr1 = TempStr1 & D.GetFormatStr(Ind) & ";"
         TempStr2 = TempStr2 & D.GetTextStr & ";"
      End If
   Next D
   
   FormatString = TempStr1
   ValueString = TempStr2
End Function

