VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSCComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public COMPNAM As String
Public PATH As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT SCP.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SCCOMP SCP "
         
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY COMPNAM "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PATH "
   Else
      OrderByStr = "ORDER BY COMPNAM "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRsExpress(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
Dim I As Long
Dim TempStr As String
Dim Ch As String
 
   COMPNAM = NVLS(Rs("COMPNAM"), "")
   PATH = NVLS(Rs("PATH"), "")
   
   TempStr = ""
   For I = 1 To Len(PATH)
      Ch = Mid(PATH, I, 1)
      If (Asc(Ch) <= 122) And (Asc(Ch) >= 32) Then
        TempStr = TempStr & Ch
      Else
         Exit For
      End If
   Next I
   PATH = TempStr
End Function


