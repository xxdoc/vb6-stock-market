VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ControlType As String
Public ControlIndex As Long
Public AllowNull As Boolean
Public AssociateLabelIndex As Long
Public TextMsg As String
Public Top As Long
Public Left As Long
Public Width As Long
Public Param1 As String
Public Param2 As String
Public ComboLoadID As Long
Public uctlLookUpID As Long
Public OldLine As Boolean
