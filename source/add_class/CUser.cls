VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public REAL_USER_ID As Long
Public USER_ID  As Long
Public USER_NAME As String
Public USER_PASSWORD As String
Public USER_DESC As String
Public USER_STATUS As String
Public GROUP_ID As Long
Public GROUP_NAME As String
Public GROUP_RIGHT As String
Public EXCEPTION_FLAG As String


Public OrderBy As Long
Public OrderType As Long
