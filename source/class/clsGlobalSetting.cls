VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID_TYPE As Long
Public DATE_TYPE As Long
Public FLAG_TYPE As Long
Public ACTION_TYPE As Long
Public REGION_TYPE As Long
Public USERNAME_TYPE As Long
Public IP_TYPE As Long
Public GROUPNAME_TYPE As Long
Public RIGHT_TYPE As Long
Public DESC_TYPE As Long
Public RIGHTNAME_TYPE As Long
Public PASSWORD_TYPE As Long
Public CODE_TYPE As Long
Public ADDRESS_TYPE As Long
Public ZIP_TYPE As Long
Public PHONE_TYPE As Long
Public FAX_TYPE As Long
Public EMAIL_TYPE As Long
Public WEB_TYPE As Long
Public TAXID_TYPE As Long
Public PORT_TYPE As Long
Public MONEY_TYPE As Long
Public NAME_TYPE As Long
Public THAI_COMPANY_LONG_NAME_LEN As Long
Public TELEPHONE_LEN As Long
Public NAME_LEN As Long
Public AMOUNT_LEN As Long
Public NOTE_LEN As Long

Private Sub Class_Initialize()
   
   ID_TYPE = 9
   DATE_TYPE = 19
   FLAG_TYPE = 1
   ACTION_TYPE = 1
   REGION_TYPE = 2
   USERNAME_TYPE = 30
   IP_TYPE = 50
   GROUPNAME_TYPE = 50
   RIGHT_TYPE = 100
   DESC_TYPE = 300
   RIGHTNAME_TYPE = 50
   PASSWORD_TYPE = 50
   CODE_TYPE = 20
   ADDRESS_TYPE = 100
   ZIP_TYPE = 30
   PHONE_TYPE = 30
   FAX_TYPE = 30
   EMAIL_TYPE = 30
   WEB_TYPE = 30
   TAXID_TYPE = 20
   PORT_TYPE = 4
   MONEY_TYPE = 10
   NAME_TYPE = 50
   
   THAI_COMPANY_LONG_NAME_LEN = 100
   TELEPHONE_LEN = 30
   NAME_LEN = 50
   AMOUNT_LEN = 10
   NOTE_LEN = 200
End Sub
