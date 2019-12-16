VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private Type POINT
    X As Long
    Y As Long
End Type
'
Private Const MF_ENABLED = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_NONOTIFY = &H80&
Private Const TPM_RETURNCMD = &H100&
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'
Public Function Popup(ParamArray param()) As Long
    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim p As POINT

' get the current cursor pos in screen coordinates
    GetCursorPos p

' create an empty popup menu
    hMenu = CreatePopupMenu()

' determine # of strings in paramarray
    nMenus = 1 + UBound(param)

' put each string in the menu
    For iMenu = 1 To nMenus
' the AppendMenu function has been superseeded by the InsertMenuItem
' function, but it is a bit easier to use.
        If Trim$(CStr(param(iMenu - 1))) = "-" Then
' if the parameter is a single dash, a separator is drawn
            AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
        Else
            AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(param(iMenu - 1))
        End If
    Next iMenu

' show the menu at the current cursor location;
' the flags make the menu aligned to the right (!); enable the right button to select
' an item; prohibit the menu from sending messages and make it return the index of
' the selected item.
' the TrackPopupMenu function returns when the user selected a menu item or cancelled
' the window handle used here may be any window handle from your application
' the return value is the (1-based) index of the menu item or 0 in case of cancelling
    iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)

' release and destroy the menu (for sanity)
    DestroyMenu hMenu

' return the selected menu item's index
    Popup = iMenu

End Function
Public Function AddMenu(TempCol As Collection) As Double
Dim iMenu As Long
Dim hMenu As Long
Dim nMenus As Long
Dim p As POINT
Dim D As CMenuItem

' get the current cursor pos in screen coordinates
    GetCursorPos p

' create an empty popup menu
    hMenu = CreatePopupMenu()

' determine # of strings in paramarray

' put each string in the menu
   iMenu = 0
    For Each D In TempCol
      iMenu = iMenu + 1
' the AppendMenu function has been superseeded by the InsertMenuItem
' function, but it is a bit easier to use.
        If Trim$(D.KEYWORD) = "-" Then
' if the parameter is a single dash, a separator is drawn
            AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
        Else
            AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(D.KEYWORD)
        End If
    Next D

' show the menu at the current cursor location;
' the flags make the menu aligned to the right (!); enable the right button to select
' an item; prohibit the menu from sending messages and make it return the index of
' the selected item.
' the TrackPopupMenu function returns when the user selected a menu item or cancelled
' the window handle used here may be any window handle from your application
' the return value is the (1-based) index of the menu item or 0 in case of cancelling
    iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)

' release and destroy the menu (for sanity)
    DestroyMenu hMenu

' return the selected menu item's index
   If iMenu > 0 Then
      AddMenu = TempCol(iMenu).KEY_ID
   Else
      AddMenu = 0
   End If
End Function



