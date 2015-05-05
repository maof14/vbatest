Attribute VB_Name = "SettingsModule"
Option Explicit

' Declare base functions
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpString As String, _
ByVal lpFileName As String) As Long

Private Function GetINIString(ByVal sApp As String, ByVal sKey As String, ByVal filePath As String) As String
    Dim sBuf As String * 256
    Dim lBuf As Long

    lBuf = GetPrivateProfileString(sApp, sKey, "", sBuf, Len(sBuf), filePath)
    GetINIString = Left$(sBuf, lBuf)
End Function

Private Function SetINIString(ByVal sApp As String, ByVal sKey As String, ByVal sString As String, lpFileName As String) As String
    Dim sBuf As String * 256
    Dim lBuf As Long
    SetINIString = WritePrivateProfileString(sApp, sKey, sString, lpFileName)
End Function

' To be copied from here downwards

Public Function setHideConvertWarning(ByVal sVal As Integer) As String
    Dim path As String
    path = ThisWorkbook.path & "\settings.ini"
    setHideConvertWarning = SetINIString("HideConvertWarning", "Option", sVal, path)
End Function

Public Function getHideConvertWarning() As String
    Dim path As String
    path = ThisWorkbook.path & "\settings.ini"
    getHideConvertWarning = GetINIString("HideConvertWarning", "Option", path)
End Function
