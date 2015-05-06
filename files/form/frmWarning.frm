VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWarning 
   Caption         =   "Warning!"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   OleObjectBlob   =   "frmWarning.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public response As Boolean
Private wt As Integer
' Logic here:
' User presses cancel: Cancel processing the conversion. Warning should appear over and over again.
' User presses OK: Proceed with conversion.
' User has ticked the checkbox and presses OK: Proceed with conversion and write true to the settings file.
' HideConversionWarning = True = Do not show the warning.

Public Enum EWarningType
    WConvertWarning = 0
    ' Whatever else warnings there could be...
End Enum

Public Sub Init(WarningType As EWarningType)
    wt = WarningType
End Sub

Private Sub btnCancel_Click()
    ' Cancel conversion.
    Me.Hide
    response = False
End Sub

Private Sub btnOK_Click()
    ' Proceed with conversion, leave settings untouched.
    If (Me.chbDontShow.value = True) Then
        If wt = WConvertWarning Then
            setHideConvertWarning ("1")
        End If
    End If
    Me.Hide
    response = True
End Sub
