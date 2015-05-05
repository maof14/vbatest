VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "Export workbook"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   OleObjectBlob   =   "frmExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public exportPath As String
Public chosenWB As String

Private Sub btnCancel_Click()
    ' Cancel exporting.
    Me.Hide
End Sub

Private Sub btnOK_Click()
    ' Click OK. If SmartApp is chosen, set property me.chosenWB to ThisWorkbook.
    ' Else, fill Listbox with all open workbooks. Pick one and set chosenWB to that workbook. Open Dialog to chose save path for the modules.
    If (Me.optThisWB.value = True) Then
        Me.chosenWB = ThisWorkbook.Name
    Else
        ' Show filedialog to pick
        Dim i As Integer
        Dim exportPath As String
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "Create or choose a file folder to export to"
            .AllowMultiSelect = False
            If (.Show <> -1) Then Exit Sub
            exportPath = .SelectedItems(1)
        End With

        For i = 0 To Me.lstWB.ListCount - 1
            If Me.lstWB.Selected(i) Then
                Exit For
            End If
        Next i
        Me.chosenWB = Me.lstWB.List(i)
    End If
    Me.Hide
End Sub

Private Sub optThisWB_Click()
    ' Checking the SmartApp button. Clear the listbox and unenable it.
    Me.lstWB.Enabled = False
    Me.lstWB.Clear
End Sub

Private Sub optOtherWB_Click()
    ' Fill the Listbox with all open workbooks.
    Me.lstWB.Clear
    Dim i As Integer
    Dim wb As Workbook
    i = 0
    For Each wb In Application.Workbooks
        Me.lstWB.AddItem wb.Name, i
        i = i + 1
    Next wb
    i = 0
    Me.lstWB.Enabled = True
End Sub
