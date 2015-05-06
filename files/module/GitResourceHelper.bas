Attribute VB_Name = "GitResourceHelper"
Option Explicit

' Helper module as workaround so that Main can get imported with Git. The code in this module should never ever change (i think), and be excluded from the import. As all logic is in the CGitResource file.

Dim gr As CGitResource
Dim path As String
Dim exportPath As String
Dim chosenWB As Workbook
Dim exportPrompt As frmExport

Sub ExportSmartApp(control As IRibbonControl)
    Set exportPrompt = New frmExport
    frmExport.Show
    Set gr = New CGitResource
    If frmExport.proceed = False Then Exit Sub
    gr.Init frmExport.chosenWB
    gr.ExportCode
    Set gr = Nothing
    Set frmExport = Nothing
End Sub

Sub ImportSmartApp(control As IRibbonControl)
    Set gr = New CGitResource
    gr.Init
    gr.ImportCode
    Set gr = Nothing
End Sub
