Attribute VB_Name = "GitResourceHelper"
Option Explicit

' Helper module as workaround so that Main can get imported with Git. The code in this module should never ever change (i think), and be excluded from the import. As all logic is in the CGitResource file.

Dim gr As CGitResource
Dim path As String

Sub ExportSmartApp(control As IRibbonControl)
    Set gr = New CGitResource
    gr.Init
    gr.ExportCode
    Set gr = Nothing
End Sub

Sub ImportSmartApp(control As IRibbonControl)
    Set gr = New CGitResource
    gr.Init
    gr.ImportCode
    Set gr = Nothing
End Sub
