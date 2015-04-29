Attribute VB_Name = "Main"
Option Explicit

'Callback for CRGOneToGlobal onAction
' Function iterates through cells in the selected area and searches the DB based on the tag specified on the control argument.
' control.Tag should contain a string containing table, have and want, separated by commas. I.e "ECB3,pCodeOne,pCode" would mean "I have pCodeOne, look in ECB4 for the corresponding pCode value".
Sub OneToOneRelation(control As IRibbonControl)
    Dim db As CDatabase
    Dim res As Collection
    Dim warning As frmWarning
    
    ' Warning not to overwrite
    If (getHideConvertWarning = "0" Or getHideConvertWarning = "") Then
        Set warning = New frmWarning
        warning.lblPrompt = "You will not be able to undo this action. If you just want to see the results, you can create a new column, copy the values you want to convert, and try there so nothing important gets overwritten."
        warning.Show
        If warning.response = False Then Exit Sub
    End If
    
    Set db = New CDatabase
    db.init
    
    Dim str, table, have, want, SQL As String
    str = Split(control.tag, ",")
    
    table = str(0)
    want = str(1)
    have = str(2)
    
    Dim c, r As Variant
    
    For Each c In Selection
        With db
            SQL = .selectQuery(table, have)
            SQL = SQL & .where(want, c)
            Set res = .fetchCollection(SQL)
        End With
        For Each r In res
            c.value = "'" & r
        Next r
    Next c
End Sub

' Requires reference Microsoft Visual Basic For Applications Extensibility Library
Sub ExportSmartApp(control As IRibbonControl)
    Const path = "C:\Users\qolsmat\Desktop\vbatest\files"
    Dim xlWb As Excel.Workbook
    Dim VBComp As VBIDE.VBComponent
    
    ' Load workbook
    Set xlWb = ThisWorkbook
    
    ' Loop through all files (components) in the workbook
    For Each VBComp In xlWb.VBProject.VBComponents
        ' Export the files
        If VBComp.Type = vbext_ct_StdModule Then
            VBComp.Export path & "\module\" & VBComp.Name & ".bas"
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            VBComp.Export path & "\class\" & VBComp.Name & ".cls"
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            VBComp.Export path & "\form\" & VBComp.Name & ".frm"
        End If
    Next VBComp
    
End Sub
