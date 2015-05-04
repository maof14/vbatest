Attribute VB_Name = "Main"
Option Explicit

' Function iterates through cells in the selected area and searches the DB based on the tag specified on the control argument.
' control.Tag should contain a string containing table, have and want, separated by commas. I.e "ECB3,pCodeOne,pCode" would mean "I have pCodeOne, look in ECB4 for the corresponding pCode value".
'Callback for OneToOneRelation onAction


Sub OneToOneRelation(control As IRibbonControl)
    Dim db As CDatabase
    Dim res As Collection
    Dim warning As frmWarning
    
    ' Warning not to overwrite
    If (getHideConvertWarning = "0" Or getHideConvertWarning = "") Then
        Set warning = New frmWarning
        warning.lblPrompt = "You will not be able to undo this action. If you just want to see the results, you can create a new column, copy the values you want to convert, and try the conversion there so nothing important gets overwritten."
        warning.Show
        If warning.response = False Then Exit Sub
        Set warning = Nothing
    End If
    
    Set db = New CDatabase
    db.Init
    
    Dim str, table, have, want, SQL As String
    str = Split(control.tag, ",")
    
    table = str(0)
    have = str(1)
    want = str(2)
    
    ' PCode special case scenario.
    If (have = "AnyPCode") Then
        If (Len(ActiveCell.value) = 3) Then ' have will be "AnyPCode". which is not len = 3. Instead, check if the cell value is len = 3. Selection(0) (array index 0, first cell)
            have = "pCodeOne"
        Else
            have = "pCode"
        End If
    End If
    
    Dim c, r As Variant
    
    For Each c In Selection
        With db
            SQL = .selectQuery(table, want)
            SQL = SQL & .where(have, c)
            Set res = .fetchCollection(SQL)
        End With
        For Each r In res
            c.value = "'" & r
        Next r
    Next c
End Sub
