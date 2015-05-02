Attribute VB_Name = "GitResources"
Option Explicit

' This module requires reference Microsoft Visual Basic For Applications Extensibility 5.1
' todo - Add functionality to check the file modified date. No need to export unmodified files. May need fso for that.
' Test of git settings..
Sub ExportSmartApp(control As IRibbonControl)
    Const path = "C:\Users\qolsmat\Desktop\vbatest\files\"
    Dim xlWb As Excel.Workbook
    Dim VBComp As VBIDE.VBComponent
    Dim i As Integer
    
    ' Load workbook
    Set xlWb = ThisWorkbook
    
    ' Loop through all files (components) in the workbook
    For Each VBComp In xlWb.VBProject.VBComponents
        ' Export the files
        If VBComp.Type = vbext_ct_StdModule Then
            VBComp.Export path & "module\" & VBComp.Name & ".bas"
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            VBComp.Export path & "class\" & VBComp.Name & ".cls"
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            VBComp.Export path & "form\" & VBComp.Name & ".frm"
        End If
        i = i + 1
    Next VBComp
    MsgBox i & " code files exported. You can now commit and push these to the Git repository. You may also want to check that no double module/class/forms has been created... ", vbInformation, "Success!"
End Sub

'Callback for ImportSA onAction
Sub ImportSmartApp(control As IRibbonControl)
    Const path = "C:\Users\qolsmat\Desktop\vbatest\files\"
    Dim xlWb As Excel.Workbook
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim fso, topDir, d, subDir, f As Variant
    'Dim f As VBIDE.VBComponent
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Load workbook
    Set xlWb = ThisWorkbook
    
    Set topDir = fso.getFolder(path)
    Set subDir = topDir.subFolders
    
    ' Loop through the files in each path.
    For Each d In subDir
        Debug.Print d
        For Each f In d.Files
            Debug.Print f
            If f.Name <> "GitResources.bas" And Right(f.Name, 4) <> ".frx" Then
                ' Must remove module before importing, which really must be tested before using on SmartApp. Always keep a backup file too.
                Set VBComp = xlWb.VBProject.VBComponents(RemoveExtension(f.Name)) ' Ok den hittar inte pga CDatabase.cls. CDatabase går.
                ThisWorkbook.VBProject.VBComponents.Remove VBComp
                ThisWorkbook.VBProject.VBComponents.Import f
            End If
            i = i + 1
        Next f
    Next d
    i = i - 1 ' Adjust file for not importing this module.
    MsgBox i & " code files imported to the file.", vbInformation, "Success!"
End Sub

Public Function RemoveExtension(ByVal str As String) As String
    Dim pos As Integer
    pos = InStr(str, ".")
    RemoveExtension = Left(str, pos - 1)
End Function
