VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGitRecource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class CGitResource - abstraction of the module GitResources
' Created by Mattias Olsson 2015-05-04

' This module requires the project to have the reference Microsoft Visual Basic For Applications Extensibility 5.1
' todo - Add functionality to check the file modified date. No need to export unmodified files. May need fso for that. Or is that a good idea?
'

' Members
Public path As String
Public filesPath As String
Private xlWb As Excel.Workbook

' Custom constructor for the class.
' Param path, the path where the files should be exported to or imported from
' Return void.
Public Sub Init()
    Me.path = ThisWorkbook.path & "\"
    Me.filesPath = Me.path & "files\"
End Sub

' Function to export the VBA files in a project to folders
' Return void.
Public Sub ExportCode()
    
    Dim VBComp As VBIDE.VBComponent
    Dim i As Integer
    
    ' Load workbook
    Set xlWb = ThisWorkbook
    
    ' Create folders if there are none.
    If (Dir(Me.filesPath, vbDirectory) = "") Then
        MkDir (Me.path & "files") ' top code dir
        MkDir (Me.path & "files\module") ' module dir
        MkDir (Me.path & "files\class") ' class dir
        MkDir (Me.path & "files\form") ' form dir
    End If
    
    ' Loop through all files (components) in the workbook
    For Each VBComp In xlWb.VBProject.VBComponents
        ' Export the files
        If VBComp.Type = vbext_ct_StdModule Then
            VBComp.Export Me.filesPath & "module\" & VBComp.Name & ".bas"
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            VBComp.Export Me.filesPath & "class\" & VBComp.Name & ".cls"
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            VBComp.Export Me.filesPath & "form\" & VBComp.Name & ".frm"
        End If
        i = i + 1
    Next VBComp
    MsgBox i & " code files exported. You can now commit and push these to the Git repository. You may also want to check that no double module/class/forms has been created... ", vbInformation, "Success!"
End Sub

' Function to Import files into the VBA project
' Return void
Public Sub ImportCode()

    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim fso, topDir, d, subDir, f As Variant
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Load workbook
    Set xlWb = ThisWorkbook
    
    Set topDir = fso.getFolder(Me.path)
    Set subDir = topDir.subFolders
    
    ' If "files" dir does not exist, exit (no files exist to import)
    If Dir(Me.filesPath) = "" Then Exit Sub
    
    ' Loop through the files in each path.
    For Each d In subDir
        Debug.Print d
        For Each f In d.Files
            Debug.Print f
            If f.Name <> "CGitResource.cls" And GetExtension(f.Name) <> ".frx" Then
                ' Must remove the project's VBComps before importing, which really must be tested before using on SmartApp. Always keep a backup file too.
                Set VBComp = xlWb.VBProject.VBComponents(RemoveExtension(f.Name))
                xlWb.VBProject.VBComponents.Remove VBComp
                xlWb.VBProject.VBComponents.Import f
                i = i + 1
            End If
        Next f
    Next d
    MsgBox i & " code files imported to the file.", vbInformation, "Success!"
End Sub

' Function to remove the extension from a filename
' Param str, a filename
' Return the filename without extension i.e. file.xls -> file
Public Function RemoveExtension(ByVal str As String) As String
    Dim pos As Integer
    pos = InStrRev(str, ".")
    RemoveExtension = Left(str, pos - 1)
End Function

' Function to get the extension from a filename
' Param str, a filename
' Return the extension i.e. file.xls -> .xls
Public Function GetExtension(ByVal str As String) As String
    Dim pos As Integer
    pos = InStrRev(str, ".")
    GetExtension = Right(str, Len(str) - pos + 1)
End Function

' Native constructor
Private Sub Class_Initialize()
    Debug.Print "Class CGitResource initialized"
End Sub

' Native destructor
Private Sub Class_Terminate()
    Debug.Print "Class CGitResource terminated"
End Sub
