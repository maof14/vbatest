VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGitResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class CGitResource - abstraction of the module GitResources
' Created by Mattias Olsson XB 20150504

' This module requires the project to have the reference Microsoft Visual Basic For Applications Extensibility 5.1
' todo - Add functionality to check the file modified date. No need to export unmodified files.
'

' Members
Public path As String
Public filesPath As String

Private xlWb As Excel.Workbook
Private VBComp As VBIDE.VBComponent

Private Enum GROverwrite
    GRNoOverwrite = 0
    GRDoOverwrite = 1
End Enum

' Custom constructor for the class.
' Param String path, the path where the files should be exported to or imported from
' Return void.
Public Sub Init(Optional ByVal wbName As String)
    ' Load workbook.
    If IsMissing(wbName) Then
        Set xlWb = ThisWorkbook
    Else
        Set xlWb = Workbooks(wbName)
    End If
    Me.path = xlWb.path & "\"
    Me.filesPath = Me.path & "files\"
End Sub

' Function to export the VBA files in a project to folders
' Return void.
Public Sub ExportCode()
    Dim i As Integer
    
    ' Create folders if there are none.
    If (Dir(Me.filesPath, vbDirectory) = "") Then
        MkDir (Me.path & "files") ' top code dir
        MkDir (Me.path & "files\module")
        MkDir (Me.path & "files\class")
        MkDir (Me.path & "files\form")
    End If
    
    ' Loop through all files (components) in the workbook VBComponents collection and export each to matching folder.
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
    InputBox i & " code files exported. You can now commit and push these to the Git repository. You may also want to check that no double module/class/forms has been created..." & vbCrLf & vbCrLf & "Copy the path to the project and paste into Git Bash.", "Success!", "cd " & Replace(xlWb.path, "\", "/")
End Sub

' Function to Import files into the VBA project
' Todo - functionality to import modules to another workbook.
' Return void
Public Sub ImportCode()
    Dim VBComps As VBIDE.VBComponents
    Dim fso, topDir, d, subDir, f As Variant
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set topDir = fso.getFolder(Me.filesPath)
    Set subDir = topDir.subFolders
    
    ' If "files" dir does not exist, exit (no files exist to import)
    If Dir(Me.filesPath) <> "" Then
        MsgBox "You do not seem to have any " & Chr(34) & "files" & Chr(34) & " directory in this folder, exiting.", vbCritical, "Error!"
        Exit Sub
    End If
    
    ' Loop through the files in each path.
    For Each d In subDir
        For Each f In d.Files
            If f.Name <> "CGitResource.cls" And GetExtension(f.Name) <> ".frx" And f.Name <> "GitResourceHelper.bas" Then
                ' CGitResource and GitResourceHelper should be imported from the VBE only. Those can't be imported as they run this, or helps to do so.
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
' Param String fileName, a filename
' Return the filename without extension i.e. file.xls -> file
Public Function RemoveExtension(ByVal fileName As String) As String
    Dim pos As Integer
    pos = InStrRev(fileName, ".")
    RemoveExtension = Left(fileName, pos - 1)
End Function

' Function to get the extension from a filename
' Param String fileName, a filename
' Return the extension i.e. file.xls -> .xls
Public Function GetExtension(ByVal fileName As String) As String
    Dim pos As Integer
    pos = InStrRev(fileName, ".")
    GetExtension = Right(fileName, Len(fileName) - pos + 1)
End Function

' Native constructor
Private Sub Class_Initialize()
    Debug.Print "Class CGitResource initialized"
End Sub

' Native destructor
Private Sub Class_Terminate()
    Debug.Print "Class CGitResource terminated"
End Sub
