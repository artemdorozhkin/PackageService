VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PackageModuleCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PackageServiceProject.src"
Option Explicit

Private Type TPackageService
#If DEV_MODE Then
    Project As VBProject
    Module As VBComponent
#Else
    Project As Object
    Module As Object
#End If
End Type

Private this As TPackageService

Public Function Create(ByRef PackageContent As PackageContent) As Object
    Set this.Project = ThisWorkbook.VBProject

    If IsModuleExists() Then
        SetContent PackageContent
        Exit Function
    End If

    Set this.Module = this.Project.VBComponents.Add(PackageServiceConstants.StdModule)
    this.Module.Name = PackageServiceConstants.PackageModuleName
    SetContent PackageContent

    Set Create = this.Module
End Function

Private Function IsModuleExists() As Boolean
    On Error Resume Next
    Set this.Module = this.Project.VBComponents(PackageServiceConstants.PackageModuleName)

    IsModuleExists = Err.Number = 0
End Function

Private Sub SetContent(ByRef PackageContent As PackageContent)
    With this.Module.codeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString PackageContent.ToString()
    End With
End Sub
