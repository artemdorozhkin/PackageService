VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PackageModuleReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PackageServiceProject.src.PackageReader"
Option Explicit

Private Type TPackageModuleReader
    Content As PackageContent
    Parser As ContentParser
#If DEV_MODE Then
    Module As VBComponent
#Else
    Module As Object
#End If
End Type

Private this As TPackageModuleReader

Implements IPackageReader

Private Function IPackageReader_Read(ByRef ProjectOrFilePath As Variant) As PackageContent
    If Not Utils.IsEqual(TypeName(ProjectOrFilePath), "VBProject") Then Err.Raise 13
    If Not IsModuleExists(ProjectOrFilePath) Then Err.Raise 91

    Dim Code As Variant: Code = this.Module.codeModule.Lines(1, this.Module.codeModule.CountOfLines)
    Set this.Parser = New ContentParser
    Set this.Content = this.Parser.Parse(Code)

    Set IPackageReader_Read = this.Content
End Function

Private Function IsModuleExists(ByRef Project As VBProject) As Boolean
    On Error Resume Next
    Set this.Module = Project.VBComponents(PackageServiceConstants.PackageModuleName)

    IsModuleExists = Err.Number = 0
End Function
