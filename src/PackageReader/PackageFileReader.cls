VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PackageFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PackageServiceProject.src.PackageReader"
Option Explicit

Private Type TPackageFileReader
    Content As PackageContent
    Parser As ContentParser
#If DEV_MODE Then
    FSO As FileSystemObject
    Stream As TextStream
#Else
    FSO As Object
    Stream As Object
#End If
End Type

Private this As TPackageFileReader

Implements IPackageReader

Private Function IPackageReader_Read(ByRef ProjectOrFilePath As Variant) As PackageContent
    If VarType(ProjectOrFilePath) <> vbString Then Err.Raise 13
    Set this.FSO = CreateObject("Scripting.FileSystemObject")
    If Not this.FSO.FileExists(ProjectOrFilePath) Then Err.Raise 53

    Dim Code As Variant: Code = ReadFile(ProjectOrFilePath)
    Set this.Parser = New ContentParser
    Set this.Content = this.Parser.Parse(Code)

    Set IPackageReader_Read = this.Content
End Function

Private Function ReadFile(ByRef Project As VBProject) As Variant
    Set this.Stream = this.FSO.GetFile(ProjectOrFilePath).OpenAsTextStream()

    Dim Lines As ArrayList: Set Lines = New ArrayList
    While this.Stream.AtEndOfStream
        Lines.Add this.Stream.Line
    Wend

    ReadFile = Lines.ToArray()
End Function
