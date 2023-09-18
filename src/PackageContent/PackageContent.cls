VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PackageContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PackageServiceProject.src.PackageContent"
Option Explicit

Private Type TPackageContent
    Name As String
    Version As String
    Description As String
    Dependencies As Dependencies
End Type

Private this As TPackageContent

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Get Version() As String
    Version = this.Version
End Property

Public Property Get Description() As String
    Description = this.Description
End Property

Public Property Get Dependencies() As Dependencies
    Set Dependencies = this.Dependencies
End Property

Public Sub Constructor(ByVal Name As String, ByVal Version As String, ByVal Description As String, ByVal Dependencies As Dependencies)
    With this
        .Name = Name
        .Version = Version
        .Description = Description
        Set .Dependencies = Dependencies
    End With
End Sub

Public Function ToString() As String
    Const Suffix As String = vbCrLf & "'@"

    Dim StringValue As String
    StringValue = StringValue & Suffix & "name: " & this.Name
    StringValue = StringValue & Suffix & "version: " & this.Version
    StringValue = StringValue & Suffix & "description: " & this.Description
    If Not this.Dependencies Is Nothing Then
        StringValue = StringValue & Suffix & "dependencies: " & this.Dependencies.ToString()
    End If

    ToString = StringValue
End Function
