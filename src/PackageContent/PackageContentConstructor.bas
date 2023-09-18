Attribute VB_Name = "PackageContentConstructor"
'@Folder "PackageServiceProject.src.PackageContent"
Option Explicit

Public Function NewPackageContent(Optional ByVal Name As String, Optional ByVal Version As String, Optional ByVal Description As String, Optional ByVal Dependencies As Dependencies) As PackageContent
    Set NewPackageContent = New PackageContent
    NewPackageContent.Constructor Name, Version, Description, Dependencies
End Function
