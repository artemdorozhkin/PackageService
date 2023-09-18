Attribute VB_Name = "DependencyConstructor"
'@Folder "PackageServiceProject.src.Dependencies"
Option Explicit

Public Function NewDependency(ByVal Name As String, ByVal Version As String) As Dependency
    Set NewDependency = New Dependency
    NewDependency.Name = Name
    NewDependency.Version = Version
End Function
