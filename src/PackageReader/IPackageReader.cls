VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPackageReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "PackageServiceProject.src.PackageReader"
Option Explicit

Public Function Read(ByRef ProjectOrFilePath As Variant) As PackageContent
End Function
