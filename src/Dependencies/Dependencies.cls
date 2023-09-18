VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dependencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PackageServiceProject.src.Dependencies"
Option Explicit

'@dependencies: [ _
                    "ArrayList": 1.2.3 _
                ]

Private Type TDependencies
    Items As ArrayList
End Type

Private this As TDependencies

Private Sub Class_Initialize()
    Set this.Items = New ArrayList
End Sub

Public Sub Add(ByRef Dependency As Dependency)
    this.Items.Add Dependency
End Sub

Public Function ToString() As String
    Const Indent4 As String = vbTab & vbTab & vbTab & vbTab
    Const Indent5 As String = Indent4 & vbTab
    Dim StringValue As String

    Dim Dependency As Variant
    For Each Dependency In this.Items.ToArray()
        StringValue = StringValue & vbCrLf & Indent5 & """" & Dependency.Name & """: " & Dependency.Version & " _"
    Next

    ToString = "[ _" & StringValue & vbCrLf & Indent4 & " ]"
End Function


