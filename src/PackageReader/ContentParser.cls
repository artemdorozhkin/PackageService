VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PackageServiceProject.src.PackageReader"
Option Explicit

Private Type TContentParser
    Content As PackageContent
End Type

Private this As TContentParser

Public Function Parse(ByRef Code As Variant) As PackageContent
    Set this.Content = NewPackageContent(Name:=GetName(Code), _
                                         Version:=GetVersion(Code), _
                                         Description:=GetDescription(Code), _
                                         Dependencies:=GetDependencies(Code))
End Function

Private Function GetName(ByRef Code As Variant) As String
    GetName = GetPatternMatch(Code, "'@name:\s*\""(.*)\""")
End Function

Private Function GetVersion(ByRef Code As Variant) As String
    GetVersion = GetPatternMatch(Code, "'@version:\s*\""(.*)\""")
End Function

Private Function GetDescription(ByRef Code As Variant) As String
    GetDescription = GetPatternMatch(Code, "'@description:\s*\""(.*)\""")
End Function

Private Function GetDependencies(ByRef Code As Variant) As Dependencies
    Dim StartLine As Long: StartLine = GetPatternLine(Code, "[\s*")
    Dim EndLine As Long: EndLine = GetPatternLine(Code, "\s*]")

    Dim Buffer As Dependencies: Set Buffer = New Dependencies
    Dim Line As Long
    For Line = StartLine To EndLine
        Dim Match As Object: Set Match = CRegExp.Execute(Code(Line), "\""(.*)\"":\s*(.*)\s+\_", i:=True)
        If Not Match Is Nothing Then
            Buffer.Add NewDependency(Name:=Match(0).SubMatches(0), _
                                     Version:=Match(0).SubMatches(1))
        End If
    Next

    Set GetDependencies = Buffer
End Function

Private Function GetPatternMatch(ByRef Code As Variant, ByVal Pattern As String) As String
    Dim Line As Variant
    For Each Line In Code
        Dim Match As Object: Set Match = CRegExp.Execute(Line, Pattern, i:=True)
        If Not Match Is Nothing Then Exit For
    Next

    If Match Is Nothing Then Exit Function
    GetPatternMatch = Match(0).SubMatches(0)
End Function

Private Function GetPatternLine(ByRef Code As Variant, ByVal Pattern As String) As Long
    Dim Line As Long
    For Line = LBound(Code) To UBound(Code)
        Dim Match As Object: Set Match = CRegExp.Execute(Code(Line), Pattern, i:=True)
        If Not Match Is Nothing Then
            GetPatternLine = Line
            Exit Function
        End If
    Next
End Function


