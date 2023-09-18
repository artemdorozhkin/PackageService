Attribute VB_Name = "Utils"
'@Folder "PackageServiceProject.src.Common"
Option Explicit

Public Sub ExportProject()
    Dim RootFolder As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Укажите папку для экспорта проекта " & ThisWorkbook.VBProject.Name
        .Show
        If .SelectedItems.Count > 0 Then
            RootFolder = .SelectedItems.Item(1)
        End If
    End With

    If IsEmpty(RootFolder) Then Exit Sub

        Dim Component As VBComponent
        For Each Component In ThisWorkbook.VBProject.VBComponents
            Dim Code As String: Code = Component.codeModule.Lines(1, Component.codeModule.CountOfLines)
            If InString(Code, vbTextCompare, "ipm-modules") And Not InString(Code, vbTextCompare, "ExportProject") Then Goto Continue
                If InString(Code, vbTextCompare, "src") Then
                    Dim Match As Object: Set Match = CRegExp.Execute(Code, "@folder\s*\(*\""(.*)\""", i:=True)
                    If Match Is Nothing Then Goto Continue

                        Dim Folder As String: Folder = Match(0).SubMatches(0)
                        Folder = Strings.Split(Folder, ".", 2)(UBound(Strings.Split(Folder, ".", 2)))
                        Folder = RootFolder & Application.PathSeparator & Strings.Replace(Folder, ".", Application.PathSeparator)

                        Dim FSO As FileSystemObject: Set FSO = New FileSystemObject
                        If Not FSO.FolderExists(Folder) Then
                            FSO.CreateFolder Folder
                        End If

                        Dim FilePath As String
                        FilePath = Folder & Application.PathSeparator & Component.Name
                        If Component.Type = vbext_ct_ClassModule Then FilePath = FilePath & ".cls"
                            If Component.Type = vbext_ct_StdModule Then FilePath = FilePath & ".bas"

                                Component.Export FilePath
                            End If
 Continue:
                            Next
End Sub

Public Function IsEqual(Byval String1 As String, Byval String2 As String, Optional Byval Compare As VbCompareMethod = vbTextCompare) As Boolean
    IsEqual = Strings.StrComp(String1, String2, Compare) = 0
End Function

Public Sub NewConstructor(Byval ClassName As String)
    Dim Project As VBProject: Set Project = ThisWorkbook.VBProject

    On Error Resume Next
    Dim Class As String: Class = Project.VBComponents(ClassName).Name
    Err.Clear: On Error Goto 0
    If Err.Number = 0 Then
        CreateConstructor Project, Class
     Exit Sub
    End If

    Dim Component As VBComponent
    For Each Component In Project.VBComponents
        If IsEqual(Component.Name, ClassName) Then
            CreateConstructor Project, Component.Name
         Exit Sub
        End If
        Next

        CreateConstructor Project, ClassName
End Sub

Private Sub CreateConstructor(Byref Project As VBProject, Byval ClassName As String)
    Dim Code As String
    Code = Code & "'@Folder(""" & Project.Name & ".src." & ClassName & """)" & vbCrLf
    Code = Code & "Option Explicit" & vbCrLf & vbCrLf
Code = Code & "Public Function New" & ClassName & "() As " & ClassName & vbCrLf
    Code = Code & vbTab & "Set New" & ClassName & " = New " & ClassName & vbCrLf
    Code = Code & "End Function"

    With Project.VBComponents.Add(PackageServiceConstants.StdModule)
        .Name = ClassName & "Constructor"
        .codeModule.DeleteLines 1, .codeModule.CountOfLines
        .codeModule.AddFromString Code
    End With

End Sub
