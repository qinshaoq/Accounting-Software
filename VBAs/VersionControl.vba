Attribute VB_Name = "VersionControl"
' Set the path where the VBA code will be saved
Public Const sVBAPath As String = "C:\VBAs\"

Sub SaveCodeModules()

'This code Exports all VBA modules
Dim i%, sName$

If Dir(sVBAPath, vbDirectory) = "" Then
MkDir (sVBAPath)
End If

With ThisWorkbook.VBProject
    For i% = 1 To .VBComponents.Count
        If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
            sName$ = .VBComponents(i%).CodeModule.Name
            If sName$ = "ThisWorkbook" Then
                sName$ = ThisWorkbook.Name
            End If
            .VBComponents(i%).Export sVBAPath & sName$ & ".vba"
        End If
    Next i
End With

End Sub

Sub ImportCodeModules()

With ThisWorkbook.VBProject
    For i% = .VBComponents.Count To 1 Step -1

        ModuleName = .VBComponents(i%).CodeModule.Name
        ModuleType = .VBComponents.Item(i%).Type

        If ModuleType = 100 And ModuleName = "ThisWorkbook" Then
            Filename = ThisWorkbook.Name
            'ImportWorkbookCode sVBAPath & Filename & ".vba"
        ElseIf ModuleType = 1 And ModuleName <> "VersionControl" Then
            If Right(ModuleName, 6) = "Macros" And Dir(sVBAPath & ModuleName & ".vba") <> "" Then
                .VBComponents.Remove .VBComponents(ModuleName)
                .VBComponents.Import sVBAPath & ModuleName & ".vba"
           End If
        End If
    Next i
End With

End Sub

Private Sub ImportWorkbookCode(sFilePathName As String)

Dim sCode As String
Set fso = CreateObject("Scripting.FileSystemObject")
Set Text = fso.OpenTextFile(sFilePathName, ForReading)
Do While Not Text.AtEndOfStream
    If Text.Line < 10 Then
        Text.SkipLine
    Else
        sCode = sCode + vbCrLf + Text.ReadLine
   End If
Loop
Text.Close
Set Text = Nothing
Set fso = Nothing

With ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    .DeleteLines StartLine:=1, Count:=.CountOfLines
    .AddFromString sCode
End With
        
End Sub
