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

        If ModuleType = 1 And ModuleName <> "VersionControl" Then
            If Right(ModuleName, 6) = "Macros" Then
                .VBComponents.Remove .VBComponents(ModuleName)
                .VBComponents.Import sVBAPath & ModuleName & ".vba"
           End If
        End If
    Next i
End With

End Sub
