Attribute VB_Name = "DevTools"
Public Sub ExportSourceFiles() 'Optional destPath As String)
'go to in VBA Editor(Extras/Tools)->(Verweise/References)-> check on "Microsoft Visual Basic for Application Extensibility 5.3"
'also necessary https://support.microsoft.com/de-at/help/813969/you-may-receive-an-run-time-error-when-you-programmatically-allow-acce
'see image in documentation folder

destPath = "U:\Tools\Fundation_VBA\Modules\" 'comment out if used with argument
'destPath == "C:\xLocal\2000_Tools\Fundation_VBA\Modules\"
Dim component As VBComponent
For Each component In Application.VBE.ActiveVBProject.VBComponents
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
        component.Export destPath & component.Name & ToFileExtension(component.Type)
        Debug.Print "Exporting: "; destPath & component.Name & ToFileExtension(component.Type)
    End If
Next
 
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function
Public Sub RemoveAllModules()
Dim project As VBProject
Set project = Application.VBE.ActiveVBProject
 
Dim comp As VBComponent
For Each comp In project.VBComponents
If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
    Debug.Print "Removing component: " & comp.Name
    project.VBComponents.Remove comp
End If
Next
End Sub
Public Sub ImportSourceFiles() 'Optional sourcePath As String)
'
SourcePath = "U:\Tools\Fundation_VBA\Modules\" 'comment out if used with argument
'sourcePath = "C:\xLocal\2000_Tools\Fundation_VBA\Modules\" 'comment out if used with argument
Dim file As String
file = Dir(SourcePath)
While (file <> vbNullString)
    If Not file = "DevTools.bas" Then
        Application.VBE.ActiveVBProject.VBComponents.Import SourcePath & file
        Debug.Print "Importing: "; SourcePath & file
    End If
    file = Dir
Wend
End Sub

