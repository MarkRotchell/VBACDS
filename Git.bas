Attribute VB_Name = "Git"
Option Explicit

Public Sub ExportSourceFiles(destPath As String)
   Dim component As VBComponent
   For Each component In Application.VBE.ActiveVBProject.VBComponents
      If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
         component.Export destPath & component.Name & ToFileExtension(component.Type)
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
    Dim Project As VBProject
    Set Project = Application.VBE.ActiveVBProject
     
    Dim comp As VBComponent
    For Each comp In Project.VBComponents
        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
        Project.VBComponents.Remove comp
        End If
    Next
End Sub

Public Sub ImportSourceFiles(sourcePath As String)
Dim file As String
file = Dir(sourcePath)
    While (file <> vbNullString)
        Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        file = Dir
    Wend
End Sub

