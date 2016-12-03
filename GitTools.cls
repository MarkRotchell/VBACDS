VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "GitTools"

Private Const FILE_EXTENSION_CLASS As String = ".cls"
Private Const FILE_EXTENSION_MODULE As String = ".bas"
Private Const FILE_EXTENSION_FORM As String = ".frm"

Private Enum GITTOOLS_ERROR
    ERROR_CODE = vbObjectError + 513
    NOT_VALID_DIRECTORY = ERROR_CODE + 1
End Enum

Private Enum MODULE_TYPE
    CLASS_MODULE = vbext_ComponentType.vbext_ct_ClassModule
    STANDARD_MODULE = vbext_ComponentType.vbext_ct_StdModule
    FORM_MODULE = vbext_ComponentType.vbext_ct_MSForm
End Enum

Public Sub Export()
    Me.ExportSourceFiles Application.VBE.VBProjects("VBACDS"), "C:\VBACDS\ForGit\"
End Sub

Public Sub Remove()
    Me.RemoveAllCodeModules Application.VBE.VBProjects("VBACDS"), "GitTools"
End Sub

Public Sub Import()
    Me.ImportSourceFiles Application.VBE.VBProjects("VBACDS"), "C:\VBACDS\ForGit\"
End Sub

Public Sub ImportSourceFiles(ByRef project As VBProject, ByRef sourcePath As String)
    
    If isDirectory(sourcePath) Then
    
        ImportSourceFilesFromDirectory project, sourcePath
    
    Else
    
        raiseNotValidDirectoryError CLASS_NAME, "importSourceFiles", sourcePath
    
    End If
    
End Sub

Private Sub ImportSourceFilesFromDirectory(ByRef project As VBProject, ByRef sourcePath As String)

    Dim fileName As String: fileName = Dir(sourcePath & "*")
    
    Do Until fileName = vbNullString
    
        project.VBComponents.Import sourcePath & fileName
        
        fileName = Dir
    
    Loop

End Sub


Public Sub RemoveAllCodeModules(ByRef project As VBProject, Optional ByRef Exception As String)

    Dim module As VBComponent
    
    For Each module In project.VBComponents
    
        RemoveIfCodeModule module, Exception
        
    Next module
    
End Sub

Private Sub RemoveIfCodeModule(ByRef module As VBComponent, Optional ByRef Exception As String)

    If isCodeModule(module) And module.Name <> Exception Then
        
        module.Collection.Remove module
    
    End If

End Sub

Public Sub ExportSourceFiles(ByRef project As VBProject, ByVal destinationPath As String)

    If isDirectory(destinationPath) Then
        
        exportSourceFilesToDirectory project, destinationPath
    
    Else
        
        raiseNotValidDirectoryError CLASS_NAME, "ExportSourceFiles", destinationPath
    
    End If

End Sub

Private Function isDirectory(ByRef directoryPath As String) As Boolean

    isDirectory = (VBA.Dir(directoryPath, vbDirectory) <> vbNullString)

End Function

Private Sub exportSourceFilesToDirectory(ByRef project As VBProject, ByVal destinationPath As String)
    
    Dim module As VBComponent
    
    For Each module In project.VBComponents
        
        exportIfCodeModule module, destinationPath
    
    Next module

End Sub

Private Sub exportIfCodeModule(ByRef module As VBComponent, ByVal destinationPath As String)
    
    If isCodeModule(module) Then
    
        module.Export destinationPath & module.Name & fileExtension(module.Type)
        
    End If
    
End Sub

Private Function isCodeModule(ByRef module As VBComponent)
    
    isCodeModule = (module.Type = CLASS_MODULE Or module.Type = STANDARD_MODULE)

End Function

Private Function fileExtension(ByRef moduleType As MODULE_TYPE) As String
    
    Select Case moduleType
        
        Case CLASS_MODULE: fileExtension = FILE_EXTENSION_CLASS
        
        Case STANDARD_MODULE: fileExtension = FILE_EXTENSION_MODULE
        
        Case FORM_MODULE: fileExtension = FILE_EXTENSION_FORM
    
    End Select

End Function

Private Sub raiseNotValidDirectoryError(ByRef className As String, ByRef functionName As String, ByRef directoryPath As String)
    
    Err.Raise NOT_VALID_DIRECTORY, CLASS_NAME, describeNotValidDirectoryError(className, functionName, directoryPath)

End Sub

Private Function describeNotValidDirectoryError(ByRef className As String, _
        ByRef functionName As String, ByRef destinationPath As String) As String
    
    describeNotValidDirectoryError = className & " >> " & functionName & _
                                     ". " & destinationPath & " is not a valid directory"

End Function
