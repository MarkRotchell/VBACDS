VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CLASS_NAME As String = "GitTools"

Private Const FILE_EXTENSION_CLASS As String = ".cls"
Private Const FILE_EXTENSION_MODULE As String = ".bas"

Private Enum GITTOOLS_ERROR
    ERROR_CODE = vbObjectError + 513
    NOT_VALID_DIRECTORY = ERROR_CODE + 1
End Enum

Private Enum MODULE_TYPE
    CLASS_MODULE = vbext_ComponentType.vbext_ct_ClassModule
    STANDARD_MODULE = vbext_ComponentType.vbext_ct_StdModule
    FORM_MODULE = vbext_ComponentType.vbext_ct_MSForm
End Enum

Public Sub ExportSourceFiles(ByRef Project As VBProject, ByVal destinationPath As String)
    If isDirectory(destinationPath) Then
        exportSourceFilesToDirectory Project, destinationPath
    Else
        raiseNotValidDirectoryError CLASS_NAME, "ExportSourceFiles", destinationPath
    End If
End Sub
Private Function isDirectory(ByRef directoryPath As String) As Boolean
    isDirectory = (Dir(directoryPath, vbDirectory) <> vbNullString)
End Function
Private Sub exportSourceFilesToDirectory(ByRef Project As VBProject, ByVal destinationPath As String)
    Dim Module As VBComponent
    For Each Module In Project.VBComponents
        exportIfCodeModule Module, destinationPath
    Next Module
End Sub
Private Sub exportIfCodeModule(ByRef Module As VBComponent, ByVal destinationPath As String)
    If isCodeModule(Module) Then Module.Export destinationPath & Module.Name & fileExtension(Module.Type)
End Sub
Private Function isCodeModule(ByRef Module As VBComponent)
    isCodeModule = (Module.Type = CLASS_MODULE Or Module.Type = STANDARD_MODULE)
End Function
Private Function fileExtension(ByRef moduleType As MODULE_TYPE) As String
    Select Case moduleType
        Case CLASS_MODULE: fileExtension = FILE_EXTENSION_CLASS
        Case STANDARD_MODULE: fileExtension = FILE_EXTENSION_MODULE
    End Select
End Function
Private Sub raiseNotValidDirectoryError(ByRef ClassName As String, ByRef functionName As String, ByRef destinationPath As String)
    Err.Raise NOT_VALID_DIRECTORY, CLASS_NAME, describeNotValidDirectoryError(ClassName, functionName, destinationPath)
End Sub
Private Function describeNotValidDirectoryError(ByRef ClassName As String, ByRef functionName As String, ByRef destinationPath As String) As String
    describeNotValidDirectoryError = ClassName & " >> " & functionName & ". " & destinationPath & " is not a valid directory"
End Function
