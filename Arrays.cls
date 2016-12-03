VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'STATIC Class module
Option Explicit

Private Const CLASS_NAME As String = "Arrays"

Private Enum ARRAYS_ERROR
    ERROR_CODE = vbObjectError + 513
    NOT_AN_ARRAY = ERROR_CODE + 1
End Enum

Public Function hasDimensions(ByRef arr As Variant)

    If VBA.IsArray(arr) Then
    
        hasDimensions = Not throwsRuntimeErrorWhenFindingUpperBound(arr)
    
    Else
    
        raiseNotArrayError CLASS_NAME, "hasDimensions", "arr"
    
    End If

End Function

Private Function throwsRuntimeErrorWhenFindingUpperBound(ByRef arr As Variant)
    
    On Error GoTo ErrorCaused
    
    Dim i As Long: i = UBound(arr, 1)

    throwsRuntimeErrorWhenFindingUpperBound = False

Exit Function

ErrorCaused:
        
    throwsRuntimeErrorWhenFindingUpperBound = True
    
End Function

Private Sub raiseNotArrayError(ByRef className As String, ByRef functionName As String, ByRef variableName As String)
    
    Err.Raise NOT_AN_ARRAY, className, describeNotArrayError(className, functionName, variableName)

End Sub

Private Function describeNotArrayError(ByRef className As String, ByRef functionName As String, ByRef variableName As String) As String

    describeNotArrayError = className & " >> " & functionName & ". " & variableName & " is not an array"

End Function
