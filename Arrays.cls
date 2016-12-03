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
    NOT_DIMENSIONED = ERROR_CODE + 2
End Enum

Public Function isSorted(ByRef arr As Variant)

    If hasDimensions(arr) Then
    
     'TODO is array sorted
    
    Else
    
        raiseNotDimensionedError CLASS_NAME, "isSorted", "arr"
    
    End If
    
End Function

Public Function numberOfDimensions(ByRef arr As Variant) As Long
    
    If VBA.IsArray(arr) Then
    
        numberOfDimensions = numberOfArrayDimensions(arr)
    
    Else
    
        raiseNotArrayError CLASS_NAME, "numberOfDimensions", "arr"
    
    End If
    
    
End Function

Private Function numberOfArrayDimensions(ByRef arr As Variant) As Long
    
    Dim i As Long: i = 0
    
    Do
        
        i = i + 1
    
    Loop Until throwsRuntimeErrorWhenFindingUpperBound(arr, i)
    
    numberOfArrayDimensions = i - 1

End Function


Public Function hasDimensions(ByRef arr As Variant)

    If VBA.IsArray(arr) Then
    
        hasDimensions = Not throwsRuntimeErrorWhenFindingUpperBound(arr)
    
    Else
    
        raiseNotArrayError CLASS_NAME, "hasDimensions", "arr"
    
    End If

End Function

Private Function throwsRuntimeErrorWhenFindingUpperBound(ByRef arr As Variant, Optional dimension As Long = 1)
    
    On Error GoTo ErrorCaused
    
        Dim u As Long: u = UBound(arr, dimension)
    
    On Error GoTo 0

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

Private Sub raiseNotDimensionedError(ByRef className As String, ByRef functionName As String, ByRef variableName As String)
    
    Err.Raise NOT_DIMENSIONED, className, describeNotDimensionedError(className, functionName, variableName)

End Sub

Private Function describeNotDimensionedError(ByRef className As String, ByRef functionName As String, ByRef variableName As String) As String

    describeNotDimensionedError = className & " >> " & functionName & ". " & variableName & " is not dimensioned"

End Function
