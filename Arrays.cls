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

Private Const NOT_FOUND As Long = -1

Private Enum ARRAYS_ERROR
    ERROR_CODE = vbObjectError + 513
    NOT_AN_ARRAY = ERROR_CODE + 1
    NOT_DIMENSIONED = ERROR_CODE + 2
    NOT_ONE_DIMENSION = ERROR_CODE + 3
    NOT_NUMERIC = ERROR_CODE + 4
End Enum

Private Type errorLocation
    Class As String
    Function As String
    Variable As String
End Type

Public Function linearSearch(ByRef arr As Variant, ByRef item As Variant) As Long
    
    Const FUNCTION_NAME As String = "linearSearch"
    
    If numberOfDimensions(arr) <> 1 Then
    
        raiseNotOneDimensionError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
        
    ElseIf Not isArrayOfNumeric(arr) Then
        
        raiseNotNumericError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
    
    ElseIf Not VBA.isNumeric(item) Then
    
        raiseNotNumericError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
    
    Else
        
        linearSearch = linearSearchNumericVector(arr, item)
        
    End If
    
End Function

Private Function linearSearchNumericVector(ByRef arr As Variant, item As Variant) As Long
        
    Dim i As Long: For i = LBound(arr) To UBound(arr)
    
        If arr(i) = item Then
            
            linearSearchNumericVector = i
        
            Exit Function
        
        End If
    
    Next i
    
    linearSearchNumericVector = NOT_FOUND
    
End Function


Public Function isSorted(ByRef arr As Variant, Optional ByRef descending As Boolean = False) As Boolean

    Const FUNCTION_NAME As String = "isSorted"
    
    If numberOfDimensions(arr) <> 1 Then
    
        raiseNotOneDimensionError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
    
    ElseIf Not isArrayOfNumeric(arr) Then
        
        raiseNotNumericError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
    
    ElseIf descending Then
    
        isSorted = isNumericVectorSortedDescending(arr)
        
    Else
    
        isSorted = isNumericVectorSortedAscending(arr)
    
    End If
    
End Function


Private Function isNumericVectorSortedAscending(ByRef arr As Variant) As Boolean
    
    Dim i As Long: For i = LBound(arr) To UBound(arr) - 1
    
        If arr(i) > arr(i + 1) Then
        
            isNumericVectorSortedAscending = False
            
            Exit Function
            
        End If
        
    Next i
    
    isNumericVectorSortedAscending = True
    
End Function

Private Function isNumericVectorSortedDescending(ByRef arr As Variant) As Boolean
    
    Dim i As Long: For i = LBound(arr) To UBound(arr) - 1
    
        If arr(i) < arr(i + 1) Then
        
            isNumericVectorSortedDescending = False
            
            Exit Function
            
        End If
        
    Next i
    
    isNumericVectorSortedDescending = True
    
End Function

Private Function varTypeOfFirstElement(ByRef arr As Variant) As VBA.VbVarType

    varTypeOfFirstElement = VBA.VarType(arr(LBound(arr, 1)))
    
End Function

Private Function isArrayOfNumeric(ByRef arr As Variant) As Boolean

    Select Case varTypeOfFirstElement(arr)
    
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
            
            isArrayOfNumeric = True
        
        Case Else
        
            isArrayOfNumeric = False
    
    End Select
        
End Function

Public Function numberOfDimensions(ByRef arr As Variant) As Long
    
    If VBA.IsArray(arr) Then
    
        numberOfDimensions = numberOfArrayDimensions(arr)
    
    Else
    
        raiseNotArrayError errorLocation(CLASS_NAME, "numberOfDimensions", "arr")
    
    End If
    
    
End Function

Private Function numberOfArrayDimensions(ByRef arr As Variant) As Long
    
    Dim i As Long: i = 0
    
    Do
        
        i = i + 1
    
    Loop Until throwsRuntimeErrorWhenFindingUpperBound(arr, i)
    
    numberOfArrayDimensions = i - 1

End Function

Public Function hasDimensions(ByRef arr As Variant) As Boolean

    If VBA.IsArray(arr) Then
    
        hasDimensions = Not throwsRuntimeErrorWhenFindingUpperBound(arr)
    
    Else
    
        raiseNotArrayError errorLocation(CLASS_NAME, "hasDimensions", "arr")
    
    End If

End Function

Private Function throwsRuntimeErrorWhenFindingUpperBound(ByRef arr As Variant, Optional dimension As Long = 1) As Boolean
    
    On Error GoTo ErrorCaused
    
        Dim u As Long
        
        u = UBound(arr, dimension)
    
    On Error GoTo 0

    throwsRuntimeErrorWhenFindingUpperBound = False

Exit Function

ErrorCaused:
        
    throwsRuntimeErrorWhenFindingUpperBound = True
    
End Function


'
' Private Functions for raising errors
'


Private Function errorLocation(ByRef className As String, ByRef functionName As String, ByRef variableName As String) As errorLocation

    errorLocation.Class = className
    
    errorLocation.Function = functionName
    
    errorLocation.Variable = variableName

End Function

Private Sub raiseNotArrayError(ByRef location As errorLocation)
    
    Err.Raise NOT_AN_ARRAY, location.Class, errorDescription(location, "is not an array")

End Sub

Private Sub raiseNotOneDimensionError(ByRef location As errorLocation)
    
    Err.Raise NOT_ONE_DIMENSION, location.Class, errorDescription(location, "does note have one dimension")

End Sub

Private Sub raiseNotNumericError(ByRef location As errorLocation)
    
    Err.Raise NOT_NUMERIC, location.Class, errorDescription(location, "is not a numeric datatype")
    
End Sub

Private Function errorDescription(ByRef location As errorLocation, ByRef description As String) As String
    
    errorDescription = location.Class & "." & location.Function & " >> " & location.Variable & " " & description
    
End Function




