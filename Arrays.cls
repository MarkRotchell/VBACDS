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
    NOT_ONE_DIMENSION = ERROR_CODE + 3
    NOT_SORTABLE_TYPE = ERROR_CODE + 4
End Enum

Private Type errorLocation
    Class As String
    Function As String
    Variable As String
End Type


Public Function isSorted(ByRef arr As Variant) As Boolean

    If numberOfDimensions(arr) <> 1 Then
    
        raiseNotOneDimensionError errorLocation(CLASS_NAME, "isSorted", "arr")
    
    Else
    
        isSorted = isVectorSorted(arr)
    
    End If
    
End Function

Private Function isVectorSorted(ByRef arr As Variant) As Boolean

    Dim left As Long, right As Long
    
    left = LBound(arr, 1)
    
    right = UBound(arr, 1)
    
    If isArrayOfStrings(arr) Then
        
        isVectorSorted = isStringVectorSortedBetween(arr, left, right)
    
    ElseIf isArrayOfNumeric(arr) Then
        
        isVectorSorted = isNumericalVectorSortedBetween(arr, left, right)
    
    Else
        
        raiseNotSortableDataTypeError errorLocation(CLASS_NAME, "isVectorSorted", "arr")
    
    End If
        
End Function

Private Function isStringVectorSortedBetween(ByRef arr As Variant, ByRef left As Long, ByRef right As Long) As Boolean

    Dim sorted As Boolean: sorted = True
    
    Do Until left = right Or Not sorted
        
        If VBA.StrComp(arr(left), arr(left + 1)) = 1 Then sorted = False
        
        left = left + 1
        
    Loop
    
    isStringVectorSortedBetween = sorted
    
End Function

Private Function isNumericalVectorSortedBetween(ByRef arr As Variant, ByRef left As Long, ByRef right As Long) As Boolean

    Dim sorted As Boolean: sorted = True
    
    Do Until left = right Or Not sorted
        
        If arr(left) > arr(left + 1) Then sorted = False
        
        left = left + 1
        
    Loop
    
    isNumericalVectorSortedBetween = sorted
    
End Function

Private Function isArrayOfNumeric(ByRef arr As Variant) As Boolean

    Select Case varTypeOfFirstElement(arr)
    
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
            
            isArrayOfNumeric = True
        
        Case Else
        
            isArrayOfNumeric = False
    
    End Select
        
End Function

Private Function isArrayOfStrings(ByRef arr As Variant) As Boolean

    isArrayOfStrings = (varTypeOfFirstElement(arr) = vbString)
        
End Function

Private Function varTypeOfFirstElement(ByRef arr As Variant) As VBA.VbVarType

    varTypeOfFirstElement = VBA.VarType(arr(LBound(arr, 1)))
    
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

Public Function hasDimensions(ByRef arr As Variant)

    If VBA.IsArray(arr) Then
    
        hasDimensions = Not throwsRuntimeErrorWhenFindingUpperBound(arr)
    
    Else
    
        raiseNotArrayError errorLocation(CLASS_NAME, "hasDimensions", "arr")
    
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

Private Sub raiseNotArrayError(ByRef location As errorLocation)
    
    Err.Raise NOT_AN_ARRAY, location.Class, errorDescription(location, "is not an array")

End Sub

Private Sub raiseNotOneDimensionError(ByRef location As errorLocation)
    
    Err.Raise NOT_ONE_DIMENSION, location.Class, errorDescription(location, "does note have one dimension")

End Sub

Private Function raiseNotSortableDataTypeError(ByRef location As errorLocation)
    
    Err.Raise NOT_SORTABLE_TYPE, location.Class, errorDescription(location, "is not a comparable datatype")
    
End Function

Private Function errorDescription(ByRef location As errorLocation, ByRef description As String)
    
    errorDescription = location.Class & "." & location.Function & " >> " & location.Variable & " " & description
    
End Function

Private Function errorLocation(ByRef className As String, ByRef functionName As String, ByRef variableName As String) As errorLocation

    errorLocation.Class = className
    errorLocation.Function = functionName
    errorLocation.Variable = variableName

End Function


