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

    Const FUNCTION_NAME As String = "isSorted"

    If numberOfDimensions(arr) <> 1 Then
    
        raiseNotOneDimensionError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
    
    ElseIf Not isArrayOfSortables(arr) Then
        
        raiseNotSortableDataTypeError errorLocation(CLASS_NAME, FUNCTION_NAME, "arr")
    
    Else
    
        isSorted = isVectorSorted(arr)
    
    End If
    
End Function

Private Function isVectorSorted(ByRef arr As Variant) As Boolean
    
    Dim isStrings As Boolean: isStrings = isArrayOfStrings(arr)
    
    Dim i As Long: For i = LBound(arr, 1) To UBound(arr, 1) - 1
    
        If isAfterNeighbour(arr, i, isStrings) Then
        
            isVectorSorted = False
            
            Exit Function
            
        End If
        
    Next
    
    isVectorSorted = True
        
End Function

Private Function isAfterNeighbour(ByRef arr As Variant, ByRef Index As Long, ByRef isStrings As Boolean)

    Const STRING1_AFTER_STRING2 As Byte = 1

    If isStrings Then
        
        isAfterNeighbour = VBA.StrComp(arr(Index), arr(Index + 1)) = STRING1_AFTER_STRING2
        
    Else
    
        isAfterNeighbour = arr(Index) > arr(Index + 1)
        
    End If
    
End Function

Private Function quickAreNumericsSorted(ByRef arr As Variant) As Boolean
    
    Dim i As Long: For i = LBound(arr) To UBound(arr) - 1
    
        If arr(i) > arr(i + 1) Then
        
            quickAreNumericsSorted = False
            
            Exit Function
            
        End If
        
    Next i
    
    quickAreNumericsSorted = True
    
End Function

Private Function quickAreStringsSorted(ByRef arr As Variant) As Boolean
    
    Const STRING1_AFTER_STRING2 As Byte = 1

    Dim i As Long: For i = LBound(arr) To UBound(arr) - 1
    
        If VBA.StrComp(arr(i), arr(i + 1)) = STRING1_AFTER_STRING2 Then
        
            quickAreStringsSorted = False
            
            Exit Function
            
        End If
        
    Next i
    
    quickAreStringsSorted = True
    
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

Private Function isArrayOfSortables(ByRef arr As Variant) As Boolean

    Select Case varTypeOfFirstElement(arr)
    
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle, vbString
            
            isArrayOfSortables = True
        
        Case Else
        
            isArrayOfSortables = False
    
    End Select
        
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


