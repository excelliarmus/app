VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_k As Double
Private p_X_train As Variant
Private p_y_train As Variant
Private p_predictions() As Double


Public Sub InitiateProperties(k As Integer)
    p_k = k
End Sub

Public Property Get k() As Variant
    k = p_k
End Property

Public Property Let k(ByVal new_k As Variant)
    p_k = new_k
End Property


Public Property Get X_train() As Variant
    X_train = p_X_train
End Property

Public Property Let X_train(ByVal new_X_train As Variant)
    p_X_train = new_X_train
End Property


Public Property Get y_train() As Variant
    y_train = p_y_train
End Property

Public Property Let y_train(ByVal new_y_train As Variant)
    p_y_train = new_y_train
End Property


Public Sub fit(X As Variant, y As Variant)
    p_X_train = X
    p_y_train = y
End Sub

Public Function predict(X() As Variant) As Integer


    Dim i As Integer
    Dim indices() As Long
    Dim lenx As Integer
    lenx = UBound(X)
    Dim distances() As Double
    ReDim Preserve distances(0 To UBound(p_X_train))
    Debug.Print "1.3"
    For i = 0 To UBound(p_X_train)
        distances(i) = euclidianDistance(X, p_X_train(i))
    Next i

'    For Each dist In distances
'        Debug.Print (dist)
'    Next dist
    

    ' Get the indices of the k smallest values
    indices = GetNSmallestIndices(distances, k)
    Dim nearest_labels() As Integer
    ReDim Preserve nearest_labels(0 To k - 1)
    ' Print the indices
    For i = LBound(indices) To UBound(indices)
        'Debug.Print (LBound(indices) & " " & UBound(indices))
        'Debug.Print ("The label of " & indices(i) & " is " & p_y_train(indices(i)))
        nearest_labels(i) = p_y_train(indices(i))
    Next i
    
    
'   Dim item As Variant
'   For Each item In nearest_labels
'      Debug.Print item
'   Next item
   
   Dim most_common As Integer
   most_common = GetMostCommonValue(nearest_labels)
   'Debug.Print ("Most common is " & most_common)
   
   predict = most_common

    
    
End Function


Private Function euclidianDistance(x1() As Variant, x2 As Variant) As Double
    Dim i As Integer
    Dim sum As Double
    sum = 0
    For i = 0 To UBound(x1)
        sum = sum + (x1(i) - x2(i)) ^ 2
    Next
    euclidianDistance = Sqr(sum)
    
End Function

Function GetNSmallestIndices(arr() As Double, n As Long) As Long()
    Dim i As Long, j As Long, min_idx As Long, temp As Double, temp_idx As Long
    Dim sorted_indices() As Long
    
    ' Initialize the sorted_indices array with the indices of the arr array
    ReDim sorted_indices(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        sorted_indices(i) = i
    Next i
    
    ' Sort the sorted_indices array by the values in the arr array
    For i = LBound(arr) To UBound(arr) - 1
        min_idx = i
        For j = i + 1 To UBound(arr)
            If arr(sorted_indices(j)) < arr(sorted_indices(min_idx)) Then
                min_idx = j
            End If
        Next j
        temp_idx = sorted_indices(i)
        sorted_indices(i) = sorted_indices(min_idx)
        sorted_indices(min_idx) = temp_idx
    Next i
    
    ' Return the n smallest values' indices
    ReDim Preserve sorted_indices(LBound(arr) To LBound(arr) + n - 1)
    GetNSmallestIndices = sorted_indices
End Function


Sub TestGetNSmallestIndices()
    Dim arr() As Double
    Dim indices() As Long
    Dim i As Long
    
    ' Initialize the arr array with some example values
    ReDim arr(1 To 10)
    For i = 1 To 10
        arr(i) = Rnd() * 10
    Next i
    
    ' Get the indices of the 3 smallest values
    indices = GetNSmallestIndices(arr, 3)
    
    ' Print the indices
    For i = LBound(indices) To UBound(indices)
        Debug.Print indices(i)
    Next i
End Sub

Function GetMostCommonValue(arr() As Integer) As Variant
    Dim count_dict As Object
    Dim i As Long
    Dim max_count As Long
    Dim most_common_value As String
    
    ' Create a dictionary to count the occurrences of each value in the arr array
    Set count_dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr) To UBound(arr)
        If count_dict.Exists(arr(i)) Then
            count_dict(arr(i)) = count_dict(arr(i)) + 1
        Else
            count_dict.Add arr(i), 1
        End If
    Next i
    
    ' Find the most common value in the arr array
    max_count = 0
    Dim key As Variant
    For Each key In count_dict.Keys
        If count_dict(key) > max_count Then
            max_count = count_dict(key)
            most_common_value = key
        End If
    Next key
    
    ' Return the most common value
    GetMostCommonValue = most_common_value
End Function
Sub TestGetMostCommonValue()
    Dim arr() As String
    Dim most_common_value As String
    Dim i As Long
    
    ' Initialize the arr array with some example values
    ReDim arr(1 To 10)
    arr(1) = "apple"
    arr(2) = "banana"
    arr(3) = "apple"
    arr(4) = "pear"
    arr(5) = "banana"
    arr(6) = "apple"
    arr(7) = "orange"
    arr(8) = "orange"
    arr(9) = "orange"
    arr(10) = "banana"
    
    ' Get the most common value in the arr array
    most_common_value = GetMostCommonValue(arr)
    
    ' Print the most common value
    Debug.Print most_common_value
End Sub

