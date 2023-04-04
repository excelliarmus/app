Attribute VB_Name = "Factory"
' function that creates a KNN classifier (Object) and returns it (requires k parameter)
Public Function CreateKNN(k As Integer) As clsKNN
    Dim knn_obj As clsKNN
    Set knn_obj = New clsKNN
    knn_obj.InitiateProperties k:=k
    Set CreateKNN = knn_obj
End Function

