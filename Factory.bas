Attribute VB_Name = "Factory"
Public Function CreateKNN(k As Integer) As clsKNN
    Dim knn_obj As clsKNN
    Set knn_obj = New clsKNN
    knn_obj.InitiateProperties k:=k
    Set CreateKNN = knn_obj
End Function

