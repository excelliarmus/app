Attribute VB_Name = "ModIncrement"
' Global variables
Dim ct As counter 'number of time the button is clicked
Dim ct_created As Boolean 'boolean to check if the object was created


Function increment()
    If Not ct_created Then
        Set ct = New counter
        ct_created = True
    End If
    ct.increment
    increment = ct.getCounts
End Function

Sub reset()
    ct.reset
End Sub


