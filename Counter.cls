VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public counter As Integer


Public Property Get getCounts() As Integer
    getCounts = counter
End Property

Public Property Let setCounts(ct As Integer)
    counter = ct
End Property

Sub increment()
    counter = counter + 1
End Sub

Sub reset()
    counter = 0
End Sub
