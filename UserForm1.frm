VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22515
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCheck_Click()
    Label3.Caption = ModIncrement.increment
End Sub

Private Sub btnReset_Click()
    ModIncrement.reset
    Label3.Caption = 0

End Sub

Private Sub frm1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Click()

End Sub
