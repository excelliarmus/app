VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpenOrders 
   Caption         =   "Open Orders"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   OleObjectBlob   =   "frmOpenOrders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpenOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UserForm_Activate()

Call ModMakeUserFormResizable.MakeFormResizable

End Sub

Private Sub UserForm_Initialize()
    With lblOpenOrders
        .ScrollBars = fmScrollBarsVertical
        .EnterFieldBehavior = fmEnterFieldBehaviorRecallSelection
    End With
End Sub

Private Sub lblOpenOrders_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With lblOpenOrders
        .SelStart = 0
        .SelLength = 0
    End With
End Sub
