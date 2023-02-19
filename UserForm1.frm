VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   12600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22305
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub btnData_Click()

 'frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

 frmData.Visible = True

End Sub

Sub btnBalances_Click()

 frmData.Visible = False
 'frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

 frmBalances.Visible = True

End Sub





Sub btnTrading_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 'frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

 frmTrading.Visible = True

End Sub

Sub btnPrediction_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 'frmPrediction.Visible = False
 frmAbout.Visible = False

 frmPrediction.Visible = True

End Sub

Sub btnAbout_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 'frmAbout.Visible = False

 frmAbout.Visible = True

End Sub

Sub btnHome_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

End Sub

Sub btnExit_Click()

Unload Me

End Sub

Private Sub frmBalances_Click()

End Sub

Private Sub frmTrading_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub
