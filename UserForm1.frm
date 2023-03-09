VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Excelliarmus"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btnBalances1_Click()

Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)


End Sub

Private Sub btnBalancesUpdateBNB_Click()
Call ModBalances.UpdateBNB(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateBTC_Click()
Call ModBalances.UpdateBTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateBUSD_Click()
Call ModBalances.UpdateBUSD(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateETH_Click()
Call ModBalances.UpdateETH(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateLTC_Click()
Call ModBalances.UpdateLTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateTRX_Click()
Call ModBalances.UpdateTRX(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateUSDT_Click()
Call ModBalances.UpdateUSDT(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnBalancesUpdateXRP_Click()
Call ModBalances.UpdateXRP(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

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
Private Sub btnStartData1_Click()
    Call ModData.activateDataStream1
    Do Until Not ModData.get_isDataStream1On
        Call ModData.writeData1(UserForm1.inputData1)
        Call ModData.displayData1
        Call ModData.displayBidAsk1(UserForm1.inputData1)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop

    
End Sub

Private Sub btnStartData2_Click()
    Call ModData.activateDataStream2
    Do Until Not ModData.get_isDataStream2On
        Call ModData.displayData2(UserForm1.inputData2)
        Call ModData.displayBidAsk2(UserForm1.inputData2)
        Application.Wait (Now + TimeValue("00:00:01"))
        DoEvents
    Loop
End Sub


Private Sub btnStopData1_Click()
    Call ModData.desactivateDataStream1
End Sub

Private Sub btnStopData2_Click()
    Call ModData.desactivateDataStream2

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

MsgBox "Click on 'Delete' for all the next prompts"

Call ModDeleteAllCharts.deleteAll

Unload Me

End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub btnTrading1_Click()

Call ModTrading.buyBTCUSDT(UserForm1.inputBalances1, UserForm1.inputBalances2)

End Sub

Private Sub frmAbout_Click()

End Sub

Private Sub frmBalances_Click()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label20_Click()

End Sub

Private Sub Label30_Click()

End Sub

Sub UserForm_Activate()

Call ModMakeUserFormResizable.MakeFormResizable
Call ModData.initializeData(UserForm1.inputData1, UserForm1.inputData2)

End Sub
