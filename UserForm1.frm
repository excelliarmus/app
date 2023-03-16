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




Private Sub btnBalancesGetBalances_Click()

Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)


End Sub

Private Sub btnBalancesStartBNB_Click()
    Call ModBalances.powerOnBNBStream
    Do Until Not ModBalances.isBNBStreamOn
        Call ModBalances.UpdateBNB(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStartBTC_Click()
    Call ModBalances.powerOnBTCStream
    Do Until Not ModBalances.isBTCStreamOn
        Call ModBalances.UpdateBTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop

End Sub

Private Sub btnBalancesStartBUSD_Click()
    Call ModBalances.powerOnBUSDStream
    Do Until Not ModBalances.isBUSDStreamOn
        Call ModBalances.UpdateBUSD(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop

End Sub

Private Sub btnBalancesStartETH_Click()
    Call ModBalances.powerOnETHStream
    Do Until Not ModBalances.isETHStreamOn
        Call ModBalances.UpdateETH(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStartGlobal_Click()
    Call ModBalances.powerOnGlobalStream
    Do Until Not ModBalances.get_isGlobalStream1On
        Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStartLTC_Click()
    Call ModBalances.powerOnLTCStream
    Do Until Not ModBalances.isLTCStreamOn
        Call ModBalances.UpdateLTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStartTRX_Click()
    Call ModBalances.powerOnTRXStream
    Do Until Not ModBalances.isTRXStreamOn
        Call ModBalances.UpdateTRX(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStartUSDT_Click()
    Call ModBalances.powerOnUSDTStream
    Do Until Not ModBalances.isUSDTStreamOn
        Call ModBalances.UpdateUSDT(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStartXRP_Click()
    Call ModBalances.powerOnXRPStream
    Do Until Not ModBalances.isXRPStreamOn
        Call ModBalances.UpdateXRP(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

Private Sub btnBalancesStopBNB_Click()
    Call ModBalances.powerOffBNBStream
End Sub

Private Sub btnBalancesStopBTC_Click()
    Call ModBalances.powerOffBTCStream
End Sub

Private Sub btnBalancesStopBUSD_Click()
    Call ModBalances.powerOffBUSDStream
End Sub

Private Sub btnBalancesStopETH_Click()
    Call ModBalances.powerOffETHStream
End Sub

Private Sub btnBalancesStopGlobal_Click()
    Call ModBalances.powerOffGlobalStream
End Sub

Private Sub btnBalancesStopLTC_Click()
    Call ModBalances.powerOffLTCStream
End Sub

Private Sub btnBalancesStopTRX_Click()
    Call ModBalances.powerOffTRXStream
End Sub

Private Sub btnBalancesStopUSDT_Click()
    Call ModBalances.powerOffUSDTStream
End Sub

Private Sub btnBalancesStopXRP_Click()
    Call ModBalances.powerOffXRPStream
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


Private Sub btnTradingDisplayAllOrders_Click()
    Call ModTrading.getAllOrders(UserForm1.inputBalances1, UserForm1.inputBalances2)
    frmAllOrders.Show
End Sub

Private Sub btnTradingDisplayOrders_Click()
    Call ModTrading.getOpenOrders(UserForm1.inputBalances1, UserForm1.inputBalances2)
    frmOpenOrders.Show
End Sub

Private Sub btnTradingPlaceOrder_Click()
    Call ModTrading.placeOrder(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

Private Sub btnTradingStartBot_Click()
    Call ModBalances.powerOnGlobalStream
    If UserForm1.tglTradingRandomBot.Value = True Then
        Call ModTrading.powerOnRandomTradingBot
        Call ModTrading.runRandomBot
    ElseIf UserForm1.tglTradingMRBot.Value = True Then
        Call ModTrading.powerOnMRTradingBot
        Call ModTrading.runMRBot
    ElseIf UserForm1.tglTradingMomentumBot.Value = True Then
        Call ModTrading.powerOnMomentumTradingBot
        Call ModTrading.runMomentumBot
    End If
    
End Sub

Private Sub btnTradingStopBot_Click()
    Call ModTrading.powerOffRandomTradingBot
    Call ModTrading.powerOffMRTradingBot
    Call ModTrading.powerOffMomentumTradingBot
    Call ModBalances.powerOffGlobalStream
End Sub

Private Sub frmAbout_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label20_Click()

End Sub

Private Sub Label60_Click()

End Sub

Private Sub Label61_Click()

End Sub

Private Sub lblAboutEmail_Click()
ActiveWorkbook.FollowHyperlink Address:="mailto:excelliarmus@proton.me", NewWindow:=True
End Sub

Private Sub lblAboutRepo_Click()
ActiveWorkbook.FollowHyperlink Address:="https://github.com/excelliarmus/app", NewWindow:=True

End Sub

Private Sub tglTradingRandomBot_Click()

End Sub

Sub UserForm_Activate()

Call ModMakeUserFormResizable.MakeFormResizable
Call ModData.initializeData(UserForm1.inputData1, UserForm1.inputData2)

End Sub
