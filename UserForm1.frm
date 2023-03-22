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
Public isMarketToggled As Boolean
Public isLimitToggled As Boolean
Public isSLToggled As Boolean
Public isRandomToggled As Boolean
Public isMRToggled As Boolean
Public isMomentumToggled As Boolean

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
 frmAccount.Visible = False

 frmData.Visible = True

End Sub

Sub btnBalances_Click()

 frmData.Visible = False
 'frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False
 frmAccount.Visible = False

 frmBalances.Visible = True

End Sub

Private Sub btnExit_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub btnHome_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub btnPredictionStart_Click()
    Dim ticker As String
    Dim qt As String
    Dim k As Integer
    Dim nsticks As Integer
    Dim freq As Integer
    Dim disc As Double
    ticker = UserForm1.inputPredictionTicker
    qt = UserForm1.inputPredictionQuantity
    k = CInt(UserForm1.inputPredictionK)
    nsticks = CInt(UserForm1.inputPredictionNumberSticks)
    freq = CInt(UserForm1.inputPredictionFrequency)
    disc = CDbl(Replace(UserForm1.inputPredictionDiscrimination, ".", ","))
    
    Call ModBalances.powerOnGlobalStream
    Call modPrediction.startBot(ticker, qt, k, nsticks, freq, disc)


End Sub

Private Sub btnPredictionStop_Click()
    Call modPrediction.desactivateMLBot
    Call ModBalances.powerOffGlobalStream
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
 frmAccount.Visible = False

 frmTrading.Visible = True

End Sub

Sub btnPrediction_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 'frmPrediction.Visible = False
 frmAbout.Visible = False
 frmAccount.Visible = False

 frmPrediction.Visible = True

End Sub

Sub btnAbout_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 'frmAbout.Visible = False
 frmAccount.Visible = False

 frmAbout.Visible = True

End Sub

Sub btnAccount_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 'frmAbout.Visible = False
 frmAbout.Visible = False
 
 
 
 frmAccount.Visible = True

 

End Sub

Sub btnHome_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False
 frmAccount.Visible = False

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

Private Sub CommandButton1_Click()
    Call modPrediction.predict(10, 2)
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

Private Sub Label92_Click()

End Sub

Private Sub Label93_Click()

End Sub

Private Sub lblAboutContact_Click()
ActiveWorkbook.FollowHyperlink Address:="mailto:excelliarmus@proton.me", NewWindow:=True
End Sub


Private Sub lblAboutDocs_Click()
ActiveWorkbook.FollowHyperlink Address:="https://excelliarmus.gitbook.io/docs/", NewWindow:=True
End Sub

Private Sub lblAboutGithub_Click()
ActiveWorkbook.FollowHyperlink Address:="https://github.com/excelliarmus/app", NewWindow:=True
End Sub


Private Sub lblAboutYoutube_Click()
ActiveWorkbook.FollowHyperlink Address:="https://www.youtube.com/@Excelliarmus", NewWindow:=True
End Sub

Private Sub lblAboutEmail_Click()

End Sub

Private Sub tglTradingLimit_Click()
    If isLimitToggled Then
        isLimitToggled = False
        UserForm1.tglTradingLimit.BackStyle = fmBackStyleTransparent
        UserForm1.tglTradingLimit.ForeColor = &HFFFFFF
    Else
        isLimitToggled = True
        UserForm1.tglTradingLimit.BackStyle = fmBackStyleOpaque
        UserForm1.tglTradingLimit.ForeColor = &H0&
    End If
End Sub

Private Sub tglTradingMarket_Click()
    If isMarketToggled Then
        isMarketToggled = False
        UserForm1.tglTradingMarket.BackStyle = fmBackStyleTransparent
        UserForm1.tglTradingMarket.ForeColor = &HFFFFFF
    Else
        isMarketToggled = True
        UserForm1.tglTradingMarket.BackStyle = fmBackStyleOpaque
        UserForm1.tglTradingMarket.ForeColor = &H0&
    End If
    
    
End Sub

Private Sub tglTradingMomentumBot_Click()
    If isMomentumToggled Then
        isMomentumToggled = False
        UserForm1.tglTradingMomentumBot.BackStyle = fmBackStyleTransparent
        UserForm1.tglTradingMomentumBot.ForeColor = &HFFFFFF
    Else
        isMomentumToggled = True
        UserForm1.tglTradingMomentumBot.BackStyle = fmBackStyleOpaque
        UserForm1.tglTradingMomentumBot.ForeColor = &H0&
    End If
End Sub

Private Sub tglTradingMRBot_Click()
    If isMRToggled Then
        isMRToggled = False
        UserForm1.tglTradingMRBot.BackStyle = fmBackStyleTransparent
        UserForm1.tglTradingMRBot.ForeColor = &HFFFFFF
    Else
        isMRToggled = True
        UserForm1.tglTradingMRBot.BackStyle = fmBackStyleOpaque
        UserForm1.tglTradingMRBot.ForeColor = &H0&
    End If
End Sub

Private Sub tglTradingRandomBot_Click()
    If isRandomToggled Then
        isRandomToggled = False
        UserForm1.tglTradingRandomBot.BackStyle = fmBackStyleTransparent
        UserForm1.tglTradingRandomBot.ForeColor = &HFFFFFF
    Else
        isRandomToggled = True
        UserForm1.tglTradingRandomBot.BackStyle = fmBackStyleOpaque
        UserForm1.tglTradingRandomBot.ForeColor = &H0&
    End If
End Sub

Private Sub tglTradingSL_Click()
    If isSLToggled Then
        isSLToggled = False
        UserForm1.tglTradingSL.BackStyle = fmBackStyleTransparent
        UserForm1.tglTradingSL.ForeColor = &HFFFFFF
    Else
        isSLToggled = True
        UserForm1.tglTradingSL.BackStyle = fmBackStyleOpaque
        UserForm1.tglTradingSL.ForeColor = &H0&
    End If
End Sub

Sub UserForm_Activate()

Call ModMakeUserFormResizable.MakeFormResizable
Call ModData.initializeData(UserForm1.inputData1, UserForm1.inputData2)

End Sub
