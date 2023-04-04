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
' sub to update balances (requires API key and Secret Key)
Private Sub btnBalancesGetBalances_Click()
    Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to start streaming BNB
Private Sub btnBalancesStartBNB_Click()
    Call ModBalances.powerOnBNBStream
    Do Until Not ModBalances.isBNBStreamOn
        Call ModBalances.UpdateBNB(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to start streaming BTC
Private Sub btnBalancesStartBTC_Click()
    Call ModBalances.powerOnBTCStream
    Do Until Not ModBalances.isBTCStreamOn
        Call ModBalances.UpdateBTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop

End Sub

' sub to start streaming BUSD
Private Sub btnBalancesStartBUSD_Click()
    Call ModBalances.powerOnBUSDStream
    Do Until Not ModBalances.isBUSDStreamOn
        Call ModBalances.UpdateBUSD(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop

End Sub

' sub to start streaming ETH
Private Sub btnBalancesStartETH_Click()
    Call ModBalances.powerOnETHStream
    Do Until Not ModBalances.isETHStreamOn
        Call ModBalances.UpdateETH(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to start streaming all cryptos
Private Sub btnBalancesStartGlobal_Click()
    Call ModBalances.powerOnGlobalStream
    Do Until Not ModBalances.get_isGlobalStream1On
        Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to start streaming LTC
Private Sub btnBalancesStartLTC_Click()
    Call ModBalances.powerOnLTCStream
    Do Until Not ModBalances.isLTCStreamOn
        Call ModBalances.UpdateLTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to start streaming TRX
Private Sub btnBalancesStartTRX_Click()
    Call ModBalances.powerOnTRXStream
    Do Until Not ModBalances.isTRXStreamOn
        Call ModBalances.UpdateTRX(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to start streaming USDT
Private Sub btnBalancesStartUSDT_Click()
    Call ModBalances.powerOnUSDTStream
    Do Until Not ModBalances.isUSDTStreamOn
        Call ModBalances.UpdateUSDT(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to start streaming XRP
Private Sub btnBalancesStartXRP_Click()
    Call ModBalances.powerOnXRPStream
    Do Until Not ModBalances.isXRPStreamOn
        Call ModBalances.UpdateXRP(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub

' sub to stop streaming of BNB
Private Sub btnBalancesStopBNB_Click()
    Call ModBalances.powerOffBNBStream
End Sub

' sub to stop streaming of BTC
Private Sub btnBalancesStopBTC_Click()
    Call ModBalances.powerOffBTCStream
End Sub

' sub to stop streaming of BUSD
Private Sub btnBalancesStopBUSD_Click()
    Call ModBalances.powerOffBUSDStream
End Sub

' sub to stop streaming of ETH
Private Sub btnBalancesStopETH_Click()
    Call ModBalances.powerOffETHStream
End Sub

' sub to stop streaming of all cryptos
Private Sub btnBalancesStopGlobal_Click()
    Call ModBalances.powerOffGlobalStream
End Sub

' sub to stop streaming of LTC
Private Sub btnBalancesStopLTC_Click()
    Call ModBalances.powerOffLTCStream
End Sub

' sub to stop streaming of TRX
Private Sub btnBalancesStopTRX_Click()
    Call ModBalances.powerOffTRXStream
End Sub

' sub to stop streaming of USDT
Private Sub btnBalancesStopUSDT_Click()
    Call ModBalances.powerOffUSDTStream
End Sub

' sub to stop streaming of XRP
Private Sub btnBalancesStopXRP_Click()
    Call ModBalances.powerOffXRPStream
End Sub

' sub to update balance of BNB
Private Sub btnBalancesUpdateBNB_Click()
    Call ModBalances.UpdateBNB(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of BTC
Private Sub btnBalancesUpdateBTC_Click()
    Call ModBalances.UpdateBTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of BUSD
Private Sub btnBalancesUpdateBUSD_Click()
    Call ModBalances.UpdateBUSD(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of ETH
Private Sub btnBalancesUpdateETH_Click()
    Call ModBalances.UpdateETH(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of LTC
Private Sub btnBalancesUpdateLTC_Click()
    Call ModBalances.UpdateLTC(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of TRX
Private Sub btnBalancesUpdateTRX_Click()
    Call ModBalances.UpdateTRX(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of USDT
Private Sub btnBalancesUpdateUSDT_Click()
    Call ModBalances.UpdateUSDT(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to update balance of XRP
Private Sub btnBalancesUpdateXRP_Click()
    Call ModBalances.UpdateXRP(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to display Data frame
Sub btnData_Click()
    frmBalances.Visible = False
    frmTrading.Visible = False
    frmPrediction.Visible = False
    frmAbout.Visible = False
    frmAccount.Visible = False
    frmData.Visible = True
End Sub

' sub to display Balances frame
Sub btnBalances_Click()
    frmData.Visible = False
    frmTrading.Visible = False
    frmPrediction.Visible = False
    frmAbout.Visible = False
    frmAccount.Visible = False
    frmBalances.Visible = True
End Sub

' sub to start ML bot (requires ticker, quantity, k, nsticks, frequency and discrimination rate)
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
    ' call sub of test prediction for explications
    ' Call modPrediction.testKNN
End Sub

' sub to stop ML bot
Private Sub btnPredictionStop_Click()
    Call modPrediction.desactivateMLBot
    Call ModBalances.powerOffGlobalStream
End Sub

' sub to start candlesticks' chart
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

' sub to start line chart
Private Sub btnStartData2_Click()
    Call ModData.activateDataStream2
    Do Until Not ModData.get_isDataStream2On
        Call ModData.displayData2(UserForm1.inputData2)
        Call ModData.displayBidAsk2(UserForm1.inputData2)
        Application.Wait (Now + TimeValue("00:00:01"))
        DoEvents
    Loop
End Sub

' sub to stop candlesticks' chart
Private Sub btnStopData1_Click()
    Call ModData.desactivateDataStream1
End Sub

' sub to stop line chart
Private Sub btnStopData2_Click()
    Call ModData.desactivateDataStream2
End Sub

' sub to display Trading frame
Sub btnTrading_Click()
    frmData.Visible = False
    frmBalances.Visible = False
    frmPrediction.Visible = False
    frmAbout.Visible = False
    frmAccount.Visible = False
    frmTrading.Visible = True
End Sub

' sub to display Prediction frame
Sub btnPrediction_Click()
    frmData.Visible = False
    frmBalances.Visible = False
    frmTrading.Visible = False
    frmAbout.Visible = False
    frmAccount.Visible = False
    frmPrediction.Visible = True
End Sub

' sub to display About frame
Sub btnAbout_Click()
    frmData.Visible = False
    frmBalances.Visible = False
    frmTrading.Visible = False
    frmPrediction.Visible = False
    frmAccount.Visible = False
    frmAbout.Visible = True
End Sub

' sub to display Account frame
Sub btnAccount_Click()
    frmData.Visible = False
    frmBalances.Visible = False
    frmTrading.Visible = False
    frmPrediction.Visible = False
    frmAbout.Visible = False
    frmAccount.Visible = True
End Sub

' sub to hide all frames (go to home, displays background)
Sub btnHome_Click()
    frmData.Visible = False
    frmBalances.Visible = False
    frmTrading.Visible = False
    frmPrediction.Visible = False
    frmAbout.Visible = False
    frmAccount.Visible = False
End Sub

' sub to exit app
Sub btnExit_Click()
    MsgBox "Click on 'Delete' for all the next prompts"
    Call ModDeleteAllCharts.deleteAll
    Unload Me
End Sub

' sub to display all orders (requires API key and Secret key)
Private Sub btnTradingDisplayAllOrders_Click()
    Call ModTrading.getAllOrders(UserForm1.inputBalances1, UserForm1.inputBalances2)
    frmAllOrders.Show
End Sub

' sub to display open orders (requires API key and Secret key)
Private Sub btnTradingDisplayOrders_Click()
    Call ModTrading.getOpenOrders(UserForm1.inputBalances1, UserForm1.inputBalances2)
    frmOpenOrders.Show
End Sub

' sub to place an order (requires API key and Secret key)
Private Sub btnTradingPlaceOrder_Click()
    Call ModTrading.placeOrder(UserForm1.inputBalances1, UserForm1.inputBalances2)
End Sub

' sub to start one the 3 trading bots
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

' sub to stop the currently running trading bot (stops all bots)
Private Sub btnTradingStopBot_Click()
    Call ModTrading.powerOffRandomTradingBot
    Call ModTrading.powerOffMRTradingBot
    Call ModTrading.powerOffMomentumTradingBot
    Call ModBalances.powerOffGlobalStream
End Sub

' sub to send an e-mail
Private Sub lblAboutContact_Click()
    ActiveWorkbook.FollowHyperlink Address:="mailto:excelliarmus@proton.me", NewWindow:=True
End Sub

' sub to redirect to the docs
Private Sub lblAboutDocs_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://excelliarmus.gitbook.io/docs/", NewWindow:=True
End Sub

' sub to redirect to the github repo
Private Sub lblAboutGithub_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://github.com/excelliarmus/app", NewWindow:=True
End Sub

' sub to redirect to the YouTube channel
Private Sub lblAboutYoutube_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.youtube.com/@Excelliarmus", NewWindow:=True
End Sub

' sub to display Limit order selected
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

' sub to display Market order selected
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

' sub to display Momentum trading bot selected
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

' sub to display Mean-reversion trading bot selected
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

' sub to display Random trading bot selected
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

' sub to display Stop Loss order selected
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

' sub to make to userform resizable and to display charts 1 and 2
Sub UserForm_Activate()
    Call ModMakeUserFormResizable.MakeFormResizable
    Call ModData.initializeData(UserForm1.inputData1, UserForm1.inputData2)
End Sub
