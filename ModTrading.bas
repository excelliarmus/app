Attribute VB_Name = "ModTrading"
Public logsArray(12) As String
Public isRandomBotOn As Boolean
Public isMRBotOn As Boolean
Public isMomentumBotOn As Boolean

Function placeMarketOrder(APIkey As String, secret_key As String, side As String, ticker As String, qt As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    On Error GoTo error
    If side = "BUY" Then
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=BUY&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp, secret_key)
        url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=BUY&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp & "&signature=" & signature
    Else
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=SELL&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp, secret_key)
        url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=SELL&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp & "&signature=" & signature
    End If
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    'On Error GoTo error_apikey
    'MsgBox xmlhttp.responseText
    If json("symbol") = ticker Then
        placeMarketOrder = "success"
    Else:
        placeMarketOrder = xmlhttp.responseText
    End If
    
    
done:
    Exit Function
error:
        MsgBox "An error occured : " & xmlhttp.responseText
End Function

Function placeLimitOrder(APIkey As String, secret_key As String, side As String, ticker As String, qt As String, limit_price As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    On Error GoTo error
    If side = "BUY" Then
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=BUY&type=LIMIT&quantity=" & qt & "&price=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp, secret_key)
        url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=BUY&type=LIMIT&quantity=" & qt & "&price=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp & "&signature=" & signature
    Else
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=SELL&type=LIMIT&quantity=" & qt & "&price=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp, secret_key)
        url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=SELL&type=LIMIT&quantity=" & qt & "&price=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp & "&signature=" & signature
    End If
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    'On Error GoTo error_apikey
    'MsgBox xmlhttp.responseText
    If json("symbol") = ticker Then
        placeLimitOrder = "success"
    Else:
        placeLimitOrder = xmlhttp.responseText
    End If
    
    
done:
    Exit Function
error:
        MsgBox "An error occured : " & xmlhttp.responseText
End Function

Function placeSLOrder(APIkey As String, secret_key As String, side As String, ticker As String, qt As String, limit_price As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    On Error GoTo error
    If side = "BUY" Then
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=BUY&type=STOP_LOSS_LIMIT&quantity=" & qt & "&price=" & limit_price & "&stopPrice=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp, secret_key)
        url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=BUY&type=STOP_LOSS_LIMIT&quantity=" & qt & "&price=" & limit_price & "&stopPrice=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp & "&signature=" & signature
    Else
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=SELL&type=STOP_LOSS_LIMIT&quantity=" & qt & "&price=" & limit_price & "&stopPrice=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp, secret_key)
        url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=SELL&type=STOP_LOSS_LIMIT&quantity=" & qt & "&price=" & limit_price & "&stopPrice=" & limit_price & "&timeInForce=GTC&timestamp=" & timestamp & "&signature=" & signature
    End If
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    'On Error GoTo error_apikey
    'MsgBox xmlhttp.responseText
    If json("symbol") = ticker Then
        placeSLOrder = "success"
    Else:
        placeSLOrder = xmlhttp.responseText
    End If
    
    
done:
    Exit Function
error:
        MsgBox "An error occured : " & xmlhttp.responseText
End Function

Sub placeOrder(APIkey As String, secret_key As String)
    Dim ticker As String
    Dim qt As String
    Dim signal As String
    Dim limit_price As String
    
    
    ticker = UserForm1.inputTradingTicker
    qt = UserForm1.inputTradingQuantity
    limit_price = UserForm1.inputTradingLimitPrice
    

    
    If UserForm1.tglTradingMarket.Value = True Then
        If UserForm1.optTradingBuy.Value = True Then
            'BUY MARKET
            signal = placeMarketOrder(APIkey, secret_key, "BUY", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9650) & " BUY " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        ElseIf UserForm1.optTradingSell.Value = True Then
            'SELL MARKET
            signal = placeMarketOrder(APIkey, secret_key, "SELL", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9660) & " SELL " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        End If
    ElseIf UserForm1.tglTradingLimit.Value = True Then
        If UserForm1.optTradingBuy.Value = True Then
            'BUY LIMIT
            signal = placeLimitOrder(APIkey, secret_key, "BUY", ticker, qt, limit_price)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9650) & " BUY LIMIT " & qt & " " & ticker & " @ " & limit_price)
            Else
                addLog (signal)
            End If
        ElseIf UserForm1.optTradingSell.Value = True Then
            'SELL LIMIT
            signal = placeLimitOrder(APIkey, secret_key, "SELL", ticker, qt, limit_price)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9660) & " SELL LIMIT " & qt & " " & ticker & " @ " & limit_price)
            Else
                addLog (signal)
            End If
        End If
    ElseIf UserForm1.tglTradingSL.Value = True Then
        If UserForm1.optTradingBuy.Value = True Then
            'BUY STOP LOSS
            signal = placeSLOrder(APIkey, secret_key, "BUY", ticker, qt, limit_price)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9650) & " BUY STOP LOSS " & qt & " " & ticker & " @ " & limit_price)
            Else
                addLog (signal)
            End If
        ElseIf UserForm1.optTradingSell.Value = True Then
            'SELL STOP LOSS
            signal = placeSLOrder(APIkey, secret_key, "SELL", ticker, qt, limit_price)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9660) & " SELL STOP LOSS " & qt & " " & ticker & " @ " & limit_price)
            Else
                addLog (signal)
            End If
        End If
    
    
    End If


End Sub

Sub displayLogs()
    UserForm1.lblTradingLog1 = logsArray(0)
    If logsArray(0) Like "*BUY*" Then
        UserForm1.lblTradingLog1.ForeColor = &HFF00&
    ElseIf logsArray(0) Like "*SELL*" Then
        UserForm1.lblTradingLog1.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog1.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog2 = logsArray(1)
    If logsArray(1) Like "*BUY*" Then
        UserForm1.lblTradingLog2.ForeColor = &HFF00&
    ElseIf logsArray(1) Like "*SELL*" Then
        UserForm1.lblTradingLog2.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog2.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog3 = logsArray(2)
    If logsArray(2) Like "*BUY*" Then
        UserForm1.lblTradingLog3.ForeColor = &HFF00&
    ElseIf logsArray(2) Like "*SELL*" Then
        UserForm1.lblTradingLog3.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog3.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog4 = logsArray(3)
    If logsArray(3) Like "*BUY*" Then
        UserForm1.lblTradingLog4.ForeColor = &HFF00&
    ElseIf logsArray(3) Like "*SELL*" Then
        UserForm1.lblTradingLog4.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog4.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog5 = logsArray(4)
    If logsArray(4) Like "*BUY*" Then
        UserForm1.lblTradingLog5.ForeColor = &HFF00&
    ElseIf logsArray(4) Like "*SELL*" Then
        UserForm1.lblTradingLog5.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog5.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog6 = logsArray(5)
    If logsArray(5) Like "*BUY*" Then
        UserForm1.lblTradingLog6.ForeColor = &HFF00&
    ElseIf logsArray(5) Like "*SELL*" Then
        UserForm1.lblTradingLog6.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog6.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog6 = logsArray(5)
    If logsArray(5) Like "*BUY*" Then
        UserForm1.lblTradingLog6.ForeColor = &HFF00&
    ElseIf logsArray(5) Like "*SELL*" Then
        UserForm1.lblTradingLog6.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog6.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog7 = logsArray(6)
    If logsArray(6) Like "*BUY*" Then
        UserForm1.lblTradingLog7.ForeColor = &HFF00&
    ElseIf logsArray(6) Like "*SELL*" Then
        UserForm1.lblTradingLog7.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog7.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog8 = logsArray(7)
    If logsArray(7) Like "*BUY*" Then
        UserForm1.lblTradingLog8.ForeColor = &HFF00&
    ElseIf logsArray(7) Like "*SELL*" Then
        UserForm1.lblTradingLog8.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog8.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog9 = logsArray(8)
    If logsArray(8) Like "*BUY*" Then
        UserForm1.lblTradingLog9.ForeColor = &HFF00&
    ElseIf logsArray(8) Like "*SELL*" Then
        UserForm1.lblTradingLog9.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog9.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog10 = logsArray(9)
    If logsArray(9) Like "*BUY*" Then
        UserForm1.lblTradingLog10.ForeColor = &HFF00&
    ElseIf logsArray(9) Like "*SELL*" Then
        UserForm1.lblTradingLog10.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog10.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog11 = logsArray(10)
    If logsArray(10) Like "*BUY*" Then
        UserForm1.lblTradingLog11.ForeColor = &HFF00&
    ElseIf logsArray(10) Like "*SELL*" Then
        UserForm1.lblTradingLog11.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog11.ForeColor = &H80FF&
    End If
    
    UserForm1.lblTradingLog12 = logsArray(11)
    If logsArray(11) Like "*BUY*" Then
        UserForm1.lblTradingLog12.ForeColor = &HFF00&
    ElseIf logsArray(11) Like "*SELL*" Then
        UserForm1.lblTradingLog12.ForeColor = &HFF&
    Else
        UserForm1.lblTradingLog12.ForeColor = &H80FF&
    End If
    
    'UserForm1.lblTradingLog2 = logsArray(1)
    'UserForm1.lblTradingLog3 = logsArray(2)
    'UserForm1.lblTradingLog4 = logsArray(3)
    'UserForm1.lblTradingLog5 = logsArray(4)
    'UserForm1.lblTradingLog6 = logsArray(5)
    'UserForm1.lblTradingLog7 = logsArray(6)
    'UserForm1.lblTradingLog8 = logsArray(7)
    'UserForm1.lblTradingLog9 = logsArray(8)
    'UserForm1.lblTradingLog10 = logsArray(9)
    'UserForm1.lblTradingLog11 = logsArray(10)
    'UserForm1.lblTradingLog12 = logsArray(11)
End Sub

Sub addLog(log As String)
    len_arr = UBound(logsArray) - LBound(logsArray)
    For i = len_arr - 1 To 1 Step -1
    ' Debug.Print (i)
        logsArray(i) = logsArray(i - 1)
    Next
    logsArray(LBound(logsArray)) = log
    displayLogs
End Sub

Sub DEBUGdisplayLogsArray()
len_arr = UBound(logsArray) - LBound(logsArray)
Debug.Print ("------------------------------")
    For i = 0 To len_arr
        Debug.Print ("logsArray(" & i & ") = " & logsArray(i))
    Next
End Sub

Sub runRandomBot()
    Dim ticker As String
    Dim qt As String
    Dim signal As String
    Dim frequence As String

    
    ticker = UserForm1.inputTradingTicker
    qt = UserForm1.inputTradingQuantity
    frequence = UserForm1.inputTradingFrequence
    
    Do Until Not isRandomBotOn
        Dim rand As Integer
        rand = Int(3 * Rnd) + 1
        If rand = 1 Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "BUY", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9650) & " BUY " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        ElseIf rand = 2 Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "SELL", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9660) & " SELL " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        Else
            addLog (get_time_for_logs & " : Do nothing " & ChrW(9787))
        End If
        Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)

        Application.Wait (Now + TimeValue("00:00:" & frequence))
        DoEvents
    Loop
End Sub

Sub runMRBot()
    Dim ticker As String
    Dim qt As String
    Dim signal As String
    Dim frequence As String
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    Dim average As Double
    Dim closeArray(5) As Double
    Dim current_price As Double
    Dim upLimit As Double
    Dim downLimit As Double
    Dim margin As Double
    
    ticker = UserForm1.inputTradingTicker
    qt = UserForm1.inputTradingQuantity
    frequence = UserForm1.inputTradingFrequence
    margin = CDbl(Replace(UserForm1.inputTradingMargin, ".", ","))

'On Error GoTo noticker



url = "https://api.binance.com/api/v3/klines?symbol=" & ticker & "&interval=1m&limit=5"
xmlhttp.Open "GET", url, False
xmlhttp.Send
Set json = JsonConverter.ParseJson(xmlhttp.responseText)

'[
'  [
'    1499040000000,      // Kline open time (1)
'    "0.01634790",       // Open price (2)
'    "0.80000000",       // High price (3)
'    "0.01575800",       // Low price (4)
'    "0.01577100",       // Close price (5)
'    "148976.11427815",  // Volume (6)
'    1499644799999,      // Kline Close time (7)
'    "2434.19055334",    // Quote asset volume (8)
'    308,                // Number of trades (9)
'    "1756.87402397",    // Taker buy base asset volume (10)
'    "28.46694368",      // Taker buy quote asset volume (11)
'    "0"                 // Unused field, ignore. (12)
'  ],
'  ...
'  [...]
']


Dim i As Integer


For i = 1 To 5
    closeArray(i) = CDbl(Replace(json(i)(5), ".", ","))
Next

average = Application.WorksheetFunction.sum(closeArray) / 5
current_price = ModData.getCurrentPrice(ticker)
upLimit = average + (average * (1 + margin))
downLimit = average + (average * (1 - margin))


    Do Until Not isMRBotOn

        If current_price < downLimit Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "BUY", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9650) & " BUY " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker) & " (down limit is " & downLimit & " )")
            Else
                addLog (signal)
            End If
        ElseIf current_price > upLimit Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "SELL", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9660) & " SELL " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker) & " (upper limit is " & upLimit & " )")
            Else
                addLog (signal)
            End If
        Else
            addLog (get_time_for_logs & " : Do nothing " & ChrW(9787))
        End If
        Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)

        Application.Wait (Now + TimeValue("00:00:" & frequence))
        DoEvents
    Loop
End Sub


Sub runMomentumBot()
    Dim ticker As String
    Dim qt As String
    Dim signal As String
    Dim frequence As String
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    Dim average As Double
    Dim closeArray(5) As Double
    Dim current_price As Double
    Dim upLimit As Double
    Dim downLimit As Double
    Dim margin As Double
    
    ticker = UserForm1.inputTradingTicker
    qt = UserForm1.inputTradingQuantity
    frequence = UserForm1.inputTradingFrequence
    margin = CDbl(Replace(UserForm1.inputTradingMargin, ".", ","))

'On Error GoTo noticker
url = "https://api.binance.com/api/v3/klines?symbol=" & ticker & "&interval=1m&limit=5"
xmlhttp.Open "GET", url, False
xmlhttp.Send
Set json = JsonConverter.ParseJson(xmlhttp.responseText)
Dim i As Integer


For i = 1 To 5
    closeArray(i) = CDbl(Replace(json(i)(5), ".", ","))
Next

average = Application.WorksheetFunction.sum(closeArray) / 5
current_price = ModData.getCurrentPrice(ticker)
upLimit = average + (average * (1 + margin))
downLimit = average + (average * (1 - margin))


    Do Until Not isMomentumBotOn

        If current_price > upLimit Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "BUY", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9650) & " BUY " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker) & " (upper limit is " & upLimit & " )")
            Else
                addLog (signal)
            End If
        ElseIf current_price < downLimit Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "SELL", ticker, qt)
            If signal = "success" Then
                addLog (get_time_for_logs & " : " & ChrW(9660) & " SELL " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker) & " (down limit is " & downLimit & " )")
            Else
                addLog (signal)
            End If
        Else
            addLog (get_time_for_logs & " : Do nothing " & ChrW(9787))
        End If
        Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)

        Application.Wait (Now + TimeValue("00:00:" & frequence))
        DoEvents
    Loop
End Sub

Sub powerOnRandomTradingBot()
    isRandomBotOn = True
    UserForm1.lblTradingBotStatus.BorderColor = &HFF00&
    UserForm1.lblTradingBotStatus.Caption = "ON"
    UserForm1.lblTradingBotStatus.ForeColor = &HFF00&

End Sub

Sub powerOffRandomTradingBot()
    isRandomBotOn = False
    UserForm1.lblTradingBotStatus.BorderColor = &HFF&
    UserForm1.lblTradingBotStatus.Caption = "OFF"
    UserForm1.lblTradingBotStatus.ForeColor = &HFF&
End Sub

Sub powerOnMRTradingBot()
    isMRBotOn = True
    UserForm1.lblTradingBotStatus.BorderColor = &HFF00&
    UserForm1.lblTradingBotStatus.Caption = "ON"
    UserForm1.lblTradingBotStatus.ForeColor = &HFF00&

End Sub

Sub powerOffMRTradingBot()
    isMRBotOn = False
    UserForm1.lblTradingBotStatus.BorderColor = &HFF&
    UserForm1.lblTradingBotStatus.Caption = "OFF"
    UserForm1.lblTradingBotStatus.ForeColor = &HFF&
End Sub

Sub powerOnMomentumTradingBot()
    isMomentumBotOn = True
    UserForm1.lblTradingBotStatus.BorderColor = &HFF00&
    UserForm1.lblTradingBotStatus.Caption = "ON"
    UserForm1.lblTradingBotStatus.ForeColor = &HFF00&

End Sub

Sub powerOffMomentumTradingBot()
    isMomentumBotOn = False
    UserForm1.lblTradingBotStatus.BorderColor = &HFF&
    UserForm1.lblTradingBotStatus.Caption = "OFF"
    UserForm1.lblTradingBotStatus.ForeColor = &HFF&
End Sub

Function get_time_for_logs()
    Dim date_now As String
    date_now = Format(Now(), "hh:mm:ss")
    get_time_for_logs = date_now
End Function

Sub getOpenOrders(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    Dim ordersString As String

    On Error GoTo error
    
    
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    
    url = "https://testnet.binance.vision/api/v3/openOrders?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    
    
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
'    [
'        {
'    "symbol": "BTCUSDT",
'    "orderId": 11118310,
'    "orderListId": -1,
'    "clientOrderId": "0E8pfX91ZoJAR9qr7PwLVR",
'    "price": "25000.00000000",
'    "origQty": "0.00100000",
'    "executedQty": "0.00000000",
'    "cummulativeQuoteQty": "0.00000000",
'    "status": "NEW",
'    "timeInForce": "GTC",
'    "type": "LIMIT",
'    "side": "SELL",
'    "stopPrice": "0.00000000",
'    "icebergQty": "0.00000000",
'    "time": 1678897799099,
'    "updateTime": 1678897799099,
'    "isWorking": true,
'    "workingTime": 1678897799099,
'    "origQuoteOrderQty": "0.00000000",
'    "selfTradePreventionMode": "NONE"
'        }
'    ]
    'frmOpenOrders.lblOpenOrders.Caption = xmlhttp.responseText
    'MsgBox xmlhttp.responseText
        For Each item In json
            ordersString = ordersString & ChrW(8658) & " "
            For Each Child In item
                ordersString = ordersString & "[ " & Child & " : " & item(Child) & " ] "
                ' Debug.Print Child & " : " & Item(Child) & vbNewLine & "Line2"
            Next Child
            ordersString = ordersString & vbNewLine & vbNewLine
        Next item
    'Debug.Print ordersString
    frmOpenOrders.lblOpenOrders.Text = ordersString
    
done:
    Exit Sub
error:
        MsgBox "An error occured : " & xmlhttp.responseText
End Sub


Sub getAllOrders(APIkey As String, secret_key As String)
    On Error GoTo error

    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    Dim ordersString As String
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/openOrders?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
        For Each item In json
            ordersString = ordersString & ChrW(8658) & " "
            For Each Child In item
                ordersString = ordersString & "[ " & Child & " : " & item(Child) & " ] "
            Next Child
            ordersString = ordersString & vbNewLine & vbNewLine
        Next item
    frmAllOrders.lblAllOrders.Text = ordersString
    
done:
    Exit Sub
error:
        MsgBox "An error occured : " & xmlhttp.responseText
End Sub
