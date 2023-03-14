Attribute VB_Name = "ModTrading"
Public logsArray(12) As String
Public isRandomBotOn As Boolean

Function placeMarketOrder(APIkey As String, secret_key As String, side As String, ticker As String, qt As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    On Error GoTo error
    If side = "BUY" Then
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=BUY&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp, secret_key)
        Url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=BUY&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp & "&signature=" & signature
    Else
        signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=" & ticker & "&side=SELL&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp, secret_key)
        Url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=" & ticker & "&side=SELL&type=MARKET&quantity=" & qt & "&timestamp=" & timestamp & "&signature=" & signature
    End If
    xmlhttp.Open "POST", Url, False
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
    
    
Done:
    Exit Function
error:
        MsgBox "An error occured : " & xmlhttp.responseText
End Function

Sub placeOrder(APIkey As String, secret_key As String)
    Dim ticker As String
    Dim qt As String
    Dim signal As String
    
    ticker = UserForm1.inputTradingTicker
    qt = UserForm1.inputTradingQuantity
    

    
    If UserForm1.tglTradingMarket.Value = True Then
        If UserForm1.optTradingBuy.Value = True Then
            'BUY MARKET
            signal = placeMarketOrder(APIkey, secret_key, "BUY", ticker, qt)
            If signal = "success" Then
                addLog ("BUY " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
            ElseIf UserForm1.optTradingSell.Value = True Then
            'SELL MARKET
            signal = placeMarketOrder(APIkey, secret_key, "SELL", ticker, qt)
            If signal = "success" Then
                addLog ("SELL " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        End If
    ElseIf True Then
    
    
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
    
    ticker = UserForm1.inputTradingTicker
    qt = UserForm1.inputTradingQuantity
    
    Do Until Not isRandomBotOn
        Dim rand As Integer
        rand = Int(3 * Rnd) + 1
        If rand = 1 Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "BUY", ticker, qt)
            If signal = "success" Then
                addLog ("BUY " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        ElseIf rand = 2 Then
            signal = placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "SELL", ticker, qt)
            If signal = "success" Then
                addLog ("SELL " & qt & " " & ticker & " @ " & ModData.getCurrentPrice(ticker))
            Else
                addLog (signal)
            End If
        Else
            addLog ("Let's take a break")
        End If
        Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
    Loop
End Sub
