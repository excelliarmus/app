Attribute VB_Name = "ModBalances"
' Booleans to check if stream is ON or OFF
Public isBalancesGlobalStreamOn As Boolean
Public isBNBStreamOn As Boolean
Public isBTCStreamOn As Boolean
Public isBUSDStreamOn As Boolean
Public isETHStreamOn As Boolean
Public isLTCStreamOn As Boolean
Public isTRXStreamOn As Boolean
Public isUSDTStreamOn As Boolean
Public isXRPStreamOn As Boolean

' sub to power ON the global stream (updates the boolean and front-end)
Sub powerOnGlobalStream()
    isBalancesGlobalStreamOn = True
    UserForm1.lblBalancesStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesStatus.Caption = "ON"
    UserForm1.lblBalancesStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the global stream (updates the boolean and front-end)
Sub powerOffGlobalStream()
    isBalancesGlobalStreamOn = False
    UserForm1.lblBalancesStatus.BorderColor = &HFF&
    UserForm1.lblBalancesStatus.Caption = "OFF"
    UserForm1.lblBalancesStatus.ForeColor = &HFF&
End Sub

' sub to power ON the BNB stream (updates the boolean and front-end)
Sub powerOnBNBStream()
    isBNBStreamOn = True
    UserForm1.lblBalancesBNBStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesBNBStatus.Caption = "ON"
    UserForm1.lblBalancesBNBStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the BNB stream (updates the boolean and front-end)
Sub powerOffBNBStream()
    isBNBStreamOn = False
    UserForm1.lblBalancesBNBStatus.BorderColor = &HFF&
    UserForm1.lblBalancesBNBStatus.Caption = "OFF"
    UserForm1.lblBalancesBNBStatus.ForeColor = &HFF&
End Sub

' sub to power ON the BTC stream (updates the boolean and front-end)
Sub powerOnBTCStream()
    isBTCStreamOn = True
    UserForm1.lblBalancesBTCStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesBTCStatus.Caption = "ON"
    UserForm1.lblBalancesBTCStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the BTC stream (updates the boolean and front-end)
Sub powerOffBTCStream()
    isBTCStreamOn = False
    UserForm1.lblBalancesBTCStatus.BorderColor = &HFF&
    UserForm1.lblBalancesBTCStatus.Caption = "OFF"
    UserForm1.lblBalancesBTCStatus.ForeColor = &HFF&
End Sub

' sub to power ON the BUSD stream (updates the boolean and front-end)
Sub powerOnBUSDStream()
    isBUSDStreamOn = True
    UserForm1.lblBalancesBUSDStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesBUSDStatus.Caption = "ON"
    UserForm1.lblBalancesBUSDStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the BUSD stream (updates the boolean and front-end)
Sub powerOffBUSDStream()
    isBUSDStreamOn = False
    UserForm1.lblBalancesBUSDStatus.BorderColor = &HFF&
    UserForm1.lblBalancesBUSDStatus.Caption = "OFF"
    UserForm1.lblBalancesBUSDStatus.ForeColor = &HFF&
End Sub

' sub to power ON the ETH stream (updates the boolean and front-end)
Sub powerOnETHStream()
    isETHStreamOn = True
    UserForm1.lblBalancesETHStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesETHStatus.Caption = "ON"
    UserForm1.lblBalancesETHStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the ETH stream (updates the boolean and front-end)
Sub powerOffETHStream()
    isETHStreamOn = False
    UserForm1.lblBalancesETHStatus.BorderColor = &HFF&
    UserForm1.lblBalancesETHStatus.Caption = "OFF"
    UserForm1.lblBalancesETHStatus.ForeColor = &HFF&
End Sub

' sub to power ON the LTC stream (updates the boolean and front-end)
Sub powerOnLTCStream()
    isLTCStreamOn = True
    UserForm1.lblBalancesLTCStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesLTCStatus.Caption = "ON"
    UserForm1.lblBalancesLTCStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the LTC stream (updates the boolean and front-end)
Sub powerOffLTCStream()
    isLTCStreamOn = False
    UserForm1.lblBalancesLTCStatus.BorderColor = &HFF&
    UserForm1.lblBalancesLTCStatus.Caption = "OFF"
    UserForm1.lblBalancesLTCStatus.ForeColor = &HFF&
End Sub

' sub to power ON the TRX stream (updates the boolean and front-end)
Sub powerOnTRXStream()
    isTRXStreamOn = True
    UserForm1.lblBalancesTRXStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesTRXStatus.Caption = "ON"
    UserForm1.lblBalancesTRXStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the TRX stream (updates the boolean and front-end)
Sub powerOffTRXStream()
    isTRXStreamOn = False
    UserForm1.lblBalancesTRXStatus.BorderColor = &HFF&
    UserForm1.lblBalancesTRXStatus.Caption = "OFF"
    UserForm1.lblBalancesTRXStatus.ForeColor = &HFF&
End Sub

' sub to power ON the USDT stream (updates the boolean and front-end)
Sub powerOnUSDTStream()
    isUSDTStreamOn = True
    UserForm1.lblBalancesUSDTStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesUSDTStatus.Caption = "ON"
    UserForm1.lblBalancesUSDTStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the USDT stream (updates the boolean and front-end)
Sub powerOffUSDTStream()
    isUSDTStreamOn = False
    UserForm1.lblBalancesUSDTStatus.BorderColor = &HFF&
    UserForm1.lblBalancesUSDTStatus.Caption = "OFF"
    UserForm1.lblBalancesUSDTStatus.ForeColor = &HFF&
End Sub

' sub to power ON the XRP stream (updates the boolean and front-end)
Sub powerOnXRPStream()
    isXRPStreamOn = True
    UserForm1.lblBalancesXRPStatus.BorderColor = &HFF00&
    UserForm1.lblBalancesXRPStatus.Caption = "ON"
    UserForm1.lblBalancesXRPStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the XRP stream (updates the boolean and front-end)
Sub powerOffXRPStream()
    isXRPStreamOn = False
    UserForm1.lblBalancesXRPStatus.BorderColor = &HFF&
    UserForm1.lblBalancesXRPStatus.Caption = "OFF"
    UserForm1.lblBalancesXRPStatus.ForeColor = &HFF&
End Sub

' function returns if Global Stream Boolean is True or False
Function get_isGlobalStream1On()
    get_isGlobalStream1On = isBalancesGlobalStreamOn
End Function

' functions returns "ok" if API keys are valid
Public Function checkKeys(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    Dim bnb As String
    On Error GoTo error
    bnb = json("balances")(1)("free")
error:
    checkKeys = "error"
    If bnb <> "" Then
        checkKeys = "ok"
    End If
End Function

' sub to update balances of each crypto and the overall USD balance
Sub UpdateBalances(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    
    
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    
    On Error GoTo error_apikey
    '{
    '  "makerCommission": 0,
    '  "takerCommission": 0,
    '  "buyerCommission": 0,
    '  "sellerCommission": 0,
    '  "commissionRates": {
    '    "maker": "0.00000000",
    '    "taker": "0.00000000",
    '    "buyer": "0.00000000",
    '    "seller": "0.00000000"
    '  },
    '  "canTrade": true,
    '  "canWithdraw": false,
    '  "canDeposit": false,
    '  "brokered": false,
    '  "requireSelfTradePrevention": false,
    '  "updateTime": 1678267449930,
    '  "accountType": "SPOT",
    '  "balances": [
    '    {
    '      "asset": "BNB",
    '      "free": "1000.00000000",
    '      "locked": "0.00000000"
    '    }(1),
    '    {
    '      "asset": "BTC",
    '      "free": "1.00000000",
    '      "locked": "0.00000000"
    '    }(2),
    '    {
    '      "asset": "BUSD",
    '      "free": "10000.00000000",
    '      "locked": "0.00000000"
    '    }(3),
    '    {
    '      "asset": "ETH",
    '      "free": "100.00000000",
    '      "locked": "0.00000000"
    '    }(4),
    '    {
    '      "asset": "LTC",
    '      "free": "500.00000000",
    '      "locked": "0.00000000"
    '    }(5),
    '    {
    '      "asset": "TRX",
    '      "free": "500000.00000000",
    '      "locked": "0.00000000"
    '    }(6),
    '    {
    '      "asset": "USDT",
    '      "free": "10000.00000000",
    '      "locked": "0.00000000"
    '    }(7),
    '    {
    '      "asset": "XRP",
    '      "free": "50000.00000000",
    '      "locked": "0.00000000"
    '    }(8)
    '  ],
    '  "permissions": [
    '    "SPOT"
    '  ]
    '}
    
    'BNB Balance ==>> json("balances")(1)("free")
    'For i = 1 To 8
    ' Sheets("Balances").Cells(i, 2) = json("balances")(i)("free")
    'Next
    
    Dim bnb, btc, busd, eth, ltc, trx, usdt, xrp
    
    bnb = json("balances")(1)("free")
    UserForm1.lblBalancesBNB.Caption = bnb
    UserForm1.lblBalancesBNBtoUSD = Replace(getToUSD("BNBUSDT", CDbl(Replace(bnb, ".", ","))), ",", ".")
    
    btc = json("balances")(2)("free")
    UserForm1.lblBalancesBTC.Caption = btc
    UserForm1.lblBalancesBTCtoUSD = Replace(getToUSD("BTCUSDT", CDbl(Replace(btc, ".", ","))), ",", ".")
    
    busd = json("balances")(3)("free")
    UserForm1.lblBalancesBUSD.Caption = busd
    UserForm1.lblBalancesBUSDtoUSD = Replace(getToUSD("BUSDUSDT", CDbl(Replace(busd, ".", ","))), ",", ".")
    
    eth = json("balances")(4)("free")
    UserForm1.lblBalancesETH.Caption = eth
    UserForm1.lblBalancesETHtoUSD = Replace(getToUSD("ETHUSDT", CDbl(Replace(eth, ".", ","))), ",", ".")
    
    ltc = json("balances")(5)("free")
    UserForm1.lblBalancesLTC.Caption = ltc
    UserForm1.lblBalancesLTCtoUSD = Replace(getToUSD("LTCUSDT", CDbl(Replace(ltc, ".", ","))), ",", ".")
    
    trx = json("balances")(6)("free")
    UserForm1.lblBalancesTRX.Caption = trx
    UserForm1.lblBalancesTRXtoUSD = Replace(getToUSD("TRXUSDT", CDbl(Replace(trx, ".", ","))), ",", ".")
    
    usdt = json("balances")(7)("free")
    UserForm1.lblBalancesUSDT.Caption = usdt
    UserForm1.lblBalancesUSDTtoUSD = Replace(CDbl(Replace(usdt, ".", ",")) * 0.989, ",", ".")
    
    xrp = json("balances")(8)("free")
    UserForm1.lblBalancesXRP.Caption = xrp
    UserForm1.lblBalancesXRPtoUSD = Replace(getToUSD("XRPUSDT", CDbl(Replace(xrp, ".", ","))), ",", ".")

    
    'UserForm1.lblBalancesBUSD.Caption = json("balances")(3)("free")
    'UserForm1.lblBalancesETH.Caption = json("balances")(4)("free")
    'UserForm1.lblBalancesLTC.Caption = json("balances")(5)("free")
    'UserForm1.lblBalancesTRX.Caption = json("balances")(6)("free")
    'UserForm1.lblBalancesUSDT.Caption = json("balances")(7)("free")
    'UserForm1.lblBalancesXRP.Caption = json("balances")(8)("free")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
    
done:
    Exit Sub
    
error_apikey:
        powerOffGlobalStream
        MsgBox "API Key / Secret Key invalid."

End Sub

' function returns the sum of each crypto balance in USD equivalent
Function getOverallBalanceUSD()
    Dim res
    res = CDbl(Replace(UserForm1.lblBalancesBNBtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesBTCtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesBUSDtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesETHtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesLTCtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesTRXtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesUSDTtoUSD, ".", ",")) _
    + CDbl(Replace(UserForm1.lblBalancesXRPtoUSD, ".", ","))
    getOverallBalanceUSD = Replace(res, ",", ".")
End Function

' function returns USD equivalent of a ticker and quantity
Function getToUSD(ticker As String, quantity As Double)
    Dim res
    res = ModData.getCurrentPrice(ticker) * quantity * 0.989 ' 1 USD = 0.989 USDT
    'MsgBox "res = " & ModData.getCurrentPrice(ticker) & " * " & quantity & " * 0.989"
    getToUSD = res
End Function

' sub to update only BNB balance + USD equivalent (requires API key + secret key)
Sub UpdateBNB(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesBNB.Caption = json("balances")(1)("free")
    bnb = json("balances")(1)("free")
    UserForm1.lblBalancesBNB.Caption = bnb
    UserForm1.lblBalancesBNBtoUSD = Replace(getToUSD("BNBUSDT", CDbl(Replace(bnb, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffBNBStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only BTC balance + USD equivalent (requires API key + secret key)
Sub UpdateBTC(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesBTC.Caption = json("balances")(2)("free")
    btc = json("balances")(2)("free")
    UserForm1.lblBalancesBTC.Caption = btc
    UserForm1.lblBalancesBTCtoUSD = Replace(getToUSD("BTCUSDT", CDbl(Replace(btc, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffBTCStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only BUSD balance + USD equivalent (requires API key + secret key)
Sub UpdateBUSD(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesBUSD.Caption = json("balances")(3)("free")
    busd = json("balances")(3)("free")
    UserForm1.lblBalancesBUSD.Caption = busd
    UserForm1.lblBalancesBUSDtoUSD = Replace(getToUSD("BUSDUSDT", CDbl(Replace(busd, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffBUSDStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only ETH balance + USD equivalent (requires API key + secret key)
Sub UpdateETH(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesETH.Caption = json("balances")(4)("free")
    eth = json("balances")(4)("free")
    UserForm1.lblBalancesETH.Caption = eth
    UserForm1.lblBalancesETHtoUSD = Replace(getToUSD("ETHUSDT", CDbl(Replace(eth, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffETHStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only LTC balance + USD equivalent (requires API key + secret key)
Sub UpdateLTC(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesLTC.Caption = json("balances")(5)("free")
    ltc = json("balances")(5)("free")
    UserForm1.lblBalancesLTC.Caption = ltc
    UserForm1.lblBalancesLTCtoUSD = Replace(getToUSD("LTCUSDT", CDbl(Replace(ltc, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffLTCStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only TRX balance + USD equivalent (requires API key + secret key)
Sub UpdateTRX(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesTRX.Caption = json("balances")(6)("free")
    trx = json("balances")(6)("free")
    UserForm1.lblBalancesTRX.Caption = trx
    UserForm1.lblBalancesTRXtoUSD = Replace(getToUSD("TRXUSDT", CDbl(Replace(trx, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffTRXStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only USDT balance + USD equivalent (requires API key + secret key)
Sub UpdateUSDT(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    ' UserForm1.lblBalancesUSDT.Caption = json("balances")(7)("free")
    usdt = json("balances")(7)("free")
    UserForm1.lblBalancesUSDT.Caption = usdt
    UserForm1.lblBalancesUSDTtoUSD = Replace(CDbl(Replace(usdt, ".", ",")) * 0.989, ",", ".")
done:
    Exit Sub
error_apikey:
        powerOffUSDTStream
        MsgBox "API Key / Secret Key invalid."
End Sub

' sub to update only XRP balance + USD equivalent (requires API key + secret key)
Sub UpdateXRP(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As Double
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    'UserForm1.lblBalancesXRP.Caption = json("balances")(8)("free")
    xrp = json("balances")(8)("free")
    UserForm1.lblBalancesXRP.Caption = xrp
    UserForm1.lblBalancesXRPtoUSD = Replace(getToUSD("XRPUSDT", CDbl(Replace(xrp, ".", ","))), ",", ".")
    UserForm1.lblBalancesOverall.Caption = getOverallBalanceUSD()
done:
    Exit Sub
error_apikey:
        powerOffXRPStream
        MsgBox "API Key / Secret Key invalid."
End Sub
