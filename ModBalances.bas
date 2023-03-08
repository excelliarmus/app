Attribute VB_Name = "ModBalances"


Sub UpdateBalances(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
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
    
    UserForm1.lblBalancesBNB.Caption = json("balances")(1)("free")
    UserForm1.lblBalancesBTC.Caption = json("balances")(2)("free")
    UserForm1.lblBalancesBUSD.Caption = json("balances")(3)("free")
    UserForm1.lblBalancesETH.Caption = json("balances")(4)("free")
    UserForm1.lblBalancesLTC.Caption = json("balances")(5)("free")
    UserForm1.lblBalancesTRX.Caption = json("balances")(6)("free")
    UserForm1.lblBalancesUSDT.Caption = json("balances")(7)("free")
    UserForm1.lblBalancesXRP.Caption = json("balances")(8)("free")
    
Done:
    Exit Sub
    
error_apikey:
        MsgBox "API Key / Secret Key invalid."

End Sub

Sub UpdateBNB(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesBNB.Caption = json("balances")(1)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateBTC(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesBTC.Caption = json("balances")(2)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateBUSD(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesBUSD.Caption = json("balances")(3)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateETH(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesETH.Caption = json("balances")(4)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateLTC(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesLTC.Caption = json("balances")(5)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateTRX(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesTRX.Caption = json("balances")(6)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateUSDT(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesUSDT.Caption = json("balances")(7)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub

Sub UpdateXRP(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/account?recvWindow=59999&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "GET", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    UserForm1.lblBalancesXRP.Caption = json("balances")(8)("free")
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub
