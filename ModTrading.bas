Attribute VB_Name = "ModTrading"
Sub buyBTCUSDT(APIkey As String, secret_key As String)
    Dim xmlhttp As Object
    Dim timestamp As LongLong
    Dim signature As String
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    timestamp = ModBinanceRequests.getTimeStampForBinance
    signature = ModBinanceRequests.getSignature("recvWindow=59999&symbol=BTCUSDT&side=BUY&type=MARKET&quantity=0.001&timestamp=" & timestamp, secret_key)
    Url = "https://testnet.binance.vision/api/v3/order?recvWindow=59999&symbol=BTCUSDT&side=BUY&type=MARKET&quantity=0.001&timestamp=" & timestamp & "&signature=" & signature
    xmlhttp.Open "POST", Url, False
    xmlhttp.setRequestHeader "X-MBX-APIKEY", APIkey
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    On Error GoTo error_apikey
    MsgBox "success"
    ' MsgBox (json("orderId"))
Done:
    Exit Sub
error_apikey:
        MsgBox "API Key / Secret Key invalid."
End Sub
