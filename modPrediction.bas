Attribute VB_Name = "modPrediction"
Public isMLBotOn As Boolean

Sub test()
    Dim knn As clsKNN
    Set knn = Factory.CreateKNN(k:=2)
    Dim arr1() As Variant
    Dim arr2() As Variant
    Dim arr_pred() As Variant
    Dim arr_y() As Variant
    Dim arrMaster() As Variant
    'training
    arr1 = Array(1, 1, 1, 1, 1)
    arr2 = Array(1, 1, 1, 1, 1)
    arr3 = Array(2, 2, 2, 2, 2)
    arr4 = Array(2, 2, 2, 2, 2)
    arr5 = Array(3, 3, 3, 3, 3)
    arrMaster = Array(arr1, arr2, arr3, arr4, arr5)
    arr_y = Array(1, 1, 2, 2, 3)
    Call knn.fit(arrMaster, arr_y)
    
    'predicting
    arr_pred = Array(3, 3, 3, 3, 3)
    Debug.Print (VarType(arr_pred))
    MsgBox (knn.predict(arr_pred))
End Sub

Sub predict(window As Integer, k As Integer)
Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Dim json As Object
Dim OHLCopen As Double
Dim OHLChigh As Double
Dim OHLClow As Double
Dim OHLCclose As Double
Dim X() As Variant
ReDim Preserve X(0 To window - 2)
Dim y() As Integer
ReDim Preserve y(0 To window - 2)

'On Error GoTo noticker



Url = "https://api.binance.com/api/v3/klines?symbol=BTCUSDT&interval=1m&limit=" & window
xmlhttp.Open "GET", Url, False
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
For i = 0 To window - 2
    OHLCopen = CDbl(Replace(json(i + 1)(2), ".", ","))
    OHLChigh = CDbl(Replace(json(i + 1)(3), ".", ","))
    OHLClow = CDbl(Replace(json(i + 1)(4), ".", ","))
    OHLCclose = CDbl(Replace(json(i + 1)(5), ".", ","))
    OHLCclose_next = CDbl(Replace(json(i + 2)(5), ".", ","))
    
    X(i) = Array(OHLCopen, OHLChigh, OHLClow, OHLCclose)
    
    If OHLCclose > OHLCclose_next Then
        y(i) = -1
    ElseIf OHLCclose < OHLCclose_next Then
        y(i) = 1
    End If
Next



    Dim OHLCopen_to_predict As Double
    Dim OHLChigh_to_predict As Double
    Dim OHLClow_to_predict As Double
    Dim OHLCclose_to_predict As Double
    
    OHLCopen_to_predict = CDbl(Replace(json(window)(2), ".", ","))
    OHLChigh_to_predict = CDbl(Replace(json(window)(3), ".", ","))
    OHLClow_to_predict = CDbl(Replace(json(window)(4), ".", ","))
    OHLCclose_to_predict = CDbl(Replace(json(window)(5), ".", ","))
    'Dim arr_to_predict As Variant
    'arr_to_predict = Array(OHLCopen_to_predict, OHLChigh_to_predict, OHLClow_to_predict, OHLCclose_to_predict)
    
    Dim knn As clsKNN
    Set knn = Factory.CreateKNN(k:=k)
    Call knn.fit(X, y)
    
    'predicting
    Dim arr_to_predict() As Variant
    arr_to_predict = Array(OHLCopen_to_predict, OHLChigh_to_predict, OHLClow_to_predict, OHLCclose_to_predict)
    MsgBox (knn.predict(arr_to_predict))

    
    'Dim knn As clsKNN
    'Set knn = Factory.CreateKNN(k:=2)
    'Call knn.fit(X, y)
    'MsgBox (knn.predict(arr_to_predict))

'OHLCopen = CDbl(Replace(json(50)(2), ".", ","))
'OHLChigh = CDbl(Replace(json(50)(3), ".", ","))
'OHLClow = CDbl(Replace(json(50)(4), ".", ","))
'OHLCclose = CDbl(Replace(json(50)(5), ".", ","))

'Sheets("Data").Cells(30, 2) = OHLCopen
'Sheets("Data").Cells(30, 3) = OHLChigh
'Sheets("Data").Cells(30, 4) = OHLClow
'Sheets("Data").Cells(30, 5) = OHLCclose

'Debug.Print OHLCopen
'Debug.Print OHLChigh
'Debug.Print OHLClow
'Debug.Print OHLCclose


'Done:  Exit Sub

'noticker:
    'isDataStream1On = False
    'MsgBox "This trading pair '" & ticker & "' is not supported on Binance."



End Sub

Sub predict2(window As Integer, k As Integer)

Dim iter As Integer
For iter = 1 To 5



Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Dim json As Object
Dim OHLCopen As Double
Dim OHLChigh As Double
Dim OHLClow As Double
Dim OHLCclose As Double
Dim X() As Variant
ReDim Preserve X(0 To window - 2)
Dim y() As Integer
ReDim Preserve y(0 To window - 2)

'On Error GoTo noticker



Url = "https://api.binance.com/api/v3/klines?symbol=BTCUSDT&interval=1m&limit=" & window
xmlhttp.Open "GET", Url, False
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
For i = 0 To window - 2
    OHLCopen = CDbl(Replace(json(i + 1)(2), ".", ","))
    OHLChigh = CDbl(Replace(json(i + 1)(3), ".", ","))
    OHLClow = CDbl(Replace(json(i + 1)(4), ".", ","))
    OHLCclose = CDbl(Replace(json(i + 1)(5), ".", ","))
    OHLCclose_next = CDbl(Replace(json(i + 2)(5), ".", ","))
    
    X(i) = Array(OHLCopen, OHLChigh, OHLClow, OHLCclose)
    
    If OHLCclose > OHLCclose_next Then
        y(i) = -1
    ElseIf OHLCclose < OHLCclose_next Then
        y(i) = 1
    End If
Next



    Dim OHLCopen_to_predict As Double
    Dim OHLChigh_to_predict As Double
    Dim OHLClow_to_predict As Double
    Dim OHLCclose_to_predict As Double
    
    OHLCopen_to_predict = CDbl(Replace(json(window)(2), ".", ","))
    OHLChigh_to_predict = CDbl(Replace(json(window)(3), ".", ","))
    OHLClow_to_predict = CDbl(Replace(json(window)(4), ".", ","))
    OHLCclose_to_predict = CDbl(Replace(json(window)(5), ".", ","))
    'Dim arr_to_predict As Variant
    'arr_to_predict = Array(OHLCopen_to_predict, OHLChigh_to_predict, OHLClow_to_predict, OHLCclose_to_predict)
    
    Dim knn As clsKNN
    Set knn = Factory.CreateKNN(k:=k)
    Call knn.fit(X, y)
    
    'predicting
    Dim arr_to_predict() As Variant
    arr_to_predict = Array(OHLCopen_to_predict, OHLChigh_to_predict, OHLClow_to_predict, OHLCclose_to_predict)
    MsgBox (knn.predict(arr_to_predict))

    
    'Dim knn As clsKNN
    'Set knn = Factory.CreateKNN(k:=2)
    'Call knn.fit(X, y)
    'MsgBox (knn.predict(arr_to_predict))

'OHLCopen = CDbl(Replace(json(50)(2), ".", ","))
'OHLChigh = CDbl(Replace(json(50)(3), ".", ","))
'OHLClow = CDbl(Replace(json(50)(4), ".", ","))
'OHLCclose = CDbl(Replace(json(50)(5), ".", ","))

'Sheets("Data").Cells(30, 2) = OHLCopen
'Sheets("Data").Cells(30, 3) = OHLChigh
'Sheets("Data").Cells(30, 4) = OHLClow
'Sheets("Data").Cells(30, 5) = OHLCclose

'Debug.Print OHLCopen
'Debug.Print OHLChigh
'Debug.Print OHLClow
'Debug.Print OHLCclose


'Done:  Exit Sub

'noticker:
    'isDataStream1On = False
    'MsgBox "This trading pair '" & ticker & "' is not supported on Binance."


Next iter
End Sub

Sub activateMLBot()
    isMLBotOn = True
End Sub

Sub desactivateMLBot()
    isMLBotOn = False
End Sub

Sub startBot(symbol As String, qt As Double, k As Integer, window As Integer, frequency As Integer, discrimination As Double)
    Call activateMLBot
    Dim knn As clsKNN
    Set knn = Factory.CreateKNN(k:=k)
    Do Until Not isMLBotOn
        Dim X() As Variant
        ReDim Preserve X(0 To window - 2)
        Dim y() As Integer
        ReDim Preserve y(0 To window - 2)
        Dim xmlhttp As Object
        Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        Dim json As Object
        Dim OHLCopen As Double
        Dim OHLChigh As Double
        Dim OHLClow As Double
        Dim OHLCclose As Double
        Dim prediction As Integer
        Dim arr_to_predict() As Variant
        Url = "https://api.binance.com/api/v3/klines?symbol=" & symbol & "&interval=1m&limit=" & window
        xmlhttp.Open "GET", Url, False
        xmlhttp.Send
        Set json = JsonConverter.ParseJson(xmlhttp.responseText)
        Dim i As Integer
        For i = 0 To window - 2
            OHLCopen = CDbl(Replace(json(i + 1)(2), ".", ","))
            OHLChigh = CDbl(Replace(json(i + 1)(3), ".", ","))
            OHLClow = CDbl(Replace(json(i + 1)(4), ".", ","))
            OHLCclose = CDbl(Replace(json(i + 1)(5), ".", ","))
            OHLCclose_next = CDbl(Replace(json(i + 2)(5), ".", ","))
            
            X(i) = Array(OHLCopen, OHLChigh, OHLClow, OHLCclose)
            
            If OHLCclose > OHLCclose_next Then
                y(i) = -1
            ElseIf OHLCclose < OHLCclose_next Then
                y(i) = 1
            End If
        Next
        Call knn.fit(X, y)
        Dim OHLCopen_to_predict As Double
        Dim OHLChigh_to_predict As Double
        Dim OHLClow_to_predict As Double
        Dim OHLCclose_to_predict As Double
        OHLCopen_to_predict = CDbl(Replace(json(window)(2), ".", ","))
        OHLChigh_to_predict = CDbl(Replace(json(window)(3), ".", ","))
        OHLClow_to_predict = CDbl(Replace(json(window)(4), ".", ","))
        OHLCclose_to_predict = CDbl(Replace(json(window)(5), ".", ","))
        arr_to_predict = Array(OHLCopen_to_predict, OHLChigh_to_predict, OHLClow_to_predict, OHLCclose_to_predict)
        prediction = knn.predict(arr_to_predict)
        MsgBox (prediction)
        Application.Wait (Now + TimeValue("00:00:" & frequency))
        DoEvents
    Loop
End Sub
