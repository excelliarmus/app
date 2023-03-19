Attribute VB_Name = "modPrediction"
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
    MsgBox (knn.predict(arr_pred))
End Sub

Sub predict()
Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Dim json As Object
Dim OHLCopen As Double
Dim OHLChigh As Double
Dim OHLClow As Double
Dim OHLCclose As Double

On Error GoTo noticker



Url = "https://api.binance.com/api/v3/klines?symbol=BTCUSDT&interval=1m&limit=100"
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

' Debug.Print json(100)(12)

Dim i As Integer
Dim j As Integer
j = 80

For i = 100 To 21 Step -1
'Debug.Print "i = " & i & " et j = " & j


OHLCopen = CDbl(Replace(json(i)(2), ".", ","))
OHLChigh = CDbl(Replace(json(i)(3), ".", ","))
OHLClow = CDbl(Replace(json(i)(4), ".", ","))
OHLCclose = CDbl(Replace(json(i)(5), ".", ","))

Sheets("Data").Cells(j, 1) = OHLCopen
Sheets("Data").Cells(j, 2) = OHLChigh
Sheets("Data").Cells(j, 3) = OHLClow
Sheets("Data").Cells(j, 4) = OHLCclose

j = j - 1
Next



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


Done:  Exit Sub

noticker:
    isDataStream1On = False
    MsgBox "This trading pair '" & ticker & "' is not supported on Binance."



End Sub
