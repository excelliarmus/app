Attribute VB_Name = "modPrediction"
Option Explicit

Public isMLBotOn As Boolean
Public predictionLogsArray(12) As String

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
Dim Y() As Integer
ReDim Preserve Y(0 To window - 2)

'On Error GoTo noticker



url = "https://api.binance.com/api/v3/klines?symbol=BTCUSDT&interval=1m&limit=" & window
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
For i = 0 To window - 2
    OHLCopen = CDbl(Replace(json(i + 1)(2), ".", ","))
    OHLChigh = CDbl(Replace(json(i + 1)(3), ".", ","))
    OHLClow = CDbl(Replace(json(i + 1)(4), ".", ","))
    OHLCclose = CDbl(Replace(json(i + 1)(5), ".", ","))
    OHLCclose_next = CDbl(Replace(json(i + 2)(5), ".", ","))
    
    X(i) = Array(OHLCopen, OHLChigh, OHLClow, OHLCclose)
    
    If OHLCclose > OHLCclose_next Then
        Y(i) = -1
    ElseIf OHLCclose < OHLCclose_next Then
        Y(i) = 1
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
    Call knn.fit(X, Y)
    
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
Dim Y() As Integer
ReDim Preserve Y(0 To window - 2)

'On Error GoTo noticker



url = "https://api.binance.com/api/v3/klines?symbol=BTCUSDT&interval=1m&limit=" & window
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
For i = 0 To window - 2
    OHLCopen = CDbl(Replace(json(i + 1)(2), ".", ","))
    OHLChigh = CDbl(Replace(json(i + 1)(3), ".", ","))
    OHLClow = CDbl(Replace(json(i + 1)(4), ".", ","))
    OHLCclose = CDbl(Replace(json(i + 1)(5), ".", ","))
    OHLCclose_next = CDbl(Replace(json(i + 2)(5), ".", ","))
    
    X(i) = Array(OHLCopen, OHLChigh, OHLClow, OHLCclose)
    
    If OHLCclose > OHLCclose_next Then
        Y(i) = -1
    ElseIf OHLCclose < OHLCclose_next Then
        Y(i) = 1
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
    Call knn.fit(X, Y)
    
    'predicting
    Dim arr_to_predict() As Variant
    arr_to_predict = Array(OHLCopen_to_predict, OHLChigh_to_predict, OHLClow_to_predict, OHLCclose_to_predict)
    MsgBox (knn.predict(arr_to_predict))
    Dim tmp() As Integer
    


    
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
    UserForm1.lblPredictionStatus.BorderColor = &HFF00&
    UserForm1.lblPredictionStatus.Caption = "ON"
    UserForm1.lblPredictionStatus.ForeColor = &HFF00&
End Sub

Sub desactivateMLBot()
    isMLBotOn = False
    UserForm1.lblPredictionStatus.BorderColor = &HFF&
    UserForm1.lblPredictionStatus.Caption = "OFF"
    UserForm1.lblPredictionStatus.ForeColor = &HFF&
End Sub

Sub startBot(symbol As String, qt As String, k As Integer, window As Integer, frequency As Integer, discrimination As Double)
    Call activateMLBot
    Dim knn As clsKNN
    Set knn = Factory.CreateKNN(k:=k)
    Dim margin As Double
    margin = CDbl(Replace(discrimination, ".", ","))
    Dim resp As String
    resp = ModBalances.checkKeys(UserForm1.inputBalances1, UserForm1.inputBalances2)
    If resp = "error" Then
        desactivateMLBot
        addLog ("API Keys invalid : bot stopped")
        Call ModBalances.powerOffGlobalStream
    Else
        Do Until Not isMLBotOn
            Dim X() As Variant
            ReDim Preserve X(0 To window - 2)
            Dim Y() As Integer
            ReDim Preserve Y(0 To window - 2)
            Dim xmlhttp As Object
            Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            Dim json As Object
            Dim OHLCopen As Double
            Dim OHLChigh As Double
            Dim OHLClow As Double
            Dim OHLCclose As Double
            Dim OHLCclose_next As Double
            Dim prediction As Integer
            Dim arr_to_predict() As Variant
            Dim url As String
            url = "https://api.binance.com/api/v3/klines?symbol=" & symbol & "&interval=1m&limit=" & window
            xmlhttp.Open "GET", url, False
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
                
                If OHLCclose_next > OHLCclose + (OHLCclose * margin) Then
                    Y(i) = 1
                ElseIf OHLCclose_next < OHLCclose - (OHLCclose * margin) Then
                    Y(i) = 1
                Else
                    Y(i) = 0
                End If
            Next
            Call knn.fit(X, Y)
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
            displayKNN (knn.knn)
            displayMost (knn.most)
            Dim signal As String
            If prediction = 1 Then
                signal = ModTrading.placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "BUY", symbol, qt)
                If signal = "success" Then
                    addLog (ModTrading.get_time_for_logs & " : " & ChrW(9650) & " BUY " & qt & " " & symbol & " @ " & ModData.getCurrentPrice(symbol) & " (KNN)")
                Else
                    addLog (signal)
                End If
            ElseIf prediction = -1 Then
                signal = ModTrading.placeMarketOrder(UserForm1.inputBalances1, UserForm1.inputBalances2, "SELL", symbol, qt)
                MsgBox signal
                If signal = "success" Then
                    addLog (ModTrading.get_time_for_logs & " : " & ChrW(9660) & " SELL " & qt & " " & symbol & " @ " & ModData.getCurrentPrice(symbol) & " (KNN)")
                Else
                    addLog (signal)
                End If
            Else
                addLog (ModTrading.get_time_for_logs & " : Do nothing " & ChrW(9787))
            End If
            
            Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)
            Application.Wait (Now + TimeValue("00:00:" & frequency))
            DoEvents
        Loop
    End If
End Sub

Sub displayKNN(arr As Variant)
    Dim str As String
    str = ""
    Dim a As Variant
    For Each a In arr
        If a = 1 Then
            str = str & "BUY "
        ElseIf a = -1 Then
            str = str & "SELL "
        Else
            str = str & "HOLD "
        End If
        
    Next a
    UserForm1.lblPredictionTime.Caption = ModTrading.get_time_for_logs
    UserForm1.inputPredictionLabels.Text = str
    
End Sub

Sub displayMost(i As Integer)
    If i = 1 Then
        UserForm1.inputPredictionMostCommon = "BUY"
    ElseIf i = -1 Then
        UserForm1.inputPredictionMostCommon = "SELL"
    Else
        UserForm1.inputPredictionMostCommon = "HOLD"
    End If
    
End Sub

Sub displayLogs()
    UserForm1.lblPredictionLog1 = predictionLogsArray(0)
    If predictionLogsArray(0) Like "*BUY*" Then
        UserForm1.lblPredictionLog1.ForeColor = &HFF00&
    ElseIf predictionLogsArray(0) Like "*SELL*" Then
        UserForm1.lblPredictionLog1.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog1.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog2 = predictionLogsArray(1)
    If predictionLogsArray(1) Like "*BUY*" Then
        UserForm1.lblPredictionLog2.ForeColor = &HFF00&
    ElseIf predictionLogsArray(1) Like "*SELL*" Then
        UserForm1.lblPredictionLog2.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog2.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog3 = predictionLogsArray(2)
    If predictionLogsArray(2) Like "*BUY*" Then
        UserForm1.lblPredictionLog3.ForeColor = &HFF00&
    ElseIf predictionLogsArray(2) Like "*SELL*" Then
        UserForm1.lblPredictionLog3.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog3.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog4 = predictionLogsArray(3)
    If predictionLogsArray(3) Like "*BUY*" Then
        UserForm1.lblPredictionLog4.ForeColor = &HFF00&
    ElseIf predictionLogsArray(3) Like "*SELL*" Then
        UserForm1.lblPredictionLog4.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog4.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog5 = predictionLogsArray(4)
    If predictionLogsArray(4) Like "*BUY*" Then
        UserForm1.lblPredictionLog5.ForeColor = &HFF00&
    ElseIf predictionLogsArray(4) Like "*SELL*" Then
        UserForm1.lblPredictionLog5.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog5.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog6 = predictionLogsArray(5)
    If predictionLogsArray(5) Like "*BUY*" Then
        UserForm1.lblPredictionLog6.ForeColor = &HFF00&
    ElseIf predictionLogsArray(5) Like "*SELL*" Then
        UserForm1.lblPredictionLog6.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog6.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog6 = predictionLogsArray(5)
    If predictionLogsArray(5) Like "*BUY*" Then
        UserForm1.lblPredictionLog6.ForeColor = &HFF00&
    ElseIf predictionLogsArray(5) Like "*SELL*" Then
        UserForm1.lblPredictionLog6.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog6.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog7 = predictionLogsArray(6)
    If predictionLogsArray(6) Like "*BUY*" Then
        UserForm1.lblPredictionLog7.ForeColor = &HFF00&
    ElseIf predictionLogsArray(6) Like "*SELL*" Then
        UserForm1.lblPredictionLog7.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog7.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog8 = predictionLogsArray(7)
    If predictionLogsArray(7) Like "*BUY*" Then
        UserForm1.lblPredictionLog8.ForeColor = &HFF00&
    ElseIf predictionLogsArray(7) Like "*SELL*" Then
        UserForm1.lblPredictionLog8.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog8.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog9 = predictionLogsArray(8)
    If predictionLogsArray(8) Like "*BUY*" Then
        UserForm1.lblPredictionLog9.ForeColor = &HFF00&
    ElseIf predictionLogsArray(8) Like "*SELL*" Then
        UserForm1.lblPredictionLog9.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog9.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog10 = predictionLogsArray(9)
    If predictionLogsArray(9) Like "*BUY*" Then
        UserForm1.lblPredictionLog10.ForeColor = &HFF00&
    ElseIf predictionLogsArray(9) Like "*SELL*" Then
        UserForm1.lblPredictionLog10.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog10.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog11 = predictionLogsArray(10)
    If predictionLogsArray(10) Like "*BUY*" Then
        UserForm1.lblPredictionLog11.ForeColor = &HFF00&
    ElseIf predictionLogsArray(10) Like "*SELL*" Then
        UserForm1.lblPredictionLog11.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog11.ForeColor = &H80FF&
    End If
    
    UserForm1.lblPredictionLog12 = predictionLogsArray(11)
    If predictionLogsArray(11) Like "*BUY*" Then
        UserForm1.lblPredictionLog12.ForeColor = &HFF00&
    ElseIf predictionLogsArray(11) Like "*SELL*" Then
        UserForm1.lblPredictionLog12.ForeColor = &HFF&
    Else
        UserForm1.lblPredictionLog12.ForeColor = &H80FF&
    End If
    
    'UserForm1.lblPredictionLog2 = predictionLogsArray(1)
    'UserForm1.lblPredictionLog3 = predictionLogsArray(2)
    'UserForm1.lblPredictionLog4 = predictionLogsArray(3)
    'UserForm1.lblPredictionLog5 = predictionLogsArray(4)
    'UserForm1.lblPredictionLog6 = predictionLogsArray(5)
    'UserForm1.lblPredictionLog7 = predictionLogsArray(6)
    'UserForm1.lblPredictionLog8 = predictionLogsArray(7)
    'UserForm1.lblPredictionLog9 = predictionLogsArray(8)
    'UserForm1.lblPredictionLog10 = predictionLogsArray(9)
    'UserForm1.lblPredictionLog11 = predictionLogsArray(10)
    'UserForm1.lblPredictionLog12 = predictionLogsArray(11)
End Sub

Sub addLog(log As String)
    Dim len_arr As Integer
    Dim i As Integer
    len_arr = UBound(predictionLogsArray) - LBound(predictionLogsArray)
    For i = len_arr - 1 To 1 Step -1
    ' Debug.Print (i)
        predictionLogsArray(i) = predictionLogsArray(i - 1)
    Next
    predictionLogsArray(LBound(predictionLogsArray)) = log
    displayLogs
End Sub
