Attribute VB_Name = "modPrediction"
Option Explicit

Public isMLBotOn As Boolean 'Boolean to check if ML bot is ON
Public predictionLogsArray(12) As String 'Array of ML bot's logs

' sub to test the KNN algo
Sub testKNN()
    Dim knn As clsKNN
    Set knn = Factory.CreateKNN(k:=2)
    Dim arr1() As Variant
    Dim arr2() As Variant
    Dim arr3() As Variant
    Dim arr4() As Variant
    Dim arr5() As Variant
    Dim arr_pred() As Variant
    Dim arr_y() As Variant
    Dim arrMaster() As Variant
    
    'training data
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

' sub to power ON the ML bot (updates boolean and label value)
Sub activateMLBot()
    isMLBotOn = True
    UserForm1.lblPredictionStatus.BorderColor = &HFF00&
    UserForm1.lblPredictionStatus.Caption = "ON"
    UserForm1.lblPredictionStatus.ForeColor = &HFF00&
End Sub

' sub to power OFF the ML bot (updates boolean and label value)
Sub desactivateMLBot()
    isMLBotOn = False
    UserForm1.lblPredictionStatus.BorderColor = &HFF&
    UserForm1.lblPredictionStatus.Caption = "OFF"
    UserForm1.lblPredictionStatus.ForeColor = &HFF&
End Sub

' sub to start the ML trading bot (requires the ticker, the quantity of each trade...
' ... the k parameter, the number of klines, the frequency fo decision making
' and the discrimination rate between decions according to returns
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
                    Y(i) = -1
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
                If signal = "success" Then
                    addLog (ModTrading.get_time_for_logs & " : " & ChrW(9660) & " SELL " & qt & " " & symbol & " @ " & ModData.getCurrentPrice(symbol) & " (KNN)")
                Else
                    addLog (signal)
                End If
            Else
                addLog (ModTrading.get_time_for_logs & " : Do nothing " & ChrW(9787))
            End If
            
            ' Call ModBalances.UpdateBalances(UserForm1.inputBalances1, UserForm1.inputBalances2)
            Application.Wait (Now + TimeValue("00:00:" & frequency))
            DoEvents
        Loop
    End If
End Sub

' sub to diplay on userform the KNNs
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

' sub to diplay the most present label among KNNs
Sub displayMost(i As Integer)
    If i = 1 Then
        UserForm1.inputPredictionMostCommon = "BUY"
    ElseIf i = -1 Then
        UserForm1.inputPredictionMostCommon = "SELL"
    Else
        UserForm1.inputPredictionMostCommon = "HOLD"
    End If
    
End Sub

' sub to display logs of the ML trading bot
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

' sub to add a log to the global logs array
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
