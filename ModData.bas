Attribute VB_Name = "ModData"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public isDataStream1On As Boolean
Public isDataStream2On As Boolean
Public chart2array() As Double
Public canIncrementData2 As Boolean









Sub writeData1(ticker As String)

Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Dim json As Object
Dim OHLCopen As Double
Dim OHLChigh As Double
Dim OHLClow As Double
Dim OHLCclose As Double

On Error GoTo noticker



Url = "https://api.binance.com/api/v3/klines?symbol=" & ticker & "&interval=1m&limit=100"
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
'  ]
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
    MsgBox "This trading pair '" & ticker & "' is not supported on Binance."



End Sub

Sub displayData1()
 
nRows = Sheets("Data").UsedRange.Rows.Count
 
Dim OHLCChart As Chart
Set OHLCChart = Charts.Add
 
With OHLCChart
    .SetSourceData Source:=Sheets("Data").Range("a1:d" & nRows)
    .ChartType = xlStockOHLC
    With .ChartGroups(1)
     
        .UpBars.Interior.ColorIndex = 10
        .DownBars.Interior.ColorIndex = 3
    End With
    .PlotArea.Format.Fill.ForeColor.RGB = RGB(4, 4, 65)
    .ChartArea.Interior.Color = RGB(4, 4, 65)
    .HasLegend = False
    .Axes(xlValue, xlPrimary).TickLabels.Font.Color = RGB(255, 255, 255)
    .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 20
    '.Axes(xlCategory, xlPrimary).TickLabels.Font.Color = RGB(255, 255, 255)
    .HasAxis(xlCategory) = False
End With



ActiveChart.Export ThisWorkbook.Path & "\chart.jpg"
f = ActiveSheet.Name
Sheets(f).Select
ActiveWindow.SelectedSheets.Visible = False


fname = ThisWorkbook.Path & "\chart.jpg"

UserForm1.imgData1.Picture = LoadPicture(fname)
End Sub


Function get_isDataStream1On()
    get_isDataStream1On = isDataStream1On
End Function
Function get_isDataStream2On()
    get_isDataStream2On = isDataStream2On
End Function

Sub activateDataStream1()
    isDataStream1On = True
End Sub

Sub activateDataStream2()
    isDataStream2On = True
End Sub

Sub desactivateDataStream1()
    isDataStream1On = False
End Sub

Sub desactivateDataStream2()
    isDataStream2On = False
End Sub


Sub displayData2(ticker As String)


Dim mychart As Chart


Set mychart = Charts.Add

Call incrementChart2Array(ticker)

With mychart
    .SeriesCollection.NewSeries
    .SeriesCollection(1).XValues = chart2array
    .SeriesCollection(1).Values = chart2array
    .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 255, 255)
    .ChartType = xlLine
    .HasAxis(xlCategory) = False
    .HasLegend = False
    .PlotArea.Format.Fill.ForeColor.RGB = RGB(4, 4, 65)
    .ChartArea.Interior.Color = RGB(4, 4, 65)
    .Axes(xlValue, xlPrimary).TickLabels.Font.Color = RGB(255, 255, 255)
    .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 20
    .SeriesCollection(1).MarkerSize = 30
    .SeriesCollection(1).MarkerBackgroundColor = RGB(255, 255, 255)
    .SeriesCollection(1).MarkerForegroundColor = RGB(255, 255, 255)
End With




ActiveChart.Export ThisWorkbook.Path & "\chart2.jpg"
f = ActiveSheet.Name
Sheets(f).Select
ActiveWindow.SelectedSheets.Visible = False


fname = ThisWorkbook.Path & "\chart2.jpg"

UserForm1.imgData2.Picture = LoadPicture(fname)
End Sub

Sub incrementChart2Array(ticker As String)

Dim len_arr As Integer


If Not canIncrementData2 Then
    ReDim Preserve chart2array(0)
    chart2array(UBound(chart2array)) = getCurrentPrice(ticker)
    canIncrementData2 = True
Else
    len_arr = UBound(chart2array) - LBound(chart2array) + 1
    If len_arr = 30 Then
    For i = 0 To len_arr - 2
    chart2array(i) = chart2array(i + 1)
    Next
    chart2array(UBound(chart2array)) = getCurrentPrice(ticker)
    Else
    ReDim Preserve chart2array(0 To UBound(chart2array) + 1)
    chart2array(UBound(chart2array)) = getCurrentPrice(ticker)
    End If

End If



End Sub


Function getCurrentPrice(symbol As String) As Double

Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Dim json As Object
Dim newPrice As Double

On Error GoTo noticker

Url = "https://api.binance.com/api/v3/ticker/price?symbol=" & symbol
xmlhttp.Open "GET", Url, False
xmlhttp.Send

Set json = JsonConverter.ParseJson(xmlhttp.responseText)
getCurrentPrice = CDbl(Replace(json("price"), ".", ","))

Done:  Exit Function

noticker:
    MsgBox "This trading pair '" & symbol & "' is not supported on Binance."



End Function