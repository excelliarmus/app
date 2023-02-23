Attribute VB_Name = "ModData"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public isDataStream1On As Boolean

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

Sub activateDataStream1()
    isDataStream1On = True
End Sub

Sub desactivateDataStream1()
    isDataStream1On = False
End Sub
