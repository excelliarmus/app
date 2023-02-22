VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   12600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22305
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub btnData_Click()

 'frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

 frmData.Visible = True

End Sub

Sub btnBalances_Click()

 frmData.Visible = False
 'frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

 frmBalances.Visible = True

End Sub





Private Sub btnDataGetChart_Click()

nRows = Sheets("Data").UsedRange.Rows.Count
 
Dim OHLCChart As Chart
Set OHLCChart = Charts.Add
 
With OHLCChart
    .SetSourceData Source:=Sheets("Data").Range("a1:e" & nRows)
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

Sub btnTrading_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 'frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

 frmTrading.Visible = True

End Sub

Sub btnPrediction_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 'frmPrediction.Visible = False
 frmAbout.Visible = False

 frmPrediction.Visible = True

End Sub

Sub btnAbout_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 'frmAbout.Visible = False

 frmAbout.Visible = True

End Sub

Sub btnHome_Click()

 frmData.Visible = False
 frmBalances.Visible = False
 frmTrading.Visible = False
 frmPrediction.Visible = False
 frmAbout.Visible = False

End Sub

Sub btnExit_Click()

MsgBox "Click on 'Delete' for all the next prompts"

Call ModDeleteAllCharts.deleteAll

Unload Me

End Sub


Sub UserForm_Activate()

Call ModMakeUserFormResizable.MakeFormResizable

End Sub
