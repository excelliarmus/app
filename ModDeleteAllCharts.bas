Attribute VB_Name = "ModDeleteAllCharts"
'DO NOT DELETE THIS FILE !

' This code deletes all charts in the workbook
' as a chart
' is generated after each refresh...

Sub deleteAll()
    Dim oChart As Chart
    For Each oChart In Application.Charts
        ' Debug.Print oChart.Name
        oChart.Delete
    Next oChart
End Sub

