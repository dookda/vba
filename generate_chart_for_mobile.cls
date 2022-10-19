VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub generate_scatterplot()
Dim ochartObj As ChartObject
Dim oChart As Chart
Dim d As Date, strDate As Date, strDate2 As String, strDate3 As Date
Dim Lr As Long

Dim FirstDate As Date
Dim IntervalType As String
Dim Number As Integer
Dim Msg As String

'#MOBILE auto chart

'find end of cell
Lr = Cells(1, Columns.Count).End(xlToLeft).Column
'MsgBox Lr

'using date function
'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dateadd-function

'import date start, using template "2018/01/"
strDate = Range("A34").Value

'Insert row
Rows(1).Insert
Rows(2).Insert

'create new date
For i = 1 To Lr - 1
    'date generate
    Range(Cells(1, i + 1), Cells(1, i + 1)).Value = DateAdd("d", i - 1, strDate)
    'assign date name
    sDate2 = Format(DateAdd("d", i - 1, strDate), "dddd")
    Range(Cells(2, i + 1), Cells(2, i + 1)).Value = sDate2
Next i

'create chart object
Set ochartObj = ActiveSheet.ChartObjects.Add(Top:=200, Left:=325, Width:=800, Height:=400)
Set oChart = ochartObj.Chart

'select chart type : Scatter
oChart.ChartType = xlXYScatter
oChart.HasLegend = False

'loop date
For i = 2 To Lr
    oChart.SeriesCollection.NewSeries
    oChart.SeriesCollection(i - 1).XValues = Range(Cells(5, 1), Cells(28, 1))
    oChart.SeriesCollection(i - 1).Values = Range(Cells(5, i), Cells(28, i))
       
    'loop hours
    For x = 1 To 24
        oChart.SeriesCollection(i - 1).Points(x).MarkerStyle = xlCircle
        oChart.SeriesCollection(i - 1).Points(x).MarkerSize = 8
        
        'assign color to date type
        If Range(Cells(2, i), Cells(2, i)).Value = "Friday" Then
            oChart.SeriesCollection(i - 1).Points(x).MarkerBackgroundColor = RGB(112, 173, 71)
            oChart.SeriesCollection(i - 1).Points(x).MarkerForegroundColor = RGB(112, 173, 71)
        ElseIf Range(Cells(2, i), Cells(2, i)).Value = "Saturday" Or Range(Cells(2, i), Cells(2, i)).Value = "Sunday" Then
            oChart.SeriesCollection(i - 1).Points(x).MarkerBackgroundColor = RGB(255, 0, 0)
            oChart.SeriesCollection(i - 1).Points(x).MarkerForegroundColor = RGB(255, 0, 0)
        Else
            oChart.SeriesCollection(i - 1).Points(x).MarkerBackgroundColor = RGB(68, 114, 196)
            oChart.SeriesCollection(i - 1).Points(x).MarkerForegroundColor = RGB(68, 114, 196)
        End If
        
    Next x
    
Next i

'create new series for MEAN
oChart.SeriesCollection.NewSeries

'defined chart type and color
oChart.SeriesCollection(Lr).ChartType = xlXYScatterLines
oChart.SeriesCollection(Lr).XValues = Range(Cells(5, 1), Cells(28, 1))
oChart.SeriesCollection(Lr).Values = Range(Cells(5, Lr + 1), Cells(28, Lr + 1))
oChart.SeriesCollection(Lr).Format.Line.ForeColor.RGB = RGB(132, 60, 12)

'assign color and size for mean symbol
For x = 1 To 24
    oChart.SeriesCollection(Lr).Points(x).MarkerStyle = xlCircle
    oChart.SeriesCollection(Lr).Points(x).MarkerSize = 8
    oChart.SeriesCollection(Lr).Points(x).MarkerBackgroundColor = RGB(132, 60, 12)
    oChart.SeriesCollection(Lr).Points(x).MarkerForegroundColor = RGB(132, 60, 12)
Next x

'set chart title
oChart.HasTitle = True
oChart.ChartTitle.Text = Range("A37").Value
oChart.ChartTitle.Font.Size = 20
oChart.ChartTitle.Font.Name = "TH SarabunPSK"

'set y axes
oChart.Axes(xlCategory).HasTitle = True
oChart.Axes(xlCategory).AxisTitle.Caption = Range("A39").Value
oChart.Axes(xlCategory).AxisTitle.Font.Size = 20
oChart.Axes(xlCategory).AxisTitle.Font.Name = "TH SarabunPSK"
oChart.Axes(xlCategory).MaximumScale = 25
oChart.Axes(xlCategory).MajorUnit = 1
oChart.Axes(xlCategory).HasMajorGridlines = True

'set x axes
oChart.Axes(xlValue).HasTitle = True
oChart.Axes(xlValue).AxisTitle.Caption = Range("A38").Value
oChart.Axes(xlValue).AxisTitle.Font.Size = 20
oChart.Axes(xlValue).AxisTitle.Font.Name = "TH SarabunPSK"
'oChart.Axes(xlValue).TickLabels.Font.Size = 16
'oChart.Axes(xlValue).TickLabels.Font.Name = "TH SarabunPSK"

'set labelto symbol
'oChart.SeriesCollection(1).HasDataLabels = True

End Sub

