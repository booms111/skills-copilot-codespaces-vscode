Option Explicit

Private Sub ApplyDarkGradientBackground(ByRef sld As Object) ' PowerPoint.Slide
    ' Ensure constants COLOR_DARK_BLUE and COLOR_BLACK are accessible
    ' (e.g., defined as Public Const in constants.vba or this sub is in that module)

    With sld.FollowMasterBackground
        .Visible = msoFalse ' Do not follow master background
    End With

    With sld.Background.Fill
        .Visible = msoTrue
        ' Apply a two-color linear gradient
        .TwoColorGradient Style:=msoGradientLinear, Variant:=1 ' Variant 1 is diagonal from top-left to bottom-right
        ' Set the start and end colors for the gradient
        .BackColor.RGB = COLOR_DARK_BLUE ' Start color (top-left)
        .ForeColor.RGB = COLOR_BLACK    ' End color (bottom-right)
    End With

    Debug.Print "Background applied."
End Sub

Private Sub AddChampionSection(ByRef sld As Object, ByVal champName As String, ByVal champRate As Single, _
                               ByVal textLeftPosition As Single, ByVal textTopPosition As Single, _
                               ByVal boxWidth As Single, ByVal boxHeight As Single)
    ' Ensure constants COLOR_CHAMPION_BOX_FILL, COLOR_GOLD, Color_WHITE, FONT_CHAMPION are accessible
    
    Dim shpChampionBox As Object ' PowerPoint.Shape
    
    ' Create the rounded rectangle shape for the champion box
    Set shpChampionBox = sld.Shapes.AddShape(msoShapeRoundedRectangle, textLeftPosition, textTopPosition, boxWidth, boxHeight)
    shpChampionBox.Adjustments(1) = 0.15 ' Set corner radius
    
    With shpChampionBox
        ' Style the shape
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = COLOR_CHAMPION_BOX_FILL ' Use defined constant
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = COLOR_GOLD
        .Line.Weight = 1.5 ' Adjusted line weight
        
        ' Add text to the shape
        With .TextFrame2
            .MarginLeft = 10 ' Add left margin
            .MarginRight = 10 ' Add right margin
            With .TextRange
                .Text = "CHAMPION: " & UCase(champName) & vbCrLf & _
                        "Overall Success Rate: " & Format(champRate, "0.00%")
                .Font.Name = FONT_CHAMPION
                .Font.Size = 18
                .Font.Fill.ForeColor.RGB = Color_WHITE
                .ParagraphFormat.Alignment = msoAlignCenter

                ' Specifically color the Champion Name
                ' Placeholder for coloring logic:
                ' If Len(champName) > 0 Then
                '    Dim startPos As Long
                '    startPos = InStr(.Text, UCase(champName))
                '    If startPos > 0 Then
                '        .Characters(startPos, Len(champName)).Font.Fill.ForeColor.RGB = COLOR_GOLD
                '    End If
                ' End If
            End With
        End With
    End With
    Debug.Print "Champion section added for " & champName
End Sub

Private Sub AddTrendLineChart(ByRef sld As Object, ByVal chartDataSheet As Object, ByVal chartDataSource As String, ByVal titleText As String, _
                             ByVal leftPos As Single, ByVal topPos As Single, _
                             ByVal widthVal As Single, ByVal heightVal As Single)
    ' Assumes chartDataSheet is an Excel.Worksheet object and chartDataSource is a String like "A1:C4"
    ' Ensure constants FONT_CHART_AXIS, COLOR_LIGHT_GREY, COLOR_CHART_GRIDLINES, Color_WHITE, FONT_TITLE are accessible

    Dim chtObj As Object ' PowerPoint.ChartObject
    Dim cht As Object    ' PowerPoint.Chart
    Dim seriesIndex As Long

    ' Add chart to the slide
    Set chtObj = sld.Shapes.AddChart2(Style:=-1, Type:=xlLineMarkers, _
                                     Left:=leftPos, Top:=topPos, _
                                     Width:=widthVal, Height:=heightVal)
    Set cht = chtObj.Chart

    ' Set chart data
    cht.SetSourceData Source:=chartDataSheet.Range(chartDataSource).Address(External:=True), PlotBy:=xlRows

    ' Chart Title
    cht.HasTitle = True
    With cht.ChartTitle
        .Text = titleText
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_TITLE ' Assuming FONT_TITLE is defined
            .Size = 18
            .Bold = True
            .Fill.ForeColor.RGB = Color_WHITE
        End With
    End With

    ' Format Axes
    With cht.Axes(xlCategory) ' X-axis
        .HasTitle = False
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_CHART_AXIS
            .Size = 10
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
    End With
    
    With cht.Axes(xlValue) ' Y-axis
        .HasTitle = True ' Set Y-axis title
        .AxisTitle.Text = "Success Rate" ' Set Y-axis title text
        With .AxisTitle.Format.TextFrame2.TextRange.Font ' Format Y-axis title font
            .Name = FONT_CHART_AXIS
            .Size = 12 ' Adjusted size
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .TickLabels.NumberFormat = "0%" 
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_CHART_AXIS
            .Size = 10
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_CHART_GRIDLINES ' Use constant for gridlines
    End With
    
    ' Format Series lines and markers (modern style)
    Dim seriesColor(1 To 4) As Long
    seriesColor(1) = RGB(0, 176, 240)  ' Blue
    seriesColor(2) = RGB(255, 192, 0)  ' Orange/Gold
    seriesColor(3) = RGB(146, 208, 80)  ' Green
    seriesColor(4) = RGB(112, 48, 160)   ' Purple

    ' Ensure there are series to format before looping
    If cht.SeriesCollection.Count > 0 Then
        For seriesIndex = 1 To cht.SeriesCollection.Count
            With cht.SeriesCollection(seriesIndex)
                ' Cycle through colors if more series than defined colors
                .Format.Line.ForeColor.RGB = seriesColor((seriesIndex - 1) Mod UBound(seriesColor) + 1)
                .Format.Line.Weight = 2.25 ' Adjusted line weight
                .MarkerStyle = xlMarkerStyleCircle 
                .MarkerSize = 7
                .MarkerBackgroundColor = seriesColor((seriesIndex - 1) Mod UBound(seriesColor) + 1)
                .MarkerForegroundColor = Color_WHITE
            End With
        Next seriesIndex
    End If
    
    ' Legend
    cht.HasLegend = True
    With cht.Legend
        .Position = xlLegendPositionBottom
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_CHART_AXIS
            .Size = 10
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .Format.Fill.Visible = msoFalse ' No background for legend
    End With
    
    ' Chart Area Formatting
    With cht.ChartArea
        .Format.Fill.Visible = msoFalse ' Transparent background
        .Format.Line.Visible = msoFalse ' No border around chart area
    End With
    
    Debug.Print "Trendline chart '" & titleText & "' added and formatted."
End Sub

Private Sub AddComparisonBarChart(ByRef sld As Object, ByVal chartData As Object, ByVal titleText As String, _
                                 ByVal leftPos As Single, ByVal topPos As Single, _
                                 ByVal widthVal As Single, ByVal heightVal As Single)
    ' Assumes chartData is a Range object from Excel, or similar data structure for a bar chart
    ' Ensure constants FONT_CHART_AXIS, COLOR_LIGHT_GREY, COLOR_CHART_GRIDLINES, Color_WHITE, FONT_TITLE are accessible

    Dim chtObj As Object ' PowerPoint.ChartObject
    Dim cht As Object    ' PowerPoint.Chart
    Dim i As Long        ' Loop counter for points

    ' Add chart to the slide
    Set chtObj = sld.Shapes.AddChart2(Style:=-1, Type:=xlColumnClustered, _
                                     Left:=leftPos, Top:=topPos, _
                                     Width:=widthVal, Height:=heightVal)
    Set cht = chtObj.Chart

    ' Set chart data (placeholder - actual implementation depends on chartData)
    ' cht.SetSourceData Source:=chartData.Address(External:=True) 
    ' For demonstration, assume 1 series with 4 points
    ' cht.ChartData.Workbook.Worksheets(1).Range("A1:B5").Value = ... 

    ' Chart Title
    cht.HasTitle = True
    With cht.ChartTitle
        .Text = titleText
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_TITLE ' Assuming FONT_TITLE is defined
            .Size = 18
            .Bold = True
            .Fill.ForeColor.RGB = Color_WHITE
        End With
    End With

    ' Format Axes
    With cht.Axes(xlCategory) ' X-axis
        .HasTitle = False ' Assuming no X-axis title for this chart type
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_CHART_AXIS
            .Size = 10
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
    End With
    
    With cht.Axes(xlValue) ' Y-axis
        .HasTitle = True ' Set Y-axis title
        .AxisTitle.Text = "Avg. Success Rate" ' Set Y-axis title text
        With .AxisTitle.Format.TextFrame2.TextRange.Font ' Format Y-axis title font
            .Name = FONT_CHART_AXIS
            .Size = 12 ' Consistent size
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .TickLabels.NumberFormat = "0%" 
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_CHART_AXIS
            .Size = 10
            .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
        End With
        .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY ' Axis line color
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_CHART_GRIDLINES ' Use constant for gridlines
    End With
    
    ' Format Series (Bars)
    ' Assuming we are working with the first series for coloring and data labels
    If cht.SeriesCollection.Count > 0 Then
        With cht.SeriesCollection(1) 
            ' Assign distinct colors to each bar (point)
            ' Ensure there are enough points if hardcoding colors like this
            If .Points.Count >= 1 Then .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 240)  ' Blue
            If .Points.Count >= 2 Then .Points(2).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)  ' Orange/Gold
            If .Points.Count >= 3 Then .Points(3).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)  ' Green
            If .Points.Count >= 4 Then .Points(4).Format.Fill.ForeColor.RGB = RGB(112, 48, 160) ' Purple
            ' For more points, a loop or different color strategy would be needed

            ' Add Data Labels to bars
            .HasDataLabels = True
            With .DataLabels
                .Position = xlLabelPositionOutsideEnd
                .NumberFormat = "0.0%" ' Example format for data labels
                With .Format.TextFrame2.TextRange.Font
                    .Name = FONT_CHART_AXIS
                    .Size = 9
                    .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
                End With
            End With
        End With ' End With cht.SeriesCollection(1)
    End If ' End If cht.SeriesCollection.Count > 0

    ' Adjust GapWidth for the chart group
    If cht.ChartGroups.Count > 0 Then
        cht.ChartGroups(1).GapWidth = 100
    End If
    
    ' Legend
    cht.HasLegend = False ' Typically bar charts with distinct point colors might not need a legend
                          ' Or set to True and format as needed:
    ' cht.HasLegend = True
    ' With cht.Legend
    '     .Position = xlLegendPositionBottom
    '     With .Format.TextFrame2.TextRange.Font
    '         .Name = FONT_CHART_AXIS
    '         .Size = 10
    '         .Fill.ForeColor.RGB = COLOR_LIGHT_GREY
    '     End With
    '     .Format.Fill.Visible = msoFalse 
    ' End With
    
    ' Chart Area Formatting
    With cht.ChartArea
        .Format.Fill.Visible = msoFalse ' Transparent background
        .Format.Line.Visible = msoFalse ' No border around chart area
    End With
    
    Debug.Print "Comparison bar chart '" & titleText & "' added and formatted."
End Sub
