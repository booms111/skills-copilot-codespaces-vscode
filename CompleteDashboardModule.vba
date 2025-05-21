Option Explicit

' Constants for Fonts
Private Const FONT_MAIN As String = "Aptos Display"
Private Const FONT_FALLBACK As String = "Calibri"
Private Const FONT_TITLE As String = FONT_MAIN
Private Const FONT_CHAMPION As String = FONT_MAIN
Private Const FONT_CHART_AXIS As String = "Calibri"

' Constants for Colors
Private Const Color_WHITE As Long = 16777215  'RGB(255, 255, 255)
Private Const COLOR_GOLD As Long = 55295 'RGB(255, 215, 0)
Private Const COLOR_DARK_BLUE As Long = 139 'RGB(0, 32, 96)
Private Const COLOR_BLACK As Long = 0 ' RGB(0, 0, 0)
Private Const COLOR_LIGHT_GREY As Long = 13882323 'RGB(220, 220, 220)
Private Const COLOR_CHAMPION_BOX_FILL As Long = RGB(50, 50, 50)
Private Const COLOR_CHART_GRIDLINES As Long = RGB(80, 80, 80)

' --- Main Subroutine ---
Public Sub CreateSampleDashboard_Jules()
    ' Dimension variables
    Dim pptApp As Object ' PowerPoint.Application
    Dim pptPres As Object ' PowerPoint.Presentation
    Dim pptSld As Object ' PowerPoint.Slide
    
    Dim excelApp As Object ' Excel.Application
    Dim wb As Object       ' Excel.Workbook
    Dim wsChartData As Object ' Excel.Worksheet
    
    Dim championName As String
    Dim championRate As Single
    Dim overallSuccessRates() As Double ' Will be ReDim'd in DefineDummyData
    Dim categories() As String      ' Will be ReDim'd in DefineDummyData
    Dim numCategories As Long
    
    ' Turn off screen updating to speed up macro execution and reduce flicker
    Application.ScreenUpdating = False
    
    ' Error Handling
    On Error GoTo ErrorHandler
    
    ' 1. Create or Get Excel Instance for data handling
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    excelApp.Visible = False ' Keep Excel hidden for data manipulation
    On Error GoTo ErrorHandler
    
    ' Add a workbook and a worksheet for chart data
    Set wb = excelApp.Workbooks.Add
    Set wsChartData = wb.Worksheets(1)
    wsChartData.Name = "ChartData"
    
    ' 2. Define Dummy Data (populates wsChartData and provides rates for IdentifyChampion)
    DefineDummyData wsChartData, categories, overallSuccessRates ' categories & overallSuccessRates are Out parameters
    numCategories = UBound(categories) - LBound(categories) + 1
                                                    
    ' 3. Identify Champion
    championName = IdentifyChampion(categories, overallSuccessRates)
    ' Find the champion's rate
    Dim cIndex As Long
    For cIndex = LBound(categories) To UBound(categories)
        If categories(cIndex) = championName Then
            championRate = overallSuccessRates(cIndex)
            Exit For
        End If
    Next cIndex

    ' 4. Create PowerPoint Instance
    Dim pptInstanceCreated As Boolean
    pptInstanceCreated = False

    On Error Resume Next ' Allow GetObject to fail silently
    Set pptApp = GetObject(, "PowerPoint.Application")
    If Err.Number = 0 Then
        pptInstanceCreated = True ' Successfully got existing instance
        Err.Clear ' Clear potential Err object state from successful GetObject
    Else
        Err.Clear ' Clear error from GetObject before trying CreateObject
        ' Attempt to create a new instance
        Set pptApp = CreateObject("PowerPoint.Application")
        If Err.Number = 0 Then
            pptInstanceCreated = True ' Successfully created new instance
            Err.Clear ' Clear potential Err object state from successful CreateObject
        Else
            ' Both GetObject and CreateObject failed
            Dim errMsg As String
            errMsg = "Failed to initialize PowerPoint application (Error " & Err.Number & ": " & Err.Description & ")." & vbCrLf & _
                   "Please ensure PowerPoint is installed, activated, and not in a problematic state (e.g., no startup dialogs open)."
            MsgBox errMsg, vbCritical, "PowerPoint Initialization Error"
            Err.Clear
            GoTo Cleanup ' Assuming Cleanup handles exiting and cleaning other objects
        End If
    End If
    On Error GoTo ErrorHandler ' Restore main error handler for subsequent operations

    If Not pptInstanceCreated Or pptApp Is Nothing Then ' Defensive check
        MsgBox "PowerPoint application object could not be confirmed. Aborting dashboard creation.", vbCritical, "PowerPoint Object Error"
        GoTo Cleanup 
    End If
    
    ' At this point, pptApp should be a valid object.
    ' Make PowerPoint visible and bring to front
    pptApp.Visible = True
    ' On Error Resume Next ' Optional: Activate may fail if another app is modal
    ' pptApp.Activate ' Optional: Tries to bring PowerPoint to the foreground
    ' On Error GoTo ErrorHandler

    ' Add a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Add a slide
    Set pptSld = pptPres.Slides.Add(1, 12) ' 12 = ppLayoutBlank (PowerPoint.PpSlideLayout.ppLayoutBlank)
    
    ' 5. Apply Background and Add Elements
    ApplyDarkGradientBackground pptSld
    AddTitle pptSld, "Jules's Performance Dashboard"
    
    AddChampionSection pptSld, championName, championRate, 50, 100, 300, 80
    
    ' Data for TrendLineChart: Header in row 1. 3 Years of data means 3 data rows. Total 4 rows.
    ' "Year" column + numCategories columns.
    Dim trendChartRange As String
    trendChartRange = wsChartData.Range("A1").Resize(3 + 1, numCategories + 1).Address(External:=False)
    AddTrendLineChart pptSld, wsChartData, trendChartRange, "Performance Trend (Success Rate %)", 50, 200, 600, 350
    
    ' Data for ComparisonBarChart
    wsChartData.Cells(1, numCategories + 3).Value = "Category" ' Starting in a new area, e.g., Col E if numCategories = 2
    wsChartData.Cells(1, numCategories + 4).Value = "Average Success Rate"
    For cIndex = LBound(categories) To UBound(categories)
        wsChartData.Cells(cIndex - LBound(categories) + 2, numCategories + 3).Value = categories(cIndex)
        wsChartData.Cells(cIndex - LBound(categories) + 2, numCategories + 4).Value = overallSuccessRates(cIndex)
    Next cIndex
    
    Dim barChartRange As String
    barChartRange = wsChartData.Cells(1, numCategories + 3).Resize(numCategories + 1, 2).Address(External:=False)
    AddComparisonBarChart pptSld, wsChartData, barChartRange, "Overall Success Rate Comparison", 700, 200, 400, 350
        
    MsgBox "Sample Dashboard Presentation created successfully!", vbInformation
    
Cleanup:
    Application.ScreenUpdating = True ' Re-enable screen updating before exiting
    
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not excelApp Is Nothing Then excelApp.Quit
    Set wsChartData = Nothing: Set wb = Nothing: Set excelApp = Nothing
    Set pptSld = Nothing: Set pptPres = Nothing: Set pptApp = Nothing
Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description & vbCrLf & "Error " & Err.Number, vbCritical
    Resume Cleanup
End Sub

' --- Data Helper Subroutines ---
Public Sub DefineDummyData(ByRef ws As Object, ByRef outCategories() As String, ByRef outOverallRates() As Double)
    Dim localCategories(1 To 2) As String
    Dim years(1 To 3) As Long
    Dim successRates(1 To 2, 1 To 3) As Double ' Category, Year
    Dim i As Long, j As Long, numYears As Long

    localCategories(1) = "Alpha"
    localCategories(2) = "Bravo"
    ReDim outCategories(LBound(localCategories) To UBound(localCategories))
    For i = LBound(localCategories) To UBound(localCategories)
        outCategories(i) = localCategories(i)
    Next i
    
    years(1) = 2021: years(2) = 2022: years(3) = 2023
    numYears = UBound(years) - LBound(years) + 1

    successRates(1, 1) = 0.65: successRates(1, 2) = 0.70: successRates(1, 3) = 0.75 ' Alpha
    successRates(2, 1) = 0.55: successRates(2, 2) = 0.60: successRates(2, 3) = 0.62 ' Bravo

    ws.Cells.ClearContents
    ws.Cells(1, 1).Value = "Year"
    For j = LBound(years) To UBound(years)
        ws.Cells(j - LBound(years) + 2, 1).Value = years(j)
    Next j

    For i = LBound(localCategories) To UBound(localCategories)
        ws.Cells(1, i - LBound(localCategories) + 2).Value = localCategories(i)
    Next i

    For i = LBound(localCategories) To UBound(localCategories)
        For j = LBound(years) To UBound(years)
            ws.Cells(j - LBound(years) + 2, i - LBound(localCategories) + 2).Value = successRates(i, j)
        Next j
    Next i
    
    ReDim outOverallRates(LBound(localCategories) To UBound(localCategories))
    For i = LBound(localCategories) To UBound(localCategories)
        Dim sumRate As Double: sumRate = 0
        For j = LBound(years) To UBound(years)
            sumRate = sumRate + successRates(i, j)
        Next j
        outOverallRates(i) = sumRate / numYears
    Next i
    Debug.Print "Dummy data defined on sheet: " & ws.Name
End Sub

Public Function IdentifyChampion(categories() As String, rates() As Double) As String
    Dim i As Long, maxRate As Double, championIndex As Long
    If LBound(rates) > UBound(rates) Then IdentifyChampion = "N/A": Exit Function
    
    maxRate = rates(LBound(rates)): championIndex = LBound(rates)
    For i = LBound(rates) + 1 To UBound(rates)
        If rates(i) > maxRate Then maxRate = rates(i): championIndex = i
    Next i
    IdentifyChampion = categories(championIndex)
End Function

Public Function CalculateSuccessRates(param1 As Variant, param2 As Variant) As Double()
    Dim placeholderRates(1 To 1) As Double: placeholderRates(1) = 0.77
    CalculateSuccessRates = placeholderRates
    Debug.Print "CalculateSuccessRates called (placeholder - not directly used by this dashboard's current data flow)"
End Function

Public Function CalculateAverageSuccessRates(successRatesArray() As Double) As Double
    Dim sumVal As Double, countVal As Long, i As Long: sumVal = 0: countVal = 0
    If LBound(successRatesArray) <= UBound(successRatesArray) Then
        For i = LBound(successRatesArray) To UBound(successRatesArray)
            sumVal = sumVal + successRatesArray(i): countVal = countVal + 1
        Next i
        If countVal > 0 Then CalculateAverageSuccessRates = sumVal / countVal Else CalculateAverageSuccessRates = 0
    Else
        CalculateAverageSuccessRates = 0
    End If
    Debug.Print "CalculateAverageSuccessRates called (placeholder - direct average used in DefineDummyData)"
End Function

' --- Graphical Helper Subroutines ---
Private Sub ApplyDarkGradientBackground(ByRef sld As Object) ' PowerPoint.Slide
    With sld.FollowMasterBackground: .Visible = False: End With ' 0 = msoFalse
    With sld.Background.Fill
        .Visible = -1 ' -1 = msoTrue
        .TwoColorGradient 1, 1 ' Style:=msoGradientLinear (1), Variant:=1
        .BackColor.RGB = COLOR_DARK_BLUE
        .ForeColor.RGB = COLOR_BLACK
    End With
    Debug.Print "Background applied."
End Sub

Public Sub AddTitle(ByRef sld As Object, ByVal titleText As String)
    Dim shpTitle As Object ' PowerPoint.Shape
    Set shpTitle = sld.Shapes.AddTextbox(1, 50, 20, 860, 50) ' 1 = msoTextOrientationHorizontal
    With shpTitle.TextFrame2.TextRange
        .Text = titleText: .Font.Name = FONT_TITLE: .Font.Size = 32: .Font.Bold = -1 ' -1 = msoTrue
        .Font.Fill.ForeColor.RGB = Color_WHITE: .ParagraphFormat.Alignment = 2 ' 2 = ppAlignCenter
    End With
    Debug.Print "Title added: " & titleText
End Sub

Private Sub AddChampionSection(ByRef sld As Object, ByVal champName As String, ByVal champRate As Single, _
                               ByVal textLeftPosition As Single, ByVal textTopPosition As Single, _
                               ByVal boxWidth As Single, ByVal boxHeight As Single)
    Dim shpChampionBox As Object ' PowerPoint.Shape
    Set shpChampionBox = sld.Shapes.AddShape(1, textLeftPosition, textTopPosition, boxWidth, boxHeight) ' 1 = msoShapeRoundedRectangle
    shpChampionBox.Adjustments(1) = 0.15
    
    With shpChampionBox
        .Fill.Visible = -1: .Fill.ForeColor.RGB = COLOR_CHAMPION_BOX_FILL
        .Line.Visible = -1: .Line.ForeColor.RGB = COLOR_GOLD: .Line.Weight = 1.5
        With .TextFrame2
            .MarginLeft = 10: .MarginRight = 10: .VerticalAnchor = 3 ' 3 = msoAnchorMiddle
            With .TextRange
                .Text = "CHAMPION: " & UCase(champName) & vbCrLf & "Overall Success Rate: " & Format(champRate, "0.00%")
                .Font.Name = FONT_CHAMPION: .Font.Size = 16: .Font.Fill.ForeColor.RGB = Color_WHITE
                .ParagraphFormat.Alignment = 2 ' ppAlignCenter
                If Len(champName) > 0 Then
                    Dim startPos As Long: startPos = InStr(.Text, UCase(champName))
                    If startPos > 0 Then
                         With .Characters(startPos, Len(champName)).Font
                            .Fill.ForeColor.RGB = COLOR_GOLD: .Bold = -1
                         End With
                    End If
                End If
            End With
        End With
    End With
    Debug.Print "Champion section added for " & champName
End Sub

Private Sub AddTrendLineChart(ByRef sld As Object, ByRef chartDataSheet As Object, ByVal chartDataSource As String, ByVal titleText As String, _
                             ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim chtObj As Object, cht As Object, seriesIndex As Long, numSeries As Long
    Set chtObj = sld.Shapes.AddChart2(-1, 74, leftPos, topPos, widthVal, heightVal) ' 74 = xlLineMarkers
    Set cht = chtObj.Chart
    cht.SetSourceData Source:=chartDataSheet.Range(chartDataSource).Address(External:=True), PlotBy:=2 ' 2 = xlColumns for typical time series
    numSeries = cht.SeriesCollection.Count

    cht.HasTitle = True
    With cht.ChartTitle
        .Text = titleText: .Format.TextFrame2.TextRange.Font.Name = FONT_TITLE
        .Format.TextFrame2.TextRange.Font.Size = 16: .Format.TextFrame2.TextRange.Font.Bold = -1
        .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color_WHITE
    End With

    With cht.Axes(1) ' xlCategory / X-axis
        .HasTitle = False
        With .Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 10: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
    End With
    With cht.Axes(2) ' xlValue / Y-axis
        .HasTitle = True: .AxisTitle.Text = "Success Rate"
        With .AxisTitle.Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 12: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .TickLabels.NumberFormat = "0%": .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
        With .Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 10: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_CHART_GRIDLINES
    End With
    
    Dim seriesColors(1 To 2) As Long
    seriesColors(1) = RGB(0, 176, 240): seriesColors(2) = RGB(255, 192, 0)
    For seriesIndex = 1 To numSeries
        If seriesIndex <= UBound(seriesColors) Then
            With cht.SeriesCollection(seriesIndex)
                .Format.Line.ForeColor.RGB = seriesColors(seriesIndex): .Format.Line.Weight = 2.25
                .MarkerStyle = 8: .MarkerSize = 7 ' 8 = xlMarkerStyleCircle
                .MarkerBackgroundColor = seriesColors(seriesIndex): .MarkerForegroundColor = Color_WHITE
            End With
        End If
    Next seriesIndex
    
    cht.HasLegend = True
    With cht.Legend: .Position = -4107 ' xlLegendPositionBottom
        With .Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 10: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .Format.Fill.Visible = False: End With
    With cht.ChartArea: .Format.Fill.Visible = False: .Format.Line.Visible = False: End With
    Debug.Print "Trendline chart '" & titleText & "' added."
End Sub

Private Sub AddComparisonBarChart(ByRef sld As Object, ByRef chartDataSheet As Object, ByVal chartDataSource As String, ByVal titleText As String, _
                                 ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim chtObj As Object, cht As Object, numPoints As Long
    Set chtObj = sld.Shapes.AddChart2(-1, 51, leftPos, topPos, widthVal, heightVal) ' 51 = xlColumnClustered
    Set cht = chtObj.Chart
    cht.SetSourceData Source:=chartDataSheet.Range(chartDataSource).Address(External:=True), PlotBy:=2 ' 2 = xlColumns

    cht.HasTitle = True
    With cht.ChartTitle
        .Text = titleText: .Format.TextFrame2.TextRange.Font.Name = FONT_TITLE
        .Format.TextFrame2.TextRange.Font.Size = 16: .Format.TextFrame2.TextRange.Font.Bold = -1
        .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color_WHITE
    End With

    With cht.Axes(1) ' xlCategory / X-axis
        .HasTitle = False
        With .Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 10: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
    End With
    With cht.Axes(2) ' xlValue / Y-axis
        .HasTitle = True: .AxisTitle.Text = "Avg. Success Rate"
        With .AxisTitle.Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 12: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .TickLabels.NumberFormat = "0%": .Format.Line.ForeColor.RGB = COLOR_LIGHT_GREY
        With .Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 10: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_CHART_GRIDLINES
    End With
    
    Dim barColors(1 To 2) As Long: barColors(1) = RGB(0, 176, 240): barColors(2) = RGB(255, 192, 0)
    If cht.SeriesCollection.Count > 0 Then
        numPoints = cht.SeriesCollection(1).Points.Count
        If numPoints > 0 Then
             For i = 1 To numPoints
                 If i <= UBound(barColors) Then cht.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = barColors(i)
             Next i
        End If
        With cht.SeriesCollection(1)
            .HasDataLabels = True
            With .DataLabels
                .Position = -4142 ' xlLabelPositionOutsideEnd
                .NumberFormat = "0.0%"
                With .Format.TextFrame2.TextRange.Font: .Name = FONT_CHART_AXIS: .Size = 9: .Fill.ForeColor.RGB = COLOR_LIGHT_GREY: End With
            End With
        End With
    End If
    If cht.ChartGroups.Count > 0 Then cht.ChartGroups(1).GapWidth = 100
    cht.HasLegend = False
    With cht.ChartArea: .Format.Fill.Visible = False: .Format.Line.Visible = False: End With
    Debug.Print "Comparison bar chart '" & titleText & "' added."
End Sub
