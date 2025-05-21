Option Explicit

Public Sub CreateSampleDashboard_Jules()
    ' Dimension variables
    Dim pptApp As Object ' PowerPoint.Application
    Dim pptPres As Object ' PowerPoint.Presentation
    Dim pptSld As Object ' PowerPoint.Slide
    Dim championName As String
    Dim championRate As Single
    ' ... other necessary variables for chart data, positions, etc.

    ' Turn off screen updating to speed up macro execution and reduce flicker
    Application.ScreenUpdating = False
    
    ' Error Handling
    On Error GoTo ErrorHandler
    
    ' Initialize Data (Example)
    championName = "Jules"
    championRate = 0.85 ' 85%
    ' ... Initialize data for charts, etc.
    
    ' 1. Create PowerPoint Instance
    On Error Resume Next ' Try to get an existing instance
    Set pptApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo ErrorHandler ' Re-enable default error handling
    pptApp.Visible = True ' Make PowerPoint visible

    ' Add a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Add a slide (e.g., Title Slide)
    Set pptSld = pptPres.Slides.Add(1, 11) ' ppLayoutTitle
    
    ' Call helper subroutines to add elements to the slide
    ' These would be defined in graphics.vba or another module
    ' ApplyDarkGradientBackground pptSld
    ' AddChampionSection pptSld, championName, championRate, 50, 50, 300, 100
    ' AddTrendLineChart pptSld, Range("A1:D5"), "Performance Trend", 50, 200, 400, 250 ' Example data range
    ' AddComparisonBarChart pptSld, Range("E1:F5"), "Category Comparison", 500, 200, 400, 250 ' Example data range

    ' ... (Rest of the slide creation logic for more slides or elements) ...
    
    MsgBox "Sample Dashboard Presentation created successfully!", vbInformation
    
    ' Re-enable screen updating before exiting
    Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
    ' Re-enable screen updating in case of an error
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
    ' ... (Cleanup: e.g., close PowerPoint objects if necessary, but be careful not to close user's other work)
    ' If pptPres Is Nothing And Not pptApp Is Nothing Then
    '    pptApp.Quit
    ' End If
    ' Set pptSld = Nothing
    ' Set pptPres = Nothing
    ' Set pptApp = Nothing
End Sub

Public Sub DefineDummyData(ByRef ws As Object) ' Excel.Worksheet
    ' Define data arrays
    Dim categories(1 To 2) As String
    Dim years(1 To 3) As Long ' Changed to Long for actual years
    Dim successRates(1 To 2, 1 To 3) As Double ' Category, Year

    ' Populate data
    categories(1) = "Alpha"
    categories(2) = "Bravo"

    ' Define years (New values)
    years(1) = 2021
    years(2) = 2022
    years(3) = 2023

    successRates(1, 1) = 0.65 ' Alpha, 2021
    successRates(1, 2) = 0.70 ' Alpha, 2022
    successRates(1, 3) = 0.75 ' Alpha, 2023
    successRates(2, 1) = 0.55 ' Bravo, 2021
    successRates(2, 2) = 0.60 ' Bravo, 2022
    successRates(2, 3) = 0.62 ' Bravo, 2023

    ' Clear previous data
    ws.Cells.ClearContents

    ' Write data to the worksheet
    ' Add headers for years (X-axis categories) - this part is relevant to AddTrendLineChart
    ws.Cells(1, 1).Value = "Year" ' This header can remain "Year"
    Dim j As Long
    For j = LBound(years) To UBound(years)
        ws.Cells(j + 1, 1).Value = years(j) ' Output actual year number, e.g., 2021
    Next j

    ' Add headers for categories (Series names)
    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        ws.Cells(1, i + 1).Value = categories(i)
    Next i

    ' Add success rates
    For i = LBound(categories) To UBound(categories)
        For j = LBound(years) To UBound(years)
            ws.Cells(j + 1, i + 1).Value = successRates(i, j)
        Next j
    Next i
    
    Debug.Print "Dummy data defined on sheet: " & ws.Name
End Sub
