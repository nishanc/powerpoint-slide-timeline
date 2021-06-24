Attribute VB_Name = "Module1"
Option Explicit
Sub Timeline()
    '------All variables------'
    Dim ppPres As PowerPoint.Presentation
    Dim ppApp As PowerPoint.Application
    Dim ppSlide As PowerPoint.Slide
    Dim slidesCount As Long
    Dim tableShape As Shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim activeSlide As Single
    Dim x As Long
    Dim past As Long
    Dim present As Long
    Dim future As Long
    Dim borders As Long
    
    '------Theme colors------'
    'Adjust these to match your power point theme
    past = RGB(165, 255, 250)
    present = RGB(0, 255, 205)
    future = RGB(2, 69, 173)
    borders = RGB(7, 32, 69)
    
    '------Get application instance------'
    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    
    If Not ppApp Is Nothing Then                            ' PowerPoint is already running
        Set ppPres = ppApp.ActivePresentation               ' use current presentation
        If ppPres Is Nothing Then                           ' if no presentation there
            Set ppPres = ppApp.Presentations.Open("...")    ' open it
        End If
    Else                                                    ' new PowerPoint instance necessary
        Set ppApp = New PowerPoint.Application              ' start new instance
        Set ppPres = ppApp.Presentations.Open("...")        ' open presentation
    End If
    
    '------Get slide width and height------'
    With ActivePresentation.PageSetup
        slideHeight = .slideHeight
        slideWidth = .slideWidth
    End With
    
    '------Set visible and activate------'
    ppApp.Visible = True
    ppApp.Activate
    
    '------Get slides count------'
    slidesCount = ppPres.Slides.Count

    '------Only do for active slide------'
    'If ppApp.ActiveWindow.Selection.Type = ppSelectionSlides Then
        'Set ppSlide = ppApp.ActiveWindow.Selection.SlideRange(1)
        '' or Set ppSlide = ppApp.ActiveWindow.View.Slide
    'End If
    'Debug.Print ppSlide.SlideID, ppSlide.SlideNumber, ppSlide.SlideIndex
    
    '------For each slide in presentation------'
    For Each ppSlide In ppPres.Slides
        'Create table with 1 row columns = number of slides on the bottom of the slide
        'https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtable
        Set tableShape = ppPres.Slides(ppSlide.SlideIndex).Shapes.AddTable(1, slidesCount, 0, slideHeight - 6, slideWidth, 20)
        '------Set styles for all borders/ cells------'
        With tableShape.Table
            For x = 1 To .Columns.Count
                .Cell(1, x).Shape.Fill.ForeColor.RGB = future
                .Columns(x).Cells.borders(ppBorderLeft).Transparency = 0
                .Columns(x).Cells.borders(ppBorderLeft).Weight = 4
                .Columns(x).Cells.borders(ppBorderLeft).ForeColor.RGB = borders
                .Columns(x).Cells.borders(ppBorderTop).ForeColor.RGB = borders
            Next
        End With
        '------Set styles for cells corresponding to previous progress------'
        If ppSlide.SlideIndex > 2 Then '2 because we are adding a different style to the cell before currunt cell
            With tableShape.Table
                For x = 1 To (ppSlide.SlideIndex - 2)
                    .Cell(1, x).Shape.Fill.ForeColor.RGB = past
                    .Cell(1, x).Shape.Fill.Transparency = 0.5
                Next
            End With
        End If
        '------Set styles related to currunt slide------'
        With tableShape.Table
            .Cell(1, ppSlide.SlideIndex).Shape.Fill.ForeColor.RGB = present
            .Cell(1, ppSlide.SlideIndex).borders(ppBorderTop).ForeColor.RGB = borders
            .Cell(1, ppSlide.SlideIndex).borders(ppBorderLeft).ForeColor.RGB = present
            .Cell(1, ppSlide.SlideIndex).borders(ppBorderRight).ForeColor.RGB = present
            .Cell(1, ppSlide.SlideIndex).borders(ppBorderTop).Weight = 4
        End With
        '------Set styles related to cells corresponding to before slide and after slide------'
        If ppSlide.SlideIndex > 1 Then
            With tableShape.Table
                .Cell(1, ppSlide.SlideIndex - 1).Shape.Fill.ForeColor.RGB = present
                .Cell(1, ppSlide.SlideIndex - 1).Shape.Fill.Transparency = 0
                If ppSlide.SlideIndex < slidesCount Then 'Because there are no slides after last slide
                    .Cell(1, ppSlide.SlideIndex + 1).Shape.Fill.ForeColor.RGB = present
                    .Cell(1, ppSlide.SlideIndex + 1).Shape.Fill.Transparency = 0
                End If
            End With
        End If
    Next ppSlide
End Sub

'-------------------------References-------------------------'
'https://stackoverflow.com/questions/24929913/powerpoint-vba-loop-all-slides-all-shapes-find-chart-set-datalabel-color-to
'https://stackoverflow.com/questions/34497087/how-to-change-table-header-color-in-ppt-vba
'https://stackoverflow.com/questions/55378289/formatting-powerpoint-table-using-vba-very-slow
'https://stackoverflow.com/questions/8999796/how-to-get-power-point-slide-dimension-using-vba
'https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtable
'https://www.tek-tips.com/viewthread.cfm?qid=1606256
'https://stackoverflow.com/questions/57117613/getting-the-active-slide-of-a-ppt-presentation-via-vba-but-from-excel
'https://stackoverflow.com/questions/45391119/powerpoint-slide-count-variable-in-vba
'https://stackoverflow.com/questions/67504136/slide-count-variable-in-vba-powerpoint
