Attribute VB_Name = "PowerBar"
Option Explicit

'Method to create a Powerbar'
Sub CreatePowerBar()

    Dim amountOfSlides As Long
    Dim SpezialSlides As Long
    Dim typ As Long
    Dim left As Long
    Dim top As Long
    Dim bottom As Long
    Dim slideHeight As Long
    Dim slideWidth As Long
    Dim progress As Long
    Dim progressbarHeight As Long
    Dim progressbarLength As Long
    Dim color As String
    Dim displayPercentage As Boolean
    
    amountOfSlides = ActivePresentation.Slides.count
    slideHeight = ActivePresentation.PageSetup.slideHeight
    slideWidth = ActivePresentation.PageSetup.slideWidth
    
    
    'Color of the PowerBar'
    color = RGB(65, 174, 189)
    
    'Amount of first Slides without the Powerbar'
    SpezialSlides = 2
    
    'Style the PowerBar'
    typ = 1                                         'Type of the shape of he Powerbar'
    displayPercentage = True                        'Display the percentage of the progress'
    progressbarHeight = 19                          'Height of the PowerBar'
    left = 20                                       'Spacing to the left border of the slide'
    bottom = 10                                     'Spacing to the bottom border of the slide'
    
    'Calculated Style Attrubutes:'
    top = slideHeight - bottom - progressbarHeight  '"Top" is the calculated equivalent to "bottom". If you change "top" make sure you keep "bottom" in mind.'
    progressbarLength = slideWidth - left * 2       'The length of the PowerBar is calculated using the space to the left border'

    Dim slide As slide
    For Each slide In ActivePresentation.Slides
        
        If slide.SlideNumber > SpezialSlides Then
            
            'Add a shape to display the progress'
            progress = progressbarLength / amountOfSlides * slide.SlideNumber
            Dim progressShape As shape
            Set progressShape = slide.shapes.AddShape(typ, left, top, progress, progressbarHeight)
            progressShape.Name = "PowerBarProgress" & slide.SlideNumber
            progressShape.Fill.ForeColor.RGB = color
            progressShape.Fill.BackColor.RGB = color

            'Add a surrounding box'
            Dim boxShape As shape
            Set boxShape = slide.shapes.AddShape(1, left, top, progressbarLength, progressbarHeight) 'typ must be 1 here to create a box'
            boxShape.Name = "PowerBarBox" & slide.SlideNumber
            boxShape.Fill.ForeColor.RGB = color
            boxShape.Fill.BackColor.RGB = color
            
            'Display the progress as percentage'
            If displayPercentage Then
                Dim percentage As Long: percentage = 100 / amountOfSlides * slide.SlideNumber
                boxShape.Fill.Transparency = 0.7
                boxShape.TextFrame.TextRange.Text = percentage & "%"
                boxShape.TextFrame.TextRange.Font.Size = 14
                boxShape.TextFrame.TextRange.Font.Name = "Calibri"
            End If
            
        End If
    Next
End Sub


Sub RemovePowerBar()
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        Dim powerbarDeleted As Boolean: powerbarDeleted = True
        Do While powerbarDeleted
            powerbarDeleted = False
            For Each shape In slide.shapes
                If shape.Name Like "PowerBar*" Then
                    shape.Delete    'Seems like: When the shape is deleted the For Each Loop does not point to the shape after the deleted one but to the one after.'
                    powerbarDeleted = True
                End If
             Next
        Loop
    Next
End Sub


Sub RefreshPowerBar()
   RemovePowerBar
   CreatePowerBar
End Sub

