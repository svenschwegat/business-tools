Sub SaveCurrentSlideAsPNG()
    Dim slideNumber As Integer
    Dim slide As slide
    Dim folder As String
    Dim filePath As String
    Dim fileName As String
    Dim year As String
    Dim month As String
      
    ' Get the current slide number, file name, year and month
    slideNumber = ActiveWindow.View.slide.SlideIndex
    presentationName = ActivePresentation.Name
    fileName = Left(presentationName, InStrRev(presentationName, ".") - 1)
    year = Format(Date, "yyyy")
    month = Format(Date, "mm")    
      
    ' Set the file path
    folder = Environ("USERPROFILE") & "\Downloads\"
    filePath = folder & year & "-" & month & " " & fileName & " " & slideNumber & ".png"

    ' Export the current slide as a PNG
    Set slide = ActivePresentation.Slides(slideNumber)
    slide.Export filePath, "PNG"
      
    MsgBox "Slide saved as PNG in this folder: " & filePath
    
End Sub