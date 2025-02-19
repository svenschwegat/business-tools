Sub SaveCurrentSlideAsPNG()
    Dim slideNumber As Integer
    Dim slide As slide
    Dim folder As String
    Dim filePath As String
    Dim fileName As String
    Dim formattedDate As String
    Dim extensionPosition As Integer
    Dim response As VbMsgBoxResult
      
    ' Get the current slide number, file name and date
    slideNumber = ActiveWindow.View.slide.SlideIndex
    fileName = ActivePresentation.Name
    extensionPosition = InStrRev(fileName, ".")
    If extensionPosition > 0 Then
        fileName = Left(presentationName, extensionPosition - 1)
    End If
    
    formattedDate = Format(Date, "yyyy-mm-dd")
      
    ' Set the file path
    folder = Environ("USERPROFILE") & "\Downloads\"
    If Right(folder, 1) <> "\" Then
        folder = folder & "\"
    End If
    filePath = folder & formattedDate & " " & fileName & " " & slideNumber & ".png"

    ' Export the current slide as a PNG
    Set slide = ActivePresentation.Slides(slideNumber)
    slide.Export filePath, "PNG"
    
    ' Show Success Message and open folder if user clicks yes
    response = MsgBox("Slide saved as PNG in this folder: " & filePath & vbCrLf & "Do you want to open the folder?", vbYesNo + vbQuestion, "Slide Exported")
    If response = vbYes Then
        Shell "explorer.exe /select," & filePath, vbNormalFocus
    End If
End Sub