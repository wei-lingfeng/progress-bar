Sub ProgressBar()
On Error Resume Next
' Record number of visible slides
Dim N As Integer
' Index for visible slides
Dim i As Integer


With ActivePresentation
' Count number of visible slides
For x = 1 To .Slides.Count
    If .Slides(x).SlideShowTransition.Hidden = False Then
        N = N + 1
    End If
Next

i = 0
' Draw progress bar and label page numbers
For x = 1 To .Slides.Count
    .Slides(x).Shapes("progress bar").Delete
    If .Slides(x).SlideShowTransition.Hidden = False Then
        i = i + 1
        Set sld = .Slides(x).Shapes.AddShape(msoShapeRectangle, _
        0, .PageSetup.SlideHeight * (1 - 0.015), _
        i * .PageSetup.SlideWidth / N, .PageSetup.SlideHeight * 0.015)
        sld.Fill.ForeColor.RGB = RGB(218, 227, 243)
        sld.Line.Visible = False
        sld.Name = "progress bar"
        
        .Slides(x).HeadersFooters.Footer.Visible = True
        .Slides(x).HeadersFooters.Footer.Text = CStr(i) ' + "/" + CStr(N)
        
    Else
        .Slides(x).HeadersFooters.Footer.Visible = False
    End If
Next

End With
End Sub
