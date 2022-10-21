Sub progressbar()
On Error Resume Next
' Record number of visible slides
Dim N As Integer
' Index for visible slides
Dim i As Integer
Dim width As Double

width = 0.01

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
        If i <> 1 And i <> N Then
            Set sld = .Slides(x).Shapes.AddShape(msoShapeRectangle, _
            0, .PageSetup.SlideHeight * (1 - width), _
            (i - 1) * .PageSetup.SlideWidth / (N - 2), .PageSetup.SlideHeight * width)
            sld.Fill.ForeColor.RGB = RGB(218, 227, 243)
            sld.Line.Visible = False
            sld.Name = "progress bar"
        
            .Slides(x).HeadersFooters.Footer.Visible = True
            .Slides(x).HeadersFooters.Footer.Text = CStr(i - 1) ' + "/" + CStr(N)
        End If
    Else
        .Slides(x).HeadersFooters.Footer.Visible = False
    End If
Next

End With
End Sub
