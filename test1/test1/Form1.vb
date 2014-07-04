Imports PowerPoint = Microsoft.Office.Interop.PowerPoint



Public Class Form1
    Dim imageFile As String = "images\bank.jpg"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ppApplication As PowerPoint.ApplicationClass = Nothing
        Dim ppPresentation As PowerPoint.Presentation = Nothing
        Dim ppSlide As PowerPoint.Slide = Nothing
        Dim ppTextRange As PowerPoint.TextRange = Nothing
        ' Dim imagepath As String
       
        Dim width As Double
        Dim height As Double
        Dim image As System.Drawing.Image = System.Drawing.Image.FromFile(imageFile)
        Try
            width = image.Width * 72.0 / image.HorizontalResolution
            height = image.Height * 72.0 / image.VerticalResolution
        Finally
            image.Dispose()
        End Try

        Try
            Dim fileTest As String = "D:\testAmr\test.pptx"
            ppApplication = New Microsoft.Office.Interop.PowerPoint.ApplicationClass()
            ppPresentation = ppApplication.Presentations.Add
            ppSlide = ppPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank)

           
            'imagepath = "images\bank1"
            ppSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 20, 20, 512, 490).TextEffect.Alignment = Microsoft.Office.Core.MsoTextEffectAlignment.msoTextEffectAlignmentRight
            ppSlide.Shapes.AddPicture(imageFile, Left, Top, width, height)


            ppTextRange = ppSlide.Shapes(1).TextFrame.TextRange
            ppTextRange.Text = "this is my ppt file , hello world"
            ppTextRange.Font.Size = 15
            ppTextRange.Font.Color.RGB = RGB(200, 31, 159)
            ppTextRange.Font.Name = "Arial"
            ppPresentation.SaveAs(fileTest)

           
        Catch ex As Exception


        Finally
            ppSlide = Nothing
            If Not ppPresentation Is Nothing Then
                ppPresentation.Close()
                ppPresentation = Nothing

            End If
            If Not ppApplication Is Nothing Then
                ppApplication.Quit()
                ppApplication = Nothing

            End If


        End Try





    End Sub


    

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ppApp As PowerPoint.Application
        Dim ppPrsn As PowerPoint.Presentation
        Dim ppSlide As PowerPoint.Slide
        Dim objShape As PowerPoint.Shape

        Dim strTemplateFile As String, strImg As String

        ppApp = New PowerPoint.Application
        ppApp.Visible = True

        strTemplateFile = "D:\testAmr\test.ppt"
        strImg = "images\bank.jpg"

        ppPrsn = ppApp.Presentations.Open(strTemplateFile)

        objShape = ppPrsn.Slides(1).Shapes.AddPicture(strImg, False, True, 0, 0, 300, 500)

        objShape.ScaleHeight(0.5, Microsoft.Office.Core.MsoTriState.msoCTrue)
        objShape.ScaleWidth(0.5, Microsoft.Office.Core.MsoTriState.msoCTrue)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        AddPicture(imageFile)




    End Sub







    Public Sub AddPicture(ByVal pictureFileName As String)
        Try
            Dim presentation As PowerPoint.Presentation = Application.ActivePresentation

            If presentation IsNot Nothing Then
                Dim slide As PowerPoint.Slide = _
                 presentation.Slides.Add( _
                 presentation.Slides.Count + 1, _
                 PowerPoint.PpSlideLayout.ppLayoutPictureWithCaption)

                ' Shapes(2) is the image shape on this layout.
                Dim shape As PowerPoint.Shape = slide.Shapes(2)

                slide.Shapes.AddPicture(pictureFileName, _
                 Microsoft.Office.Core.MsoTriState.msoFalse, _
                 Microsoft.Office.Core.MsoTriState.msoTrue, _
                 shape.Left, shape.Top, shape.Width, shape.Height)

                ' Insert the file name.
                slide.Shapes(1).TextFrame.TextRange.Text = pictureFileName
            End If

        Catch ex As Exception
            MessageBox.Show("Unable to insert selected picture: " & _
             ex.Message)
        End Try
    End Sub



End Class
