Imports Microsoft.VisualBasic
Imports System.Drawing.Imaging
Imports System.Drawing

Namespace framework
    Public Class Image

        Public Sub CreateThumbnail(ByVal largura As Integer, ByVal altura As Integer, ByVal srcpath As String, ByVal destpath As String)
            Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(srcpath)
            Dim imgthumb As System.Drawing.Image = img.GetThumbnailImage(largura, altura, Nothing, New System.IntPtr(0))
            imgthumb.Save(destpath, ImageFormat.Jpeg)

            img.Dispose()
            imgthumb.Dispose()
        End Sub


        Public Shared Sub ResizeImage(originalFile As String, newFile As String, newWidth As Integer, maxHeight As Integer, onlyResizeIfWider As Boolean)
            Dim fullsizeImage As System.Drawing.Image = System.Drawing.Image.FromFile(originalFile)
            fullsizeImage.RotateFlip(RotateFlipType.Rotate180FlipNone)
            fullsizeImage.RotateFlip(RotateFlipType.Rotate180FlipNone)
            If onlyResizeIfWider Then
                If fullsizeImage.Width <= newWidth Then
                    newWidth = fullsizeImage.Width
                End If
            End If
            Dim newHeight As Integer = fullsizeImage.Height * newWidth / fullsizeImage.Width
            If newHeight > maxHeight Then
                newWidth = fullsizeImage.Width * maxHeight / fullsizeImage.Height
                newHeight = maxHeight
            End If
            Dim newImage As System.Drawing.Image = fullsizeImage.GetThumbnailImage(newWidth, newHeight, Nothing, IntPtr.Zero)
            fullsizeImage.Dispose()
            newImage.Save(newFile)
        End Sub

    End Class
End Namespace