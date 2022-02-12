Imports System.Drawing

Namespace Barcode

    Public Class clsCode39

        'http://www.vbforums.com/showthread.php?349118-EAN-13-CODE39-UPC-CODE128-Code-Bar-Generator

        Public Function Generate(ByVal Code As String) As Image
            Dim bmp As Image
            Dim ValidInput As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
            Dim ValidCodes As String = "4191459566472786097041902596264733841710595784729059950476626106644590602984801043246599624767444602600464775861090446866032248034439186013047842447705803036526582823575858090365863556658042365383495434978353624150635770"
            Dim i As Integer

            If Code = "" Then Throw New Exception("The code si incorrect")
            For i = 0 To Code.Length - 1
                If ValidInput.IndexOf(Code.Substring(i, 1)) = -1 Then Throw New Exception("The code is incorrect")
            Next
            Code = "*" & Code & "*"
            ValidInput &= "*"

            bmp = New Bitmap(Code.Length * 16, 58)
            Dim g As Graphics = Graphics.FromImage(bmp)
            g.FillRectangle(New SolidBrush(Color.White), 0, 0, Code.Length * 16, 58)
            Dim p As New Pen(Color.Black, 1)
            Dim BarValue, BarX As Integer
            Dim BarSlice As Short

            ' Create font and brush.
            Dim drawFont As New Font("Arial", 8)
            Dim drawBrush As New SolidBrush(Color.Black)
            ' Create point for upper-left corner of drawing.
            Dim x As Single = (Code.Length * 16) / 3
            Dim y As Single = 42
            ' Set format of string.
            Dim drawFormat As New StringFormat
            drawFormat.FormatFlags = StringFormatFlags.NoWrap

            For i = 0 To Code.Length - 1
                Try
                    BarValue = Val(ValidCodes.Substring(ValidInput.IndexOf(Code.Substring(i, 1)) * 5, 5))
                    If BarValue = 0 Then BarValue = 36538
                    For BarSlice = 15 To 0 Step -1
                        If BarValue >= 2 ^ BarSlice Then
                            g.DrawLine(p, BarX, 0, BarX, 40)
                            BarValue = BarValue - (2 ^ BarSlice)
                        End If
                        BarX += 1
                    Next
                Catch
                End Try
            Next
            g.DrawString(Code, drawFont, drawBrush, x, y, drawFormat)

            Return bmp
        End Function
    End Class

End Namespace
