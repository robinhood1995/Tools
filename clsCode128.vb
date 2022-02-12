Imports System.Drawing

Namespace Barcode

    Public Class clsCode128
        Private bitsCode As ArrayList

        'http://www.vbforums.com/showthread.php?349118-EAN-13-CODE39-UPC-CODE128-Code-Bar-Generator

        Public Sub New()
            bitsCode = New ArrayList
            bitsCode.Add("11011001100")
            bitsCode.Add("11001101100")
            bitsCode.Add("11001100110")
            bitsCode.Add("10010011000")
            bitsCode.Add("10010001100")
            bitsCode.Add("10001001100")
            bitsCode.Add("10011001000")
            bitsCode.Add("10011000100")
            bitsCode.Add("10001100100")
            bitsCode.Add("11001001000")
            bitsCode.Add("11001000100")
            bitsCode.Add("11000100100")
            bitsCode.Add("10110011100")
            bitsCode.Add("10011011100")
            bitsCode.Add("10011001110")
            bitsCode.Add("10111001100")
            bitsCode.Add("10011101100")
            bitsCode.Add("10011100110")
            bitsCode.Add("11001110010")
            bitsCode.Add("11001011100")
            bitsCode.Add("11001001110")
            bitsCode.Add("11011100100")
            bitsCode.Add("11001110100")
            bitsCode.Add("11101101110")
            bitsCode.Add("11101001100")
            bitsCode.Add("11100101100")
            bitsCode.Add("11100100110")
            bitsCode.Add("11101100100")
            bitsCode.Add("11100110100")
            bitsCode.Add("11100110010")
            bitsCode.Add("11011011000")
            bitsCode.Add("11011000110")
            bitsCode.Add("11000110110")
            bitsCode.Add("10100011000")
            bitsCode.Add("10001011000")
            bitsCode.Add("10001000110")
            bitsCode.Add("10110001000")
            bitsCode.Add("10001101000")
            bitsCode.Add("10001100010")
            bitsCode.Add("11010001000")
            bitsCode.Add("11000101000")
            bitsCode.Add("11000100010")
            bitsCode.Add("10110111000")
            bitsCode.Add("10110001110")
            bitsCode.Add("10001101110")
            bitsCode.Add("10111011000")
            bitsCode.Add("10111000110")
            bitsCode.Add("10001110110")
            bitsCode.Add("11101110110")
            bitsCode.Add("11010001110")
            bitsCode.Add("11000101110")
            bitsCode.Add("11011101000")
            bitsCode.Add("11011100011")
            bitsCode.Add("11011101110")
            bitsCode.Add("11101011000")
            bitsCode.Add("11101000110")
            bitsCode.Add("11100010110")
            bitsCode.Add("11101101000")
            bitsCode.Add("11101100010")
            bitsCode.Add("11100011010")
            bitsCode.Add("11101111010")
            bitsCode.Add("11001000010")
            bitsCode.Add("11110001010")
            bitsCode.Add("10100110000")
            bitsCode.Add("10100001100")
            bitsCode.Add("10010110000")
            bitsCode.Add("10010000110")
            bitsCode.Add("10000101100")
            bitsCode.Add("10000100110")
            bitsCode.Add("10110010000")
            bitsCode.Add("10110000100")
            bitsCode.Add("10011010000")
            bitsCode.Add("10011000010")
            bitsCode.Add("10000110100")
            bitsCode.Add("10000110010")
            bitsCode.Add("11000010010")
            bitsCode.Add("11001010000")
            bitsCode.Add("11110111010")
            bitsCode.Add("11000010100")
            bitsCode.Add("10001111010")
            bitsCode.Add("10100111100")
            bitsCode.Add("10010111100")
            bitsCode.Add("10010011110")
            bitsCode.Add("10111100100")
            bitsCode.Add("10011110100")
            bitsCode.Add("10011110010")
            bitsCode.Add("11110100100")
            bitsCode.Add("11110010100")
            bitsCode.Add("11110010010")
            bitsCode.Add("11011011110")
            bitsCode.Add("11011110110")
            bitsCode.Add("11110110110")
            bitsCode.Add("10101111000")
            bitsCode.Add("10100011110")
            bitsCode.Add("10001011110")
            bitsCode.Add("10111101000")
            bitsCode.Add("10111100010")
            bitsCode.Add("11110101000")
            bitsCode.Add("11110100010")
            bitsCode.Add("10111011110")
            bitsCode.Add("10111101110")
            bitsCode.Add("11101011110")
            bitsCode.Add("11110101110")
            bitsCode.Add("11010000100")
            bitsCode.Add("11010010000")
            bitsCode.Add("11010011100")
            bitsCode.Add("1100011101011")
        End Sub
        Public Function Generate(ByVal Code As String) As Image
            Dim bmp As Image
            Dim ValidInput As String = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
            Dim CheckSum As Integer
            Dim BarCode As String
            Dim i As Integer

            If Code = "" Then Throw New Exception("The code is incorrect")
            BarCode = bitsCode(104)
            For i = 0 To Code.Length - 1
                If ValidInput.IndexOf(Code.Substring(i, 1)) = -1 Then Throw New Exception("The code is incorrect")
                CheckSum += ((i + 1) * ValidInput.IndexOf(Code.Substring(i, 1)))
                BarCode &= bitsCode(ValidInput.IndexOf(Code.Substring(i, 1)))
            Next
            CheckSum += 104 'Start B
            CheckSum = CheckSum Mod 103
            BarCode &= bitsCode(CheckSum)
            BarCode &= bitsCode(106) 'Stop symbol

            bmp = New Bitmap(BarCode.Length, 58)
            Dim g As Graphics = Graphics.FromImage(bmp)
            g.FillRectangle(New SolidBrush(Color.White), 0, 0, BarCode.Length, 58)
            Dim p As New Pen(Color.Black, 1)
            Dim BarX As Integer

            ' Create font and brush.
            Dim drawFont As New Font("Arial", 8)
            Dim drawBrush As New SolidBrush(Color.Black)
            ' Create point for upper-left corner of drawing.
            Dim y As Single = 42
            ' Set format of string.
            Dim drawFormat As New StringFormat
            drawFormat.FormatFlags = StringFormatFlags.NoWrap
            drawFormat.Alignment = StringAlignment.Center

            For i = 0 To BarCode.Length - 1
                Try
                    If BarCode.Chars(i) = "1" Then g.DrawLine(p, BarX, 0, BarX, 40)
                    BarX += 1
                Catch
                End Try
            Next
            g.DrawString(Code, drawFont, drawBrush, New RectangleF(0, y, BarCode.Length, 16), drawFormat)

            Return bmp
        End Function
    End Class

End Namespace
