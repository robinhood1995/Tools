Imports System.Math
Imports log4net
Imports log4net.Config
Imports System.Configuration
Imports System.Reflection
Imports System.IO
Imports Quiksoft.FreeSMTP
Imports System.Windows.Forms
Imports System.Net.NetworkInformation
Imports System.Drawing
Imports System.Text




Public Class clsFunctions
    ''' <summary>
    ''' A Database connection for .NET
    ''' </summary>
    ''' <remarks>
    ''' This class does not hold open a connection but 
    ''' instead is stateless: for each request it 
    ''' connects, performs the request and disconnects.
    ''' </remarks>
    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsFunctions))


    'Private Sub New()
    '    _log.Info("Starting " & MethodBase.GetCurrentMethod().ToString())
    'End Sub

#Region " Faxing Out "
    ''' <summary>
    ''' Sending a Fax out
    ''' </summary>
    ''' <param name="faxsrvname">Fax Server HostName or IP</param>
    ''' <param name="PathandDoc">Path and Document Name to send</param>
    ''' <param name="PhoneNum">Contact Fax Number</param>
    ''' <param name="Contact">Contact Name</param>
    ''' <param name="Email">Contact Email</param>
    ''' <remarks></remarks>
    'Public Shared Function Fax(ByRef faxsrvname As String, ByRef PathandDoc As String, ByRef PhoneNum As String _
    '               , ByRef Contact As String, ByRef Email As String)
    '    Dim objFaxServer As New FAXCOMEXLib.FaxServer    'connection to the server
    '    Dim objFaxDocument As New FAXCOMEXLib.FaxDocument 'fax document to send
    '    Dim strFaxPDFtoSend As String
    '    'local document to send
    '    strFaxPDFtoSend = "" & PathandDoc & ""
    '    Try
    '        'now we have all the info, we can try and send the job out
    '        'Connect to the fax server
    '        objFaxServer.Connect(faxsrvname)
    '        'Set the fax body   
    '        objFaxDocument.Body = strFaxPDFtoSend
    '        'Name the document
    '        objFaxDocument.DocumentName = "Fax from Kiwiplan's ESP Server"
    '        'Set the fax priority
    '        objFaxDocument.Priority = FAXCOMEXLib.FAX_PRIORITY_TYPE_ENUM.fptNORMAL
    '        'Add the recipient with the fax no 
    '        objFaxDocument.Recipients.Add(PhoneNum, Contact)
    '        'Set the cover page type and the path to the cover page
    '        objFaxDocument.CoverPageType = FAXCOMEXLib.FAX_COVERPAGE_TYPE_ENUM.fcptSERVER
    '        'objFaxDocument.CoverPage = "generic.cov"
    '        objFaxDocument.CoverPage = FAXCOMEXLib.FAX_COVERPAGE_TYPE_ENUM.fcptNONE
    '        'Provide the address for the fax receipt
    '        objFaxDocument.ReceiptAddress = Email
    '        'Dont attach the original fax to the email receipt 
    '        objFaxDocument.AttachFaxToReceipt = False
    '        'Set the receipt type to email
    '        objFaxDocument.ReceiptType = FAXCOMEXLib.FAX_RECEIPT_TYPE_ENUM.frtMAIL
    '        'Subject into the cover
    '        objFaxDocument.Subject = "Fax"
    '        'Set the sender properties.
    '        objFaxDocument.Sender.Name = "ESP SERVER"
    '        'Submit the document to the connected fax server
    '        objFaxDocument.ConnectedSubmit(objFaxServer)

    '    Catch ex As Exception
    '        _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
    '        Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
    '        'MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Function
#End Region

#Region " Test Fax Server "
    Public Shared Function TestFaxing(ByRef faxsrvname As String)
        Dim objFaxServer As New FAXCOMEXLib.FaxServer
        Dim objFaxOutgoingQueue As FAXCOMEXLib.FaxOutgoingQueue
        Try
            'Connect to the fax server.
            objFaxServer.Connect(faxsrvname)
            'Create the outgoing queue 
            objFaxOutgoingQueue = objFaxServer.Folders.OutgoingQueue
            'Ensure that the queue is not blocked or paused
            objFaxOutgoingQueue.Blocked = False
            objFaxOutgoingQueue.Paused = False
            'Set the number of retries
            objFaxOutgoingQueue.Retries = 3
            'Set the retry delay no of minutes
            objFaxOutgoingQueue.RetryDelay = 1
            'save the prefs
            objFaxOutgoingQueue.Save()
            MsgBox("Fax Settings Saved")
        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Function
#End Region

#Region "From 16ths"
    ''' <summary>
    ''' Converts Sixteenths stored to inches
    ''' </summary>
    ''' <param name="Value">Integer in as 16ths</param>
    ''' <returns>A String representing "inches.16ths" formatted to 2 decimal places</returns>
    ''' <remarks></remarks>
    Public Shared Function To16ths(ByVal Value As Object) As String
        If Value = Nothing Then
            To16ths = Nothing
            'To16ths = ""
        Else
            'To16ths = Format(CInt((iValue / 16) + (iValue Mod 16) / 100), "0.00")
            To16ths = Convert.ToString(Format(Int(Value / 16) + (Value Mod 16) / 100, "0.00"))
            'to16ths = ((iValue / 16) + (iValue Mod 16) / 100)
        End If
    End Function
#End Region

#Region "To 16ths"
    ''' <summary>
    ''' Converts inches to Sixteenths
    ''' </summary>
    ''' <param name="measurement">in 16ths as Integer</param>
    ''' <returns>A String representing "inches.16ths" formatted to 2 decimal places</returns>
    ''' <remarks></remarks>
    Public Shared Function To16ths(ByVal measurement As Double) As String
        Dim wholeNumber As Integer = Math.Floor(measurement)
        Dim frac As Integer = Math.Round((measurement - wholeNumber) * 16)
        Return (wholeNumber.ToString & " " & frac.ToString & "/16")
    End Function
#End Region

#Region "Encryption String"
    ''' <summary>
    ''' Encrypt a String
    ''' </summary>
    ''' <param name="cryptStr">String to/from Encryption</param>
    ''' <returns>Encrypted or Decrypted Value</returns>
    ''' <remarks></remarks>
    Public Shared Function EncryptDecryptString(ByVal cryptStr As Object) As String

        If cryptStr = "" Then
            EncryptDecryptString = ""
            Exit Function
        End If
        Dim temp As String = Nothing
        Dim PwdChr As Integer
        Dim EncryptKey As Integer

        EncryptKey = Int(Sqrt(Len(cryptStr) * 81)) + 23

        For PwdChr = 1 To Len(cryptStr)
            temp = temp + Chr(Asc(Mid(cryptStr, PwdChr, 1)) Xor EncryptKey)
        Next PwdChr

        EncryptDecryptString = temp

    End Function
#End Region

#Region "Check if Numberic"
    ''' <summary>
    ''' Value Numeric Check
    ''' </summary>
    ''' <param name="Number">Number Value</param>
    ''' <returns>If the Value is a Number</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNumeric(ByVal Number As Object) As Boolean
        If Number = Nothing Then
            Number = Nothing
        Else
            Dim i As Integer
            For i = 0 To Number.Length - 1
                If Not Char.IsNumber(Number, i) Then
                    Return False
                End If
            Next
            Return True
        End If
    End Function
#End Region

#Region "Check if Alpha Only"
    ''' <summary>
    ''' Value String Check
    ''' </summary>
    ''' <param name="Text">Value in Text</param>
    ''' <returns>If the Value is Text</returns>
    ''' <remarks></remarks>
    Public Shared Function IsChar(ByVal Text As String) As Boolean
        Dim i As Integer
        For i = 0 To Text.Length - 1
            If Not Char.IsLetter(Text, i) Then
                Return False
            End If
        Next
        Return True
    End Function
#End Region

#Region "Convert Julian to Date"
    'Public Shared Function Julian2Date(ByVal datDate As Date) As Double
    '    Dim GGG
    '    Dim DD, MM, YY
    '    Dim S, A
    '    Dim JD, J1

    '    MM = Month(datDate)
    '    DD = Day(datDate)
    '    YY = Year(datDate)
    '    GGG = 1

    '    If (YY <= 1585) Then
    '        GGG = 0
    '    End If

    '    JD = -1 * Int(7 * (Int((MM + 9) / 12) + YY) / 4)
    '    S = 1

    '    If ((MM - 9) < 0) Then
    '        S = -1
    '    End If

    '    A = Abs(MM - 9)
    '    J1 = Int(YY + S * Int(A / 7))
    '    J1 = -1 * Int((Int(J1 / 100) + 1) * 3 / 4)
    '    JD = JD + Int(275 * MM / 9) + DD + (GGG * J1)
    '    JD = JD + 1721027 + 2 * GGG + 367 * YY

    '    If ((DD = 0) And (MM = 0) And (YY = 0)) Then
    '        MsgBox("Please enter a meaningful date!")
    '    Else
    '        Julian2Date = JD
    '    End If

    'End Function


    Public Function JulianToDate(ByVal vntJulianDate As Object) As Date
        ' Convert Julian date to VB Date format
        Dim lJulianDate, mlYear, mlday, mlDaysInYear
        Dim lYear, DaysInTheYear, lDaysInYear, mdDate, lDay

        If Not IsNothing(vntJulianDate) Then    'did they supply the Julian date in the function
            lJulianDate = vntJulianDate
        End If
        If Len(CStr(lJulianDate)) > 5 Then
            Err.Raise(vbObjectError + 2, "clsJulianToDate:JulianToDate", "Julian date greater than 5 characters.")
            Exit Function
        ElseIf Len(CStr(lJulianDate)) < 1 Then
            Err.Raise(vbObjectError + 3, "clsJulianToDate:JulianToDate", "Julian date less than one characters.")
            Exit Function
        End If

        mlYear = lJulianDate \ 1000             'get the year part
        mlday = lJulianDate - lYear * 1000      'get the day of the year part
        mlDaysInYear = DaysInTheYear(DateSerial(lYear, 1, 1)) 'number of days in the year
        If mlday >= 1 And mlday <= lDaysInYear Then         'within the range?
            mdDate = DateSerial(lYear, 1, 1) + lDay - 1     'yes, return what we found
            JulianToDate = mdDate                           'and return in the function
        Else
            Err.Raise(vbObjectError + 1, "clsJulianToDate:JulianToDate", "Invalid Julian day, less than 1 or greater than " & lDaysInYear & ".")
        End If
    End Function
#End Region

#Region "Convert Numbers to Letter"
    'Author: Alan L. Lesmerises
    'From: http://www.freevbcode.com/ShowCode.asp?ID=4303

    Function ColumnLetter(ByVal ColumnNumber As Integer) As String
        If ColumnNumber > 26 Then

            ' 1st character:  Subtract 1 to map the characters to 0-25,
            '                 but you don't have to remap back to 1-26
            '                 after the 'Int' operation since columns
            '                 1-26 have no prefix letter

            ' 2nd character:  Subtract 1 to map the characters to 0-25,
            '                 but then must remap back to 1-26 after
            '                 the 'Mod' operation by adding 1 back in
            '                 (included in the '65')

            ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) &
                           Chr(((ColumnNumber - 1) Mod 26) + 65)
        Else
            ' Columns A-Z
            ColumnLetter = Chr(ColumnNumber + 64)
        End If
    End Function
#End Region

#Region "Convert Multi-Letters to Numbers"
    Function MultiLetter(ByVal InputNumber) As String
        Dim CumSum As Object, InputValue As Object
        Dim StringPosition As Integer
        Dim i As Integer, Modulus As Integer
        Dim TempString As String, PartialValue As Object
        On Error GoTo Err_MultiLetter

        InputValue = CDec(InputNumber)

        If InputValue < 1 Then
            MultiLetter = ""
        Else
            StringPosition = 0
            CumSum = CDec(0)
            TempString = ""
            Do
                PartialValue = Int(CDec((InputValue - CumSum - 1) / (26 ^ StringPosition)))
                ' The code above should be all on 1 line ...

                Modulus = PartialValue - Int(CDec(PartialValue / 26)) * 26
                TempString = Chr(Modulus + 65) & TempString
                StringPosition = StringPosition + 1
                CumSum = CDec(0)
                For i = 1 To StringPosition
                    CumSum = CDec((CumSum + 1) * 26)
                Next i
            Loop While InputValue > CumSum
            MultiLetter = TempString
        End If

        Exit Function

Err_MultiLetter:
        MsgBox("Error " & Err.Number & ": " & Err.Description)

    End Function
#End Region

#Region "Column to Letter"
    ''' <summary>
    ''' Column number to letter
    ''' </summary>
    ''' <param name="colnum"></param>
    ''' <returns></returns>
    Public Function ColLetter(ByVal colnum As Long) As String
        '**************** Column Letter Value *****************
        '                 Recursive function
        ' Translate a column number into a letter value
        '******************************************************
        ' Converting a decimal number to a series of letters
        ' is the same as converting one number system to
        ' another (eg: decimal to hex).
        '
        ' Each column represents a power of 26 -- so if the
        '   number is > 26
        '    1. Run the integer portion of (columnNumber/26)
        '       thru the process and
        '    2. Prefix the returned letter to the letter
        '       representing the remainder
        ' The function will call itself as many times as
        '   needed to get the column number below or equal
        '   to 26
        '
        ' Since the alphabet does not have a representation
        ' of zero, and the integer arithmetic involved
        ' returns a zero at the boundaries of the range, we
        ' need to make some adjustments. Whenever a mod
        ' operation returns a zero, we reset it to 26 and,
        ' where necessary, decrement the next higher column.
        '***************************************************
        Dim wk As String
        Dim wkn As Long

        If colnum > 26 Then
            wkn = colnum \ 26
            If colnum Mod 26 = 0 Then
                wkn = wkn - 1
            End If
            wk = ColLetter(wkn)
        End If
        wkn = (colnum Mod 26)
        If wkn = 0 Then wkn = 26

        ColLetter = wk & Chr(Asc("A") - 1 + wkn)
    End Function
#End Region

#Region " Create Directory "
    ''' <summary>
    ''' Create Directory
    ''' </summary>
    ''' <param name="Path"></param>
    ''' <returns></returns>
    Public Shared Function CreateDir(ByRef Path As String)
        Try
            If Not Directory.Exists(Path) Then
                Directory.CreateDirectory(Path)
            End If

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Function
#End Region

#Region " Email "
    ''' <summary>
    ''' Send a email
    ''' </summary>
    ''' <param name="FromEmail"></param>
    ''' <param name="ToEmail"></param>
    ''' <param name="Subject"></param>
    ''' <param name="FileName"></param>
    ''' <param name="SMTPServer"></param>
    ''' <param name="SMTPPort"></param>
    ''' <returns></returns>
    Public Shared Function NetEmail(ByRef FromEmail As String, ByRef ToEmail As String,
                                    ByRef Subject As String, ByRef FileName As String,
                                          Optional ByRef SMTPServer As String = "localhost",
                                    Optional ByRef SMTPPort As Integer = 25)
        Try
            'Create Message
            Dim oMail As New System.Net.Mail.MailMessage

            'set the addresses
            oMail.From = New System.Net.Mail.MailAddress(FromEmail)
            oMail.To.Add(ToEmail)

            'set the content
            oMail.Subject = Subject
            oMail.Body = "This email is from Kiwiplan's ESP systems."

            'add an attachment from the filesystem
            oMail.Attachments.Add(New System.Net.Mail.Attachment("" & FileName & ""))

            Dim smtp As New System.Net.Mail.SmtpClient(SMTPServer, SMTPPort)
            smtp.Send(oMail)

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Function
#End Region

#Region " Check Text in Combo "
    ''' <summary>
    ''' Verify text in a combo box
    ''' </summary>
    ''' <param name="objCombo"></param>
    ''' <param name="TextToFind"></param>
    ''' <returns></returns>
    Public Function CheckIfExistInCombo(ByVal objCombo As Object, ByVal TextToFind As String) As Boolean
        Dim NumOfItems As Object 'The Number Of Items In ComboBox
        Dim IndexNum As Integer 'Index

        NumOfItems = objCombo.ListCount
        For IndexNum = 0 To NumOfItems - 1
            If objCombo.List(IndexNum) = TextToFind Then
                CheckIfExistInCombo = True
                Exit Function
            End If
        Next IndexNum

        CheckIfExistInCombo = False
    End Function
#End Region

#Region " Loop throught rows "
    ''' <summary>
    ''' Loop through rows in a datagridview
    ''' </summary>
    ''' <param name="datagrd">DataGridView Object</param>
    ''' <returns>Message Box in text of each cell value</returns>
    ''' <remarks></remarks>
    Public Shared Function RowLoop(ByVal datagrd As DataGridView)
        Dim output As String = String.Empty
        For Each row As DataGridViewRow In datagrd.Rows
            For Each cell As DataGridViewCell In row.Cells
                output += cell.Value & ":"
            Next
            output += vbCrLf
        Next
        MsgBox(output)

    End Function
#End Region

#Region " Get Rows in Data Grid "
    ''' <summary>
    ''' Get a selected row from a DataGridView
    ''' </summary>
    ''' <param name="dataGridView"></param>
    ''' <returns></returns>
    Public Shared Function GetSelectedRowsDGV(ByVal dataGridView As DataGridView) As List(Of DataRow)
        Dim list As New List(Of DataRow)()
        'If the dataGridView selected rows are null
        If dataGridView.SelectedRows Is Nothing OrElse dataGridView.SelectedRows.Count = 0 Then
            'Return list
            Return list
        End If
        For Each dataGridViewRow As DataGridViewRow In dataGridView.SelectedRows
            'Declare a null DataRow
            Dim dataRow As DataRow = Nothing
            Try
                'Declare a DataRowView(DataRowView)
                Dim dataRowView As DataRowView = DirectCast(dataGridViewRow.DataBoundItem, DataRowView)
                dataRow = dataRowView.Row

            Catch ex As Exception
                'Catch the exception
                'Write the error message
                _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
                MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            'If the row is null
            If dataRow Is Nothing Then
                'Continue
                Continue For
            End If
            'Add a row(DataRow) to list
            list.Add(dataRow)
        Next
        'Return list
        Return list
    End Function
#End Region

#Region " Get a Row in DGV to Datatable "
    ''' <summary>
    ''' Get a specific row in a DataGridview by index
    ''' </summary>
    ''' <param name="intSelectedIndex"></param>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    Public Shared Function GetSelected_RowInDataTable(ByVal intSelectedIndex As Integer, ByVal dgv As DataGridView) As DataTable
        Dim dt As New DataTable

        _log.Info("Selected " & dgv.SelectedRows.Count & " rows")
        'Add the columns
        Dim col As DataColumn

        'For each columns in the datagridveiw add a new column to data table

        For Each dgvCol As DataGridViewColumn In dgv.Columns
            col = New DataColumn(dgvCol.Name)
            dt.Columns.Add(col)
        Next

        'Add the selected rows from the datagridview

        Dim row As DataRow

        row = dt.Rows.Add

        For Each column As DataGridViewColumn In dgv.Columns
            row.Item(column.Index) = dgv.Rows.Item(intSelectedIndex).Cells(column.Index).Value
        Next

        Return dt
    End Function

#End Region

#Region " Get Selected Rows in DGV to Datatable "
    ''' <summary>
    ''' Loop through a DataGridView and return the Selected Rows
    ''' </summary>
    ''' <param name="dgv">DataGridView Object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSelectedRowsInDataGridView(ByVal dgv As DataGridView) As DataTable
        Try
            Dim dt As New DataTable

            _log.Info("Selected " & dgv.SelectedRows.Count & " rows")
            For Each col As DataGridViewColumn In dgv.Columns
                dt.Columns.Add(col.DataPropertyName, col.ValueType)
            Next
            For Each gridRow As DataGridViewRow In dgv.SelectedRows
                If gridRow.IsNewRow Then
                    Continue For
                End If

                Dim dtRow As DataRow = dt.NewRow()
                For i1 As Integer = 0 To dgv.Columns.Count - 1
                    dtRow(i1) = (If(gridRow.Cells(i1).Value Is Nothing, DBNull.Value, gridRow.Cells(i1).Value))
                Next
                dt.Rows.Add(dtRow)
            Next

            Return dt
        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        Finally

        End Try

    End Function

#End Region

#Region " Display a dataset table and values "
    Public Shared Function DisplayTable(ByVal ds As DataSet)
        Try
            Dim output As String = Nothing

            For Each r As DataRow In ds.Tables(0).Rows
                For Each c As DataColumn In ds.Tables(0).Columns
                    output += c.ColumnName & vbTab & ":" & r(c)
                    output += vbCrLf
                Next
            Next

            MsgBox(output)

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        Finally

        End Try
    End Function
#End Region

#Region " CountOccurences "
    Shared Function CountOccurences(ByVal searchIn As String, ByVal searchFor As String) As Integer
        Dim ipos As Integer = 1
        Dim IntCount As Integer = 0

        If searchIn.Length > 0 Then
            Do While InStr(ipos, searchIn, searchFor) > 0
                IntCount += 1
                If InStr(ipos, searchIn, searchFor) < searchIn.Length Then
                    ipos = InStr(ipos, searchIn, searchFor) + 1
                Else
                    Exit Do
                End If
            Loop
        End If
        Return IntCount
    End Function
#End Region

#Region " Get/set text file contents "
    ''' <summary>
    ''' Get/Read file content
    ''' </summary>
    ''' <param name="FullPath"></param>
    ''' <param name="ErrInfo"></param>
    ''' <returns></returns>
    Public Shared Function GetFileContents(ByVal FullPath As String,
       Optional ByRef ErrInfo As String = "") As String

        Dim strContents As String
        Dim objReader As StreamReader
        Try

            objReader = New StreamReader(FullPath)
            strContents = objReader.ReadToEnd()
            objReader.Close()
            Return strContents
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Save data to a file
    ''' </summary>
    ''' <param name="strData"></param>
    ''' <param name="FullPath"></param>
    ''' <param name="ErrInfo"></param>
    ''' <returns></returns>
    Public Shared Function SaveTextToFile(ByVal strData As String,
     ByVal FullPath As String,
       Optional ByVal ErrInfo As String = "") As Boolean

        Dim Contents As String
        Dim bAns As Boolean = False
        Dim objReader As StreamWriter
        Try


            objReader = New StreamWriter(FullPath)
            objReader.Write(strData)
            objReader.Close()
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message

        End Try
        Return bAns
    End Function
#End Region

#Region " Ping Server"
    ''' <summary>
    ''' Ping Server
    ''' </summary>
    ''' <param name="nameOrAddress"></param>
    ''' <returns></returns>
    Public Shared Function PingHost(ByVal nameOrAddress As String) As Boolean
        Dim pingable As Boolean = False
        Dim pinger As Ping = New Ping
        Try
            Dim reply As PingReply = pinger.Send(nameOrAddress)
            pingable = (reply.Status = IPStatus.Success)
        Catch ex As PingException
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            'Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        End Try

        Return pingable
    End Function

#End Region

#Region " Image Resixe and Center "
    ''' <summary>
    ''' AutoSize and center and image in a picture box
    ''' </summary>
    ''' <param name="ImagePath"></param>
    ''' <param name="picBox"></param>
    ''' <param name="pSizeMode"></param>
    Public Sub AutosizeImage(ByVal ImagePath As String, ByVal picBox As PictureBox, Optional ByVal pSizeMode As PictureBoxSizeMode = PictureBoxSizeMode.CenterImage)
        Try
            picBox.Image = Nothing
            picBox.SizeMode = pSizeMode
            If System.IO.File.Exists(ImagePath) Then
                Dim imgOrg As Bitmap
                Dim imgShow As Bitmap
                Dim g As Graphics
                Dim divideBy, divideByH, divideByW As Double
                imgOrg = DirectCast(Bitmap.FromFile(ImagePath), Bitmap)

                divideByW = imgOrg.Width / picBox.Width
                divideByH = imgOrg.Height / picBox.Height
                If divideByW > 1 Or divideByH > 1 Then
                    If divideByW > divideByH Then
                        divideBy = divideByW
                    Else
                        divideBy = divideByH
                    End If

                    imgShow = New Bitmap(CInt(CDbl(imgOrg.Width) / divideBy), CInt(CDbl(imgOrg.Height) / divideBy))
                    imgShow.SetResolution(imgOrg.HorizontalResolution, imgOrg.VerticalResolution)
                    g = Graphics.FromImage(imgShow)
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(imgOrg, New Rectangle(0, 0, CInt(CDbl(imgOrg.Width) / divideBy), CInt(CDbl(imgOrg.Height) / divideBy)), 0, 0, imgOrg.Width, imgOrg.Height, GraphicsUnit.Pixel)
                    g.Dispose()
                Else
                    imgShow = New Bitmap(imgOrg.Width, imgOrg.Height)
                    imgShow.SetResolution(imgOrg.HorizontalResolution, imgOrg.VerticalResolution)
                    g = Graphics.FromImage(imgShow)
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(imgOrg, New Rectangle(0, 0, imgOrg.Width, imgOrg.Height), 0, 0, imgOrg.Width, imgOrg.Height, GraphicsUnit.Pixel)
                    g.Dispose()
                End If
                imgOrg.Dispose()

                picBox.Image = imgShow
            Else
                picBox.Image = Nothing
            End If


        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region

#Region " Scale Image "
    ''' <summary>
    ''' Scale and Imaage to a certain size
    ''' </summary>
    ''' <param name="OldImage"></param>
    ''' <param name="TargetHeight"></param>
    ''' <param name="TargetWidth"></param>
    ''' <returns></returns>
    Public Shared Function ScaleImage(ByVal OldImage As Image, ByVal TargetHeight As Integer,
                           ByVal TargetWidth As Integer) As Image

        Dim NewHeight As Integer = TargetHeight
        Dim NewWidth As Integer = CInt(NewHeight / OldImage.Height * OldImage.Width)

        If NewWidth > TargetWidth Then
            NewWidth = TargetWidth
            NewHeight = CInt(NewWidth / OldImage.Width * OldImage.Height)
        End If

        Dim NewImage As Image = New Bitmap(OldImage, NewWidth, NewHeight)

        Return NewImage

    End Function
#End Region

#Region " CSV to Datatable"
    ''' <summary>
    ''' Covert and CSV file with Headers to a datatable
    ''' </summary>
    ''' <param name="strfilename"></param>
    ''' <param name="delimeter"></param>
    ''' <returns></returns>
    Public Shared Function csvToDatatable(ByVal strfilename As String, ByVal delimeter As Char)
        Try
            Dim dt As New System.Data.DataTable
            Dim firstLine As Boolean = True
            If IO.File.Exists(strfilename) Then
                Dim SR As StreamReader = New StreamReader(strfilename)
                Dim line As String = SR.ReadLine()
                Dim strArray As String() = line.Split(delimeter)
                Dim row As DataRow

                For Each s As String In strArray
                    dt.Columns.Add(New DataColumn(s))
                Next

                Do
                    line = SR.ReadLine
                    If Not line = String.Empty Then
                        row = dt.NewRow()
                        row.ItemArray = line.Split(delimeter)
                        dt.Rows.Add(row)
                    Else
                        Exit Do
                    End If
                Loop
            End If

            Return dt

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            GC.Collect()
        End Try
    End Function
#End Region

#Region " Export DataGridView to CSV "
    ''' <summary>
    ''' Export DatagridView Data to a CSV
    ''' </summary>
    ''' <param name="datagridview"></param>
    ''' <param name="strfilename"></param>
    ''' <param name="delimeter"></param>
    ''' <returns></returns>
    Public Shared Function DataGridViewTocsv(ByRef datagridview As DataGridView, ByRef strfilename As String, ByRef delimeter As Char)
        Try
            Dim sb As New StringBuilder
            For i As Integer = 0 To datagridview.Rows.Count - 1
                Dim array As String() = New String(datagridview.Columns.Count - 1) {}
                If i.Equals(0) Then
                    For j As Integer = 0 To datagridview.Columns.Count - 1
                        array(j) = datagridview.Columns(j).HeaderText.Replace(delimeter, "~")
                    Next
                    sb.AppendLine(String.Join(delimeter, array))
                End If
                For j As Integer = 0 To datagridview.Columns.Count - 1
                    If Not datagridview.Rows(i).IsNewRow Then
                        array(j) = datagridview(j, i).Value.Replace(delimeter, "~").ToString
                    End If
                Next
                If Not datagridview.Rows(i).IsNewRow Then
                    sb.AppendLine(String.Join(delimeter, array))
                End If
            Next
            File.WriteAllText(strfilename, sb.ToString)

            Return strfilename

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            GC.Collect()
        End Try
    End Function
#End Region

#Region " Check Email Addresses "
    Function IsValidEmailFormat(ByVal s As String) As Boolean
        Try
            Dim a As New Net.Mail.MailAddress(s)
        Catch
            Return False
        End Try
        Return True
    End Function
#End Region

#Region " Random Codes "
    Public Shared Function GetRandomString(ByVal iLength As Integer) As String
        Dim sResult As String = ""
        Dim rdm As New Random()

        For i As Integer = 1 To iLength
            sResult &= ChrW(rdm.Next(32, 126))
        Next

        Return sResult
    End Function

    Public Shared Function GenerateRandomString(ByRef iLength As Integer) As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray()
        Dim sResult As String = ""

        For i As Integer = 0 To iLength - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next

        Return sResult
    End Function

#End Region

    'Function GetUserName() As String

    '    Dim selectQuery As Management.SelectQuery = New Management.SelectQuery("Win32_Process")
    '    Dim searcher As Management.ManagementObjectSearcher = New Management.ManagementObjectSearcher(selectQuery)
    '    Dim y As System.Management.ManagementObjectCollection
    '    y = searcher.Get

    '    For Each proc As Management.ManagementObject In y
    '        Dim s(1) As String
    '        proc.InvokeMethod("GetOwner", CType(s, Object()))
    '        Dim n As String = proc("Name").ToString()
    '        If n = "explorer.exe" Then
    '            Return s(0)
    '        End If
    '    Next
    'End Function


End Class
