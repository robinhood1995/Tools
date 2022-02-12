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
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Net
Imports System.Globalization

Public Class clsFunctions
    ''' <summary>
    ''' A set of functions for .NET
    ''' </summary>
    ''' <remarks>
    ''' This class does not hold open a connection but 
    ''' instead is stateless: for each request it 
    ''' connects, performs the request and disconnects.
    ''' </remarks>
    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsFunctions))

    Private Shared ReadOnly onValidating As MethodInfo = GetType(Control).GetMethod("OnValidating", BindingFlags.Instance Or BindingFlags.NonPublic)
    Private Shared ReadOnly onValidated As MethodInfo = GetType(Control).GetMethod("OnValidated", BindingFlags.Instance Or BindingFlags.NonPublic)


    Private Sub New()
        _log.Info("Starting Tools DLL" & MethodBase.GetCurrentMethod().ToString())
    End Sub

#Region " Validate Controls "
    Public Function Validate(ByVal control As Control) As Boolean
        Dim e As System.ComponentModel.CancelEventArgs = New System.ComponentModel.CancelEventArgs()
        onValidating.Invoke(control, New Object() {e})
        If e.Cancel Then Return False
        onValidated.Invoke(control, New Object() {EventArgs.Empty})
        Return True
    End Function
#End Region

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
    Public Shared Function Fax(ByRef faxsrvname As String, ByRef PathandDoc As String, ByRef PhoneNum As String _
                   , ByRef Contact As String, ByRef Email As String)
        Dim objFaxServer As New FAXCOMEXLib.FaxServer    'connection to the server
        Dim objFaxDocument As New FAXCOMEXLib.FaxDocument 'fax document to send
        Dim strFaxPDFtoSend As String
        'local document to send
        strFaxPDFtoSend = "" & PathandDoc & ""
        Try
            'now we have all the info, we can try and send the job out
            'Connect to the fax server
            objFaxServer.Connect(faxsrvname)
            'Set the fax body   
            objFaxDocument.Body = strFaxPDFtoSend
            'Name the document
            objFaxDocument.DocumentName = "Fax from Kiwiplan's ESP Server"
            'Set the fax priority
            objFaxDocument.Priority = FAXCOMEXLib.FAX_PRIORITY_TYPE_ENUM.fptNORMAL
            'Add the recipient with the fax no 
            objFaxDocument.Recipients.Add(PhoneNum, Contact)
            'Set the cover page type and the path to the cover page
            objFaxDocument.CoverPageType = FAXCOMEXLib.FAX_COVERPAGE_TYPE_ENUM.fcptSERVER
            'objFaxDocument.CoverPage = "generic.cov"
            objFaxDocument.CoverPage = FAXCOMEXLib.FAX_COVERPAGE_TYPE_ENUM.fcptNONE
            'Provide the address for the fax receipt
            objFaxDocument.ReceiptAddress = Email
            'Dont attach the original fax to the email receipt 
            objFaxDocument.AttachFaxToReceipt = False
            'Set the receipt type to email
            objFaxDocument.ReceiptType = FAXCOMEXLib.FAX_RECEIPT_TYPE_ENUM.frtMAIL
            'Subject into the cover
            objFaxDocument.Subject = "Fax"
            'Set the sender properties.
            objFaxDocument.Sender.Name = "ESP SERVER"
            'Submit the document to the connected fax server
            objFaxDocument.ConnectedSubmit(objFaxServer)

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            Throw New Exception(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            'MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
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

#Region " From 16ths"
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

#Region " To 16ths"
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

#Region " Encryption String "
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

#Region " Check if Numberic "
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

#Region " Check if Alpha Only "
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

#Region " Convert Julian to Date "
    Public Shared Function Julian2Date(ByVal datDate As Date) As Double
        Dim GGG
        Dim DD, MM, YY
        Dim S, A
        Dim JD, J1

        MM = DateAndTime.Month(datDate)
        DD = DateAndTime.Day(datDate)
        YY = DateAndTime.Year(datDate)
        GGG = 1

        If (YY <= 1585) Then
            GGG = 0
        End If

        JD = -1 * Int(7 * (Int((MM + 9) / 12) + YY) / 4)
        S = 1

        If ((MM - 9) < 0) Then
            S = -1
        End If

        A = Abs(MM - 9)
        J1 = Int(YY + S * Int(A / 7))
        J1 = -1 * Int((Int(J1 / 100) + 1) * 3 / 4)
        JD = JD + Int(275 * MM / 9) + DD + (GGG * J1)
        JD = JD + 1721027 + 2 * GGG + 367 * YY

        If ((DD = 0) And (MM = 0) And (YY = 0)) Then
            MsgBox("Please enter a meaningful date!")
        Else
            Julian2Date = JD
        End If

    End Function


    Public Shared Function JulianToDate(ByVal vntJulianDate As Object) As Date
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

#Region " Convert Numbers to Letter "
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

#Region " Convert Multi-Letters to Numbers "
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

#Region " Column to Letter "
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
    Public Shared Function NetEmail(ByRef ToEmail As String,
                                    ByRef Subject As String, ByRef FileName As String,
                                    Optional ByRef FromEmail As String = "no_repy@myfflbook.com",
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
            oMail.Body = "This email is from our MyFFLbook system."

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

#Region " CSV to Datatable "
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
        'https://github.com/ukushu/DataExporter
        'https://stackoverflow.com/questions/4959722/how-can-i-turn-a-datatable-to-a-csv/52684280#52684280
        Try
            Dim sb As New StringBuilder
            For i As Integer = 0 To datagridview.Rows.Count - 1
                Dim array As String() = New String(datagridview.Columns.Count - 1) {}
                If i.Equals(0) Then
                    For j As Integer = 0 To datagridview.Columns.Count - 1
                        array(j) = datagridview.Columns(j).HeaderText.Replace(delimeter, "~") '.Replace("ID", "Scan")
                        _log.Info("Exporting header " & j.ToString)
                    Next
                    sb.AppendLine(String.Join(delimeter, array))
                End If
                For j As Integer = 0 To datagridview.Columns.Count - 1
                    If Not datagridview.Rows(i).IsNewRow Then
                        array(j) = datagridview(j, i).Value.ToString.Replace(delimeter, "~").Replace(vbCr, "").Replace(vbLf, "").ToString
                        _log.Info("Exporting data " & j.ToString)
                    End If
                Next
                If Not datagridview.Rows(i).IsNewRow Then
                    sb.AppendLine(String.Join(delimeter, array))
                    _log.Info("Adding a new line to export")

                End If
            Next
            File.WriteAllText(strfilename, sb.ToString)
            _log.Info("Adding line " & sb.ToString)

            Return strfilename

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Function
#End Region

#Region " Export to CSV "
    Public Shared Function exportToCSV(ByRef datagridview As DataGridView, ByRef strfilename As String, ByRef delimeter As Char)
        Try
            'Build the CSV file data as a Comma separated string.
            Dim csv As String = String.Empty

            'Add the Header row for CSV file.
            For Each column As DataGridViewColumn In datagridview.Columns
                csv += column.HeaderText & ","c
            Next

            'Add new line.
            csv += vbCr & vbLf

            'Adding the Rows
            For Each row As DataGridViewRow In datagridview.Rows
                For Each cell As DataGridViewCell In row.Cells
                    'Add the Data rows.
                    csv += cell.Value.ToString().Replace(",", ";") & ","c
                Next

                'Add new line.
                csv += vbCr & vbLf
            Next

            'Exporting to Excel
            File.WriteAllText(strfilename, csv)

            Return strfilename

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

#End Region

#Region " Check Email Addresses "
    Public Shared Function IsValidEmailFormat(ByVal s As String) As Boolean
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
        Try
            Dim rdm As New Random()
            Dim allowChrs() As Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray()
            Dim sResult As String = ""

            For i As Integer = 0 To iLength - 1
                sResult += allowChrs(rdm.Next(0, allowChrs.Length))
                _log.Info("Randomising " & sResult)
            Next

            Return sResult
        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
        End Try

    End Function

#End Region

#Region " Convert String to Zero "
    Public Shared Function ConvertToInteger(ByRef value As String) As Integer
        If String.IsNullOrEmpty(value) Then
            value = "0"
        End If
        Return Convert.ToInt32(value)
    End Function
#End Region

#Region " Convert String to Decimal "
    Public Shared Function ConvertToDecimal(ByRef value As String) As Decimal
        If String.IsNullOrEmpty(value) Then
            value = "0"
        End If
        Return Convert.ToDecimal(value)
    End Function
#End Region

#Region " Convert Letters to Numbers "
    ''' <summary>
    ''' Convert Alphabet to Numbers
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function getNumericValueInAlphabet(ByVal value As String) As String
        Dim Alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        ValidateInput(value)

        Dim RV As New System.Text.StringBuilder

        value = value.ToUpper

        For i As Integer = 0 To value.Length - 1
            Dim Index As Integer = Alphabet.IndexOf(value.Chars(i))
            If Index > -1 Then
                RV.Append((Index + 1).ToString.PadLeft(2, "0"c))
            Else
                Throw New ArgumentOutOfRangeException("The provided String is not inside then range")
            End If
        Next

        Return RV.ToString
    End Function

    Private Sub ValidateInput(ByVal value As String)

        If value Is Nothing Then
            Throw New ArgumentNullException
        End If

        If String.IsNullOrEmpty(value.Trim) Then
            Throw New ArgumentException("The provided String argument is empty")
        End If

    End Sub
#End Region

#Region " Unicode to String "
    ''' <summary>
    ''' This is to convert data columns from image to string
    ''' </summary>
    ''' <param name="bytes"></param>
    ''' <returns></returns>
    ''' https://www.experts-exchange.com/questions/28693849/VB-net-DataGridView-Error.html
    Public Function UnicodeBytesToString(ByVal bytes() As Byte) As String

        Return System.Text.Encoding.Unicode.GetString(bytes)

    End Function
#End Region

#Region " Get Age "
    'https://stackoverflow.com/questions/16874911/compute-age-from-given-birthdate
    Public Function GetCurrentAge(ByVal dob As Date) As Integer
        Dim age As Integer
        age = Today.Year - dob.Year
        If (dob > Today.AddYears(-age)) Then age -= 1
        Return age
    End Function

    'https://stackoverflow.com/questions/16874911/compute-age-from-given-birthdate
    Public Shared Function Age(DOfB As Object) As String
        If (Month(Date.Today) * 100) + Date.Today.Day >= (Month(DOfB) * 100) + DOfB.Day Then
            Return DateDiff(DateInterval.Year, DOfB, Date.Today)
        Else
            Return DateDiff(DateInterval.Year, DOfB, Date.Today) - 1
        End If
    End Function
#End Region

#Region " Get only Number in String "
    ''' <summary>
    ''' Get the numbers only from a string
    ''' </summary>
    ''' <param name="ValueSrting"></param>
    ''' <returns></returns>
    Public Shared Function GetNumbersFromString(ByVal ValueSrting As String) As String
        'Convert the barcode field to numeric only as we have letter appened to it
        Dim nonNumericCharacters As New System.Text.RegularExpressions.Regex("\D")
        Dim numericOnlyString As String = nonNumericCharacters.Replace(ValueSrting, String.Empty)

        Return numericOnlyString
    End Function
#End Region

#Region " Run SQLSMD "
    Public Shared Function runSQLCMD(ByVal connString As String, ByVal WorkingDirectory As String, ByVal QueryToExceute As String, ByVal ExportFileName As String) As Boolean
        Try

            'https://stackoverflow.com/questions/42733358/how-to-run-sqlcmd-script-using-vb-net
            Dim builder As New Common.DbConnectionStringBuilder

            builder.ConnectionString = connString
            Dim ServerName = builder("Data Source")
            Dim DatabaseName = builder("Initial Catalog")

            _log.Info("Staring sqlcmd script run process")

            Dim DoubleQuote As String = Chr(34)
            Dim Process = New Process()
            Process.StartInfo.UseShellExecute = False
            Process.StartInfo.RedirectStandardOutput = True
            Process.StartInfo.RedirectStandardError = True
            Process.StartInfo.CreateNoWindow = True
            Process.StartInfo.FileName = "SQLCMD.EXE"
            'Dim strargs As String = "SQLCMD.EXE -S " & ServerName & " -d " & DatabaseName & " -E -i """ & QueryToExceute & """ -o " & ExportFileName & "  -h-1 -s"","" -w 700"
            'Process.StartInfo.Arguments = "-S " & ServerName & " -d " & DatabaseName & " -E -i " & DoubleQuote & DoubleQuote & QueryToExceute & DoubleQuote & DoubleQuote & " -o " & ExportFileName & " -h-1 -s"","" -w 700"
            'Process.StartInfo.Arguments = $"-S {ServerName} -d {DatabaseName} -E -i """"{QueryToExceute}"""" -o {ExportFileName} -h-1 -s"","" -w 700"
            Process.StartInfo.Arguments = String.Format("{0} {1} {2} {3} {4} ""{5}"" {6} ""{7}"" {8}", "-S", ServerName, "-d ", DatabaseName, "-E -i ", QueryToExceute, "-o", ExportFileName, "-h -1 -s"","" -w 700")
            'Process.StartInfo.Arguments = "-S " \ {ServerName} \ " -d " \ {DatabaseName} \ " -E -i " & String.Format("\{0}\", QueryToExceute) & " -o " & ExportFileName & "  -h-1 -s"","" -w 700"
            '_log.Info("Arguments sent over to start the DB setup: " & strargs)
            _log.Info("Running the following " & Process.startinfo.filename.ToString)
            _log.Info("With the following Arguments " & Process.StartInfo.Arguments.ToString)
            'Process.StartInfo.WorkingDirectory = {WorkingDirectory}
            _log.Info($"Exporting results in this directory {ExportFileName}")
            Process.Start()
            Process.WaitForExit()
            _log.Info("Finished")

            'Dim startInfo = New ProcessStartInfo()
            'startInfo.FileName = "SQLCMD.EXE"
            'startInfo.Arguments = String.Format("-S {0} -d {1}, -U {2} -P {3} -i {4}", server, database, user, password, File)
            'Process.Start(startInfo)

            Return True

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally


        End Try
    End Function
#End Region

#Region " Convert Numbers to Words "
    Shared Function convertToWords(input As Integer) As String

        'https://stackoverflow.com/questions/45550911/convert-numbers-to-words-using-only-select-case-in-vb-net

        Dim words As String = ""

        Select Case input
            Case 0
                words = "Zero"
            Case 1
                words = "one"
            Case 2
                words = "Two"
            Case 3
                words = "Three"
            Case 4
                words = "Four"
            Case 5
                words = "Five"
            Case 6
                words = "Six"
            Case 7
                words = "Seven"
            Case 8
                words = "Eight"
            Case 9
                words = "Nine"
            Case 10
                words = "Ten"
            Case 11
                words = "Eleven"
            Case 12
                words = "Twelve"
            Case 13
                words = "Thirteen"
            Case 14
                words = "Fourteen"
            Case 15
                words = "Fiftheen"
            Case 16
                words = "Sixteen"
            Case 17
                words = "Seventeen"
            Case 18
                words = "Eighteen"
            Case 19
                words = "Nineteen"
            Case 20
                words = "Twenty"
            Case 30
                words = "Thirty"
            Case 40
                words = "Fourty"
            Case 50
                words = "Fifty"
            Case 60
                words = "Sixty"
            Case 70
                words = "Seventy"
            Case 80
                words = "Eighty"
            Case 90
                words = "Ninety"
            Case 100
                words = "Hundred"

            Case Is >= 1000
                Dim thousands As Integer = (input \ 1000) '<= how many 1000's are there?
                Select Case thousands
                    Case Is > 0
                        input = (input Mod 1000) '<= the remainder is the new value which will be used by calling the same function again
                        words &= convertToWords(thousands) & " thousand " & convertToWords(input)
                End Select
        End Select

        Return words
    End Function
#End Region

#Region " Scan Drivers License "

#End Region

#Region " Split string and keep seprators "

    Public Shared Function SplitAndKeepSeparators(ByVal value As String, ByVal separators As Char(), ByVal splitOptions As StringSplitOptions) As String()
        Dim splitValues As List(Of String) = New List(Of String)()
        Dim itemStart As Integer = 0

        For pos As Integer = 0 To value.Length - 1

            For sepIndex As Integer = 0 To separators.Length - 1

                If separators(sepIndex) = value(pos) Then

                    If itemStart <> pos OrElse splitOptions = StringSplitOptions.None Then
                        splitValues.Add(value.Substring(itemStart, pos - itemStart))
                    End If

                    itemStart = pos + 1
                    splitValues.Add(separators(sepIndex).ToString())
                    Exit For
                End If
            Next
        Next

        If itemStart <> value.Length OrElse splitOptions = StringSplitOptions.None Then
            splitValues.Add(value.Substring(itemStart, value.Length - itemStart))
        End If

        Return splitValues.ToArray()
    End Function

#End Region

#Region " Cancelled input box cancel "
    Public Shared Function StrPtr(ByVal obj As Object) As Integer
        Dim Handle As GCHandle = GCHandle.Alloc(obj, GCHandleType.Pinned)
        Dim intReturn As Integer = Handle.AddrOfPinnedObject.ToInt32
        Handle.Free()
        Return intReturn
    End Function
#End Region

#Region "Check for change"
    Public Shared Function DataChanged(ByVal form As Form) As Boolean
        Dim changed As Boolean = False
        If form Is Nothing Then Return changed

        For Each c As Control In form.Controls

            Select Case c.[GetType]().ToString()
                Case "TextBox"
                    changed = (CType(c, TextBox)).Modified
                Case "TextBox"
                    changed = (CType(c, TextBox)).Modified
            End Select

            If changed Then Exit For
        Next

        Return changed
    End Function
#End Region

#Region " Convert List to Datatable but very slow "
    Public Shared Function ConvertToDataTable(Of T)(ByVal list As List(Of T)) As DataTable
        Dim entityType = GetType(T)

        If entityType = GetType(String) Then
            Dim dataTable = New DataTable(entityType.Name)
            dataTable.Columns.Add(entityType.Name)

            For Each item As T In list
                Dim row = dataTable.NewRow()
                row(0) = item
                dataTable.Rows.Add(row)
            Next

            Return dataTable
        ElseIf entityType.BaseType = GetType([Enum]) Then
            Dim dataTable = New DataTable(entityType.Name)
            dataTable.Columns.Add(entityType.Name)

            For Each namedConstant As String In [Enum].GetNames(entityType)
                Dim row = dataTable.NewRow()
                row(0) = namedConstant
                dataTable.Rows.Add(row)
            Next

            Return dataTable
        End If

        Dim underlyingType = Nullable.GetUnderlyingType(entityType)
        Dim primitiveTypes = New List(Of Type) From {
        GetType(Byte),
        GetType(Char),
        GetType(Decimal),
        GetType(Double),
        GetType(Int16),
        GetType(Int32),
        GetType(Int64),
        GetType(SByte),
        GetType(Single),
        GetType(UInt16),
        GetType(UInt32),
        GetType(UInt64)
    }
        If underlyingType Is Nothing Then underlyingType = entityType
        Dim typeIsPrimitive = primitiveTypes.Contains(underlyingType)

        If typeIsPrimitive Then
            Dim dataTable = New DataTable(underlyingType.Name)
            dataTable.Columns.Add(underlyingType.Name)

            For Each item As T In list
                Dim row = dataTable.NewRow()
                row(0) = item
                dataTable.Rows.Add(row)
            Next

            Return dataTable
        Else
            Dim dataTable = New DataTable(entityType.Name)
            Dim propertyDescriptorCollection = TypeDescriptor.GetProperties(entityType)

            For Each propertyDescriptor As PropertyDescriptor In propertyDescriptorCollection
                Dim propertyType = If(Nullable.GetUnderlyingType(propertyDescriptor.PropertyType), propertyDescriptor.PropertyType)
                dataTable.Columns.Add(propertyDescriptor.Name, propertyType)
            Next

            For Each item As T In list
                Dim row = dataTable.NewRow()

                For Each propertyDescriptor As PropertyDescriptor In propertyDescriptorCollection
                    Try
                        Dim value = propertyDescriptor.GetValue(item)
                        row(propertyDescriptor.Name) = If(value, DBNull.Value)
                    Catch ex As Exception
                        _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
                    End Try
                Next

                dataTable.Rows.Add(row)
            Next

            Return dataTable
        End If
    End Function
#End Region

#Region " Check Remote File Exists "
    Public Function RemoteFileExists(ByVal fileurl As String) As Boolean
        Try
            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(fileurl), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.GetFileSize
            Dim response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
            If response.StatusCode = FtpStatusCode.ActionNotTakenFileUnavailable Then
                Return False 'Return instead of Exit Function
            End If
            Dim fileSize As Long = response.ContentLength
            MsgBox(fileSize)
            If fileSize > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception 'Catch all errors
            'Log the error if you'd like, you can find the error message and location in "ex.Message" and "ex.StackTrace".
            MessageBox.Show("An error occurred:" & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False 'Return False since the checking failed.
        End Try
    End Function
#End Region

#Region " Get Mac Address "
    Public Shared Function getMacAddress() As String
        Try
            Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
            Dim adapter As NetworkInterface
            Dim myMac As String = String.Empty

            For Each adapter In adapters
                Select Case adapter.NetworkInterfaceType
                'Exclude Tunnels, Loopbacks and PPP
                    Case NetworkInterfaceType.Tunnel, NetworkInterfaceType.Loopback, NetworkInterfaceType.Ppp
                    Case Else
                        If Not adapter.GetPhysicalAddress.ToString = String.Empty And Not adapter.GetPhysicalAddress.ToString = "00000000000000E0" Then
                            myMac = adapter.GetPhysicalAddress.ToString
                            Exit For ' Got a mac so exit for
                        End If

                End Select
            Next adapter

            Return myMac
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
#End Region

#Region " Validate US Date "
    Public Shared Function ValidateUSDate(ByVal checkInputValue As String) As Boolean
        Dim returnError As Boolean
        Dim dateVal As DateTime

        ' several possible format styles
        Dim formats() As String = {"MM-d-yyyy", "MM-dd-yyyy", "M-dd-yyyy", "M-d-yyyy"}

        If Date.TryParseExact(checkInputValue, formats, System.Globalization.CultureInfo.CurrentCulture, DateTimeStyles.None, dateVal) Then
            returnError = True
        Else
            returnError = False
        End If
        Return returnError
    End Function
#End Region

#Region " Get Datatable from Objects "
    Public Shared Function GetDataTableFromObjects(Of TDataClass As Class)(ByVal dataList As List(Of TDataClass)) As DataTable
        Dim t As Type = GetType(TDataClass)
        Dim dt As DataTable = New DataTable(t.Name)

        For Each pi As PropertyInfo In t.GetProperties()
            dt.Columns.Add(New DataColumn(pi.Name))
        Next

        If dataList IsNot Nothing Then

            For Each item As TDataClass In dataList
                Dim dr As DataRow = dt.NewRow()

                For Each dc As DataColumn In dt.Columns
                    dr(dc.ColumnName) = item.[GetType]().GetProperty(dc.ColumnName).GetValue(item, Nothing)
                Next

                dt.Rows.Add(dr)
            Next
        End If

        Return dt
    End Function
#End Region

#Region " Create Data Tables "
    Public Shared Function CreateDataTable(Of T)(ByVal list As IEnumerable(Of T)) As DataTable
        Dim type As Type = GetType(T)
        Dim properties = type.GetProperties()
        Dim dataTable As DataTable = New DataTable()
        dataTable.TableName = GetType(T).FullName
        _log.Info("Table Name: " & GetType(T).FullName)

        For Each info As PropertyInfo In properties
            dataTable.Columns.Add(New DataColumn(info.Name, If(Nullable.GetUnderlyingType(info.PropertyType), info.PropertyType)))
            _log.Info("Column Header: " & info.Name)
        Next

        For Each entity As T In list
            Dim values As Object() = New Object(properties.Length - 1) {}

            For i As Integer = 0 To properties.Length - 1
                values(i) = properties(i).GetValue(entity)
                '_log.Info("Field: " & values(i).ToString)
            Next

            dataTable.Rows.Add(values)
            '_log.Info("Fields: " & values.ToString)
        Next

        Return dataTable
    End Function
    Public Shared Function ToDataTable(Of T)(ByVal self As IEnumerable(Of T)) As DataTable
        Dim properties = GetType(T).GetProperties()
        Dim dataTable = New DataTable()

        For Each Info As PropertyInfo In properties
            dataTable.Columns.Add(Info.Name, If(Nullable.GetUnderlyingType(Info.PropertyType), Info.PropertyType))
            _log.Info("Column Header: " & Info.Name)
        Next

        For Each entity As T In self
            dataTable.Rows.Add(properties.[Select](Function(p) p.GetValue(entity)).ToArray())
            _log.Info("Fields: " & properties.[Select](Function(p) p.GetValue(entity)).ToArray().ToString)
        Next

        Return dataTable
    End Function
#End Region

#Region " Set the combobox to the correct item"
    Public Shared Sub FindItemByID(ByRef cbCombo As ComboBox, ByVal strID As String)
        Dim bLoading = False
        ' This sub is used to find an Item in a combobox and set the selected index of the combo box to that item...
        Dim bOrigLoading As Boolean = bLoading ' So I can restore the old value at the end...
        bLoading = True ' Stops this trigger when loading the data into the ComboBox...
        Dim bFound As Boolean
        Dim iCount As Integer

        ' We are searching on the ValueMember of the ComboBox...
        For iCount = 0 To cbCombo.Items.Count - 1
            cbCombo.SelectedIndex = iCount ' Set the SelectedIndex to move items in the ComboBox...
            If cbCombo.SelectedValue = strID Then ' We have found what we are looking for...
                bFound = True ' Flag we have found what we are looking for...
                Exit For ' And exit the For..Next loop...
            End If
        Next

        ' Did we find it?...
        If bFound Then ' Yes...
            cbCombo.SelectedIndex = iCount
        Else ' Nope...
            ' Setting it to -1 once does not seem to work, have to do it a second time...
            cbCombo.SelectedIndex = -1
            cbCombo.SelectedIndex = -1
        End If
        bLoading = bOrigLoading ' Restore my original Loading flag...
    End Sub
#End Region

    '#Region " Change Controls to True or False "
    '    Public Sub ControlsTrueFalse(ByVal Status As Boolean)
    '        Try

    '            'Load blank defaults
    '            Dim a As Control
    '            For Each a In Me.Controls
    '                _log.Info("Changing control " & a.GetType.Name & "to the status of " & Status)
    '                If TypeOf a Is TextBox Then
    '                    a.Enabled = Status
    '                End If
    '                If TypeOf a Is ComboBox Then
    '                    a.Enabled = Status
    '                End If
    '                If TypeOf a Is DateTimePicker Then
    '                    a.Enabled = Status
    '                End If
    '                If TypeOf a Is CheckBox Then
    '                    a.Enabled = Status
    '                End If
    '                If TypeOf a Is GroupBox Then
    '                    a.Enabled = Status
    '                End If
    '                If TypeOf a Is Button Then
    '                    a.Enabled = Status
    '                End If
    '            Next
    '            For Each GroupBoxCntrol As Control In Me.Controls
    '                If TypeOf GroupBoxCntrol Is GroupBox Then
    '                    _log.Info("Changing control to " & Status & " in a groupbox except of name of group box is Searches")
    '                    If GroupBoxCntrol.Name <> "Searches" Then
    '                        For Each cntrl As Control In GroupBoxCntrol.Controls
    '                            _log.Info("Changing control " & cntrl.GetType.Name & "to the status of " & Status)
    '                            If TypeOf cntrl Is TextBox Then
    '                                cntrl.Enabled = Status
    '                            End If
    '                            If TypeOf cntrl Is ComboBox Then
    '                                cntrl.Enabled = Status
    '                            End If
    '                            If TypeOf cntrl Is DateTimePicker Then
    '                                cntrl.Enabled = Status
    '                            End If
    '                            If TypeOf cntrl Is CheckBox Then
    '                                cntrl.Enabled = Status
    '                            End If
    '                            If TypeOf cntrl Is GroupBox Then
    '                                cntrl.Enabled = Status
    '                            End If
    '                            If TypeOf cntrl Is Button Then
    '                                cntrl.Enabled = Status
    '                            End If
    '                        Next
    '                        GroupBoxCntrol.Enabled = Status
    '                    End If
    '                End If

    '            Next

    '        Catch ex As Exception
    '            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
    '            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Finally

    '        End Try
    '    End Sub
    '#End Region

#Region " Simple FTP Upload File"
    ''' <summary>
    ''' Simple Upload file subroutine
    ''' Format "ftp://www.yourwebsitename.com/yourfilename.fileextension"
    ''' </summary>
    ''' <param name="filetoupload"></param>
    ''' <param name="ftpuri"></param>
    ''' <param name="ftpusername"></param>
    ''' <param name="ftppassword"></param>
    Public Shared Sub FtpUploadFile(ByVal filetoupload As String, ByVal ftpuri As String, ByVal ftpusername As String, ByVal ftppassword As String)
        ' Create a web request that will be used to talk with the server and set the request method to upload a file by ftp.
        Dim ftpRequest As FtpWebRequest = CType(WebRequest.Create(ftpuri), FtpWebRequest)

        Try
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile

            ' Confirm the Network credentials based on the user name and password passed in.
            ftpRequest.Credentials = New NetworkCredential(ftpusername, ftppassword)

            ' Read into a Byte array the contents of the file to be uploaded 
            Dim bytes() As Byte = System.IO.File.ReadAllBytes(filetoupload)

            ' Transfer the byte array contents into the request stream, write and then close when done.
            ftpRequest.ContentLength = bytes.Length
            Using UploadStream As Stream = ftpRequest.GetRequestStream()
                UploadStream.Write(bytes, 0, bytes.Length)
                UploadStream.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        MessageBox.Show("Process Complete")
    End Sub
#End Region

#Region " Simple FTP Downlad File "
    ''' <summary>
    ''' Simple Upload file subroutine
    ''' format "ftp://ftp.yourwebsitename/nameoffileonserver.fileext"
    ''' </summary>
    ''' <param name="downloadpath"></param>
    ''' <param name="ftpuri"></param>
    ''' <param name="ftpusername"></param>
    ''' <param name="ftppassword"></param>
    Public Shared Sub FTPDownloadFile(ByVal downloadpath As String, ByVal ftpuri As String, ByVal ftpusername As String, ByVal ftppassword As String)
        'Create a WebClient.
        Dim request As New WebClient()

        ' Confirm the Network credentials based on the user name and password passed in.
        request.Credentials = New NetworkCredential(ftpusername, ftppassword)

        'Read the file data into a Byte array
        Dim bytes() As Byte = request.DownloadData(ftpuri)

        Try
            '  Create a FileStream to read the file into
            Dim DownloadStream As FileStream = IO.File.Create(downloadpath)
            '  Stream this data into the file
            DownloadStream.Write(bytes, 0, bytes.Length)
            '  Close the FileStream
            DownloadStream.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        MessageBox.Show("Process Complete")

    End Sub
#End Region

End Class
