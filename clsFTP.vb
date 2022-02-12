'----------------------------------------------------------------
'       Class       :   clsFTP.vb
'       Description :   This class will enable a developer to
'                       perform FTP Processing in VB.NET.
'                       The class supports such features as:
'                           - Uploading a file.
'                           - Create/Remove a directory.
'                           - Remove a file.
'                           - and much much more...
'       Date        :   7th February 2006.
'       Conversion  :   The code was converted from C# code.
'       Note        :   Also note that the code DOES NOT
'                       ACCOMODATE for proxy servers.  Only
'                       DIRECT connections to FTP Sites.
'
'   Bit Shifting within VB.NET 2002 is not supported via the
'   traditional bit shifting operators (i.e. << - Bitwise left
'   and >> - Bitwise right).
'   Dividing a number by 2^16 in VB.NET 2002 is the same as bit-
'   shifting right 16 positions.
'   Multiplying a number by 2^16 in VB.NET 2002 is the same as
'   bit-shifting left 16 positions.
'
'   Check out the following MSDN Article on bitshifting in 
'   VB.NET 2003:
'       -   Visual Basic .NET 2003 Language Changes
'       -   By Duncan Mackenzie.
'----------------------------------------------------------------
Imports System
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Net.Sockets
Imports log4net
Imports log4net.Config

Namespace FTP

    ' Main FTP Class.
    Public Class clsFTP

#Region "Main Class Variable Declarations"
        Private m_sRemoteHost, m_sRemotePath, m_sRemoteUser As String
        Private m_sRemotePassword, m_sMess As String
        Private m_iRemotePort, m_iBytes As Int32
        Private m_objClientSocket As Socket

        Private m_iRetValue As Int32
        Private m_bLoggedIn As Boolean    ' Change to loggedIn
        Private m_sMes, m_sReply As String

        ' Set the size of the packet that is used to read and
        '  write data to the FTP Server to the spcified size below.
        Public Const BLOCK_SIZE = 512
        Private m_aBuffer(BLOCK_SIZE) As Byte
        Private ASCII As Encoding = Encoding.ASCII

        ' General variables
        Private m_sMessageString As String
        Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsFTP))
#End Region

#Region "Class Constructors"
        '
        ' Main class constructor.
        Public Sub New()
            m_sRemoteHost = "microsoft"
            m_sRemotePath = "."
            m_sRemoteUser = "anonymous"
            m_sRemotePassword = ""
            m_sMessageString = ""
            m_iRemotePort = 21
            m_bLoggedIn = False
        End Sub

        ' Parametized constructor.
        Public Sub New(ByVal sRemoteHost As String, _
                       ByVal sRemotePath As String, _
                       ByVal sRemoteUser As String, _
                       ByVal sRemotePassword As String, _
                       ByVal iRemotePort As Int32)
            m_sRemoteHost = sRemoteHost
            m_sRemotePath = sRemotePath
            m_sRemoteUser = sRemoteUser
            m_sRemotePassword = sRemotePassword
            m_sMessageString = ""
            m_iRemotePort = iRemotePort
            m_bLoggedIn = False
        End Sub
#End Region

#Region "Public Properties"
        '
        ' Set/Get the name of the FTP Server.
        Public Property RemoteHost() As String
            Get
                Return m_sRemoteHost
            End Get
            Set(ByVal Value As String)
                m_sRemoteHost = Value
            End Set
        End Property

        ' Set/Get the FTP Port Number.
        Public Property RemotePort() As Int32
            Get
                Return m_iRemotePort
            End Get
            Set(ByVal Value As Int32)
                m_iRemotePort = Value
            End Set
        End Property

        ' Set/Get the remote path.
        Public Property RemotePath() As String
            Get
                Return m_sRemotePath
            End Get
            Set(ByVal Value As String)
                m_sRemotePath = Value
            End Set
        End Property

        ' Set the remote password.
        Public Property RemotePassword() As String
            Get
                Return m_sRemotePassword
            End Get
            Set(ByVal Value As String)
                m_sRemotePassword = Value
            End Set
        End Property

        ' Set/Get the remote user.
        Public Property RemoteUser() As String
            Get
                Return m_sRemoteUser
            End Get
            Set(ByVal Value As String)
                m_sRemoteUser = Value
            End Set
        End Property

        ' Set the class messagestring.
        Public Property MessageString() As String
            Get
                Return m_sMessageString
            End Get
            Set(ByVal Value As String)
                m_sMessageString = Value
            End Set
        End Property

#End Region

#Region "Public Subs and Functions"
        '
        ' Return a list of files within a string() array from the
        '  file system.
        Public Function GetFileList(ByVal sMask As String) As String()
            Dim cSocket As Socket
            Dim bytes As Int32
            Dim seperator As Char = ControlChars.Lf
            Dim mess() As String

            m_sMes = ""
            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            cSocket = CreateDataSocket()
            SendCommand("NLST " & sMask)

            If (Not (m_iRetValue = 150 Or m_iRetValue = 125)) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            m_sMes = ""
            Do While (True)
                m_aBuffer.Clear(m_aBuffer, 0, m_aBuffer.Length)
                bytes = cSocket.Receive(m_aBuffer, m_aBuffer.Length, 0)
                m_sMes += ASCII.GetString(m_aBuffer, 0, bytes)

                If (bytes < m_aBuffer.Length) Then
                    Exit Do
                End If
            Loop

            mess = m_sMes.Split(seperator)
            cSocket.Close()
            ReadReply()

            If (m_iRetValue <> 226) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))

            End If

            Return mess
        End Function

        '
        ' Get the size of the file on the FTP Server.
        Public Function GetFileSize(ByVal sFileName As String) As Long
            Dim size As Long

            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SendCommand("SIZE " & sFileName)
            size = 0

            If (m_iRetValue = 213) Then
                size = Int64.Parse(m_sReply.Substring(4))
            ElseIf (m_iRetValue = 550) Then
                'do nothing as file is not there
                size = 0
            Else
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
                _log.Error(m_sReply.Substring(4))
            End If

            _log.Info("File Size: " & size)

            Return size
        End Function

        '
        ' Log into the FTP Server.
        Public Function Login() As Boolean
            m_objClientSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            Dim ep As New IPEndPoint(Dns.Resolve(m_sRemoteHost).AddressList(0), m_iRemotePort)

            Try
                m_objClientSocket.Connect(ep)
            Catch ex As Exception
                MessageString = m_sReply
                Throw New IOException("Couldn't connect to remote server")
            End Try

            ReadReply()
            If (m_iRetValue <> 220) Then
                CloseConnection()
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            SendCommand("USER " & m_sRemoteUser)
            If (Not (m_iRetValue = 331 Or m_iRetValue = 230)) Then
                Cleanup()
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            If (m_iRetValue <> 230) Then
                SendCommand("PASS " & m_sRemotePassword)
                If (Not (m_iRetValue = 230 Or m_iRetValue = 202)) Then
                    Cleanup()
                    MessageString = m_sReply
                    Throw New IOException(m_sReply.Substring(4))
                End If
            End If

            m_bLoggedIn = True
            ChangeDirectory(m_sRemotePath)

            ' Return the end result.
            Return m_bLoggedIn
        End Function

        '
        ' If the value of mode is true, set binary mode for 
        '  downloads.
        ' Else, set Ascii mode.
        Public Sub SetBinaryMode(ByVal bMode As Boolean)

            If (bMode) Then
                SendCommand("TYPE I")
            Else
                SendCommand("TYPE A")
            End If

            If (m_iRetValue <> 200) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If
        End Sub

        '
        ' Download a file to the Assembly's local directory,
        ' keeping the same file name.
        Public Sub DownloadFile(ByVal sFileName As String)
            DownloadFile(sFileName, "", False)
        End Sub

        '
        ' Download a remote file to the Assembly's local 
        '  directory, keeping the same file name, and set 
        '  the resume flag.
        Public Sub DownloadFile(ByVal sFileName As String, _
                                ByVal bResume As Boolean)
            DownloadFile(sFileName, "", bResume)
        End Sub

        '
        ' Download a remote file to a local file name which can 
        '  include a path. The local file name will be created or 
        '  overwritten, but the path must exist.
        Public Sub DownloadFile(ByVal sFileName As String, _
                                ByVal sLocalFileName As String)
            DownloadFile(sFileName, sLocalFileName, False)
        End Sub

        '
        ' Download a remote file to a local file name which can 
        '  include a path, and set the resume flag. The local file 
        '  name will be created or overwritten, but the path must 
        '  exist.
        Public Sub DownloadFile(ByVal sFileName As String, _
                                ByVal sLocalFileName As String, _
                                ByVal bResume As Boolean)
            Dim st As Stream
            Dim output As FileStream
            Dim cSocket As Socket
            Dim offset, npos As Long

            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SetBinaryMode(True)

            If (sLocalFileName.Equals("")) Then
                sLocalFileName = sFileName
            End If

            If (Not (File.Exists(sLocalFileName))) Then
                st = File.Create(sLocalFileName)
                st.Close()
            End If

            output = New FileStream(sLocalFileName, FileMode.Open)
            cSocket = CreateDataSocket()
            offset = 0

            If (bResume) Then
                offset = output.Length

                If (offset > 0) Then
                    SendCommand("REST " & offset)
                    If (m_iRetValue <> 350) Then
                        'throw new IOException(reply.Substring(4));
                        'Some servers may not support resuming.
                        offset = 0
                    End If
                End If

                If (offset > 0) Then
                    npos = output.Seek(offset, SeekOrigin.Begin)
                End If
            End If

            SendCommand("RETR " & sFileName)

            If (Not (m_iRetValue = 150 Or m_iRetValue = 125)) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            Do While (True)
                m_aBuffer.Clear(m_aBuffer, 0, m_aBuffer.Length)
                m_iBytes = cSocket.Receive(m_aBuffer, m_aBuffer.Length, 0)
                output.Write(m_aBuffer, 0, m_iBytes)

                If (m_iBytes <= 0) Then
                    Exit Do
                End If
            Loop

            output.Close()
            If (cSocket.Connected) Then
                cSocket.Close()
            End If

            ReadReply()
            If (Not (m_iRetValue = 226 Or m_iRetValue = 250)) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

        End Sub

        '
        ' Upload a file.
        Public Sub UploadFile(ByVal sFileName As String)
            UploadFile(sFileName, False)
        End Sub

        '
        ' Upload a file and set the resume flag.
        Public Sub UploadFile(ByVal sFileName As String, _
                              ByVal bResume As Boolean)
            Dim cSocket As Socket
            Dim offset As Long
            Dim input As FileStream
            Dim bFileNotFound As Boolean

            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            cSocket = CreateDataSocket()
            offset = 0

            If (bResume) Then
                Try
                    SetBinaryMode(True)
                    offset = GetFileSize(sFileName)
                Catch ex As Exception
                    offset = 0
                End Try
            End If

            If (offset > 0) Then
                SendCommand("REST " & offset)
                If (m_iRetValue <> 350) Then
                    'throw new IOException(reply.Substring(4));
                    'Remote server may not support resuming.
                    offset = 0
                End If
            End If

            SendCommand("STOR " & Path.GetFileName(sFileName))
            If (Not (m_iRetValue = 125 Or m_iRetValue = 150)) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            ' Check to see if the file exists before the upload.
            bFileNotFound = False
            If (File.Exists(sFileName)) Then
                ' Open input stream to read source file
                input = New FileStream(sFileName, FileMode.Open)
                If (offset <> 0) Then
                    input.Seek(offset, SeekOrigin.Begin)
                End If

                ' Upload the file 
                m_iBytes = input.Read(m_aBuffer, 0, m_aBuffer.Length)
                Do While (m_iBytes > 0)
                    cSocket.Send(m_aBuffer, m_iBytes, 0)
                    m_iBytes = input.Read(m_aBuffer, 0, m_aBuffer.Length)
                Loop
                input.Close()
            Else
                bFileNotFound = True
            End If

            If (cSocket.Connected) Then
                cSocket.Close()
            End If

            ' No point in reading the return value if the file was
            '  not found.
            If (bFileNotFound) Then
                MessageString = m_sReply
                Throw New IOException("The file: " & sFileName & " was not found.  Can not upload the file to the FTP Site.")
            End If

            ReadReply()
            If (Not (m_iRetValue = 226 Or m_iRetValue = 250)) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If
        End Sub

        '
        ' Delete a file from the remote FTP server.
        Public Function DeleteFile(ByVal sFileName As String) As Boolean
            Dim bResult As Boolean

            bResult = True
            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SendCommand("DELE " & sFileName)
            If (m_iRetValue <> 250) Then
                bResult = False
                MessageString = m_sReply
            End If

            ' Return the final result.
            Return bResult
        End Function

        '
        ' Rename a file on the remote FTP server.
        Public Function RenameFile(ByVal sOldFileName As String, _
                                   ByVal sNewFileName As String) As Boolean
            Dim bResult As Boolean

            bResult = True
            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SendCommand("RNFR " & sOldFileName)
            If (m_iRetValue <> 350) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            '  known problem
            '  rnto will not take care of existing file.
            '  i.e. It will overwrite if newFileName exist
            SendCommand("RNTO " & sNewFileName)
            If (m_iRetValue <> 250) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            Return bResult
        End Function

        '
        ' Create a directory on the remote FTP server.
        Public Function CreateDirectory(ByVal sDirName As String) As Boolean
            Dim bResult As Boolean

            bResult = True
            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SendCommand("MKD " & sDirName)
            If (m_iRetValue <> 257) Then
                bResult = False
                MessageString = m_sReply
            End If

            ' Return the final result.
            Return bResult
        End Function

        '
        ' Delete a directory on the remote FTP server.
        Public Function RemoveDirectory(ByVal sDirName As String) As Boolean
            Dim bResult As Boolean

            bResult = True
            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SendCommand("RMD " & sDirName)
            If (m_iRetValue <> 250) Then
                bResult = False
                MessageString = m_sReply
            End If

            ' Return the final result.
            Return bResult
        End Function

        '
        ' Change the current working directory on the remote FTP 
        '  server.
        Public Function ChangeDirectory(ByVal sDirName As String) As Boolean
            Dim bResult As Boolean

            bResult = True
            If (sDirName.Equals(".")) Then
                Exit Function
            End If

            If (Not (m_bLoggedIn)) Then
                Login()
            End If

            SendCommand("CWD " & sDirName)
            If (m_iRetValue <> 250) Then
                bResult = False
                MessageString = m_sReply
            End If

            Me.m_sRemotePath = sDirName

            ' Return the final result.
            Return bResult
        End Function

        '
        ' Close the FTP connection.
        Public Sub CloseConnection()
            If (Not (m_objClientSocket Is Nothing)) Then
                SendCommand("QUIT")
            End If

            Cleanup()
        End Sub

#End Region

#Region "Private Subs and Functions"
        '
        ' Read the reply from the FTP Server
        Private Sub ReadReply()
            m_sMes = ""
            m_sReply = ReadLine()
            m_iRetValue = Int32.Parse(m_sReply.Substring(0, 3))
        End Sub

        '
        ' Clean up some variables.
        Private Sub Cleanup()
            If Not (m_objClientSocket Is Nothing) Then
                m_objClientSocket.Close()
                m_objClientSocket = Nothing
            End If

            m_bLoggedIn = False
        End Sub

        '
        ' Read a line from the server.
        Private Function ReadLine(Optional ByVal bClearMes As Boolean = False) As String
            Dim seperator As Char = ControlChars.Lf
            Dim mess() As String

            If (bClearMes) Then
                m_sMes = ""
            End If
            Do While (True)
                m_aBuffer.Clear(m_aBuffer, 0, BLOCK_SIZE)
                m_iBytes = m_objClientSocket.Receive(m_aBuffer, m_aBuffer.Length, 0)
                m_sMes += ASCII.GetString(m_aBuffer, 0, m_iBytes)
                If (m_iBytes < m_aBuffer.Length) Then
                    Exit Do
                End If
            Loop

            mess = m_sMes.Split(seperator)
            If (m_sMes.Length > 2) Then
                m_sMes = mess(mess.Length - 2)
            Else
                m_sMes = mess(0)
            End If

            If (Not (m_sMes.Substring(3, 1).Equals(" "))) Then
                Return ReadLine(True)
            End If

            Return m_sMes
        End Function

        '
        ' Send a command to the FTP Server.
        Private Sub SendCommand(ByVal sCommand As String)
            sCommand = sCommand & ControlChars.CrLf
            Dim cmdbytes As Byte() = ASCII.GetBytes(sCommand)

            m_objClientSocket.Send(cmdbytes, cmdbytes.Length, 0)
            ReadReply()
        End Sub

        '
        ' Create a Data socket.
        Private Function CreateDataSocket() As Socket
            Dim index1, index2, len As Int32
            Dim partCount, i, port As Int32
            Dim ipData, buf, ipAddress As String
            Dim parts(6) As Int32
            Dim ch As Char
            Dim s As Socket
            Dim ep As IPEndPoint

            SendCommand("PASV")
            If (m_iRetValue <> 227) Then
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If

            index1 = m_sReply.IndexOf("(")
            index2 = m_sReply.IndexOf(")")
            ipData = m_sReply.Substring(index1 + 1, index2 - index1 - 1)

            len = ipData.Length
            partCount = 0
            buf = ""

            For i = 0 To ((len - 1) And partCount <= 6)
                ch = Char.Parse(ipData.Substring(i, 1))
                If (Char.IsDigit(ch)) Then
                    buf += ch
                ElseIf (ch <> ",") Then
                    MessageString = m_sReply
                    Throw New IOException("Malformed PASV reply: " & m_sReply)
                End If

                If ((ch = ",") Or (i + 1 = len)) Then
                    Try
                        parts(partCount) = Int32.Parse(buf)
                        partCount += 1
                        buf = ""
                    Catch ex As Exception
                        MessageString = m_sReply
                        Throw New IOException("Malformed PASV reply: " & m_sReply)
                    End Try
                End If
            Next

            ipAddress = parts(0) & "." & parts(1) & "." & parts(2) & "." & parts(3)

            ' Make this call in VB.NET 2002.  We would like to 
            '  bitshift the number by 8 bits, so in VB.NET 2002 we
            '  multiply the number by 2 to the power of 8.
            port = parts(4) * (2 ^ 8)

            ' Make this call and comment out the above line for
            '  VB.NET 2003.
            'port = parts(4) << 8

            ' Determine the data port number.
            port = port + parts(5)

            s = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            ep = New IPEndPoint(Dns.Resolve(ipAddress).AddressList(0), port)

            Try
                s.Connect(ep)
            Catch ex As Exception
                MessageString = m_sReply
                Throw New IOException("Can't connect to remote server")
            End Try

            Return s
        End Function

#End Region

#Region " *** Example Code *** "
        '
        ' Copy and paste the code below into a VB WebForm or WinForm
        '  application and then do the following:
        '
        '       1).  From within the ASP.NET or WinForm app set a
        '            reference to the FTP.dll and BitOperators.dll
        '            files.
        '       2).  At the top of the application code file 
        '            (E.g WebForm1.aspx.vb or Form1.vb) type in
        '               Imports FTP
        '       3).  Compile the application and run.
        '       4).  Have fun.

        'Protected Sub TestFTP()
        '    Dim ff As clsFTP

        '    Try
        '        '-------------------------------------------
        '        ' OPTION 1
        '        ' --------
        '        '
        '        ' Create an instance of the FTP Class.
        '        'ff = New clsFTP()

        '        ' Setup the appropriate properties.
        '        'ff.RemoteHost = "microsoft"
        '        'ff.RemoteUser = "ftpuser"
        '        'ff.RemotePassword = "password"
        '        '-------------------------------------------

        '        '-------------------------------------------
        '        ' OPTION 2
        '        ' --------
        '        '  Pass the values into the constructor 
        '        '  instead.  These can be overridden by simply
        '        '  setting the appropriate properties on the
        '        '  instance of the clsFTP Class.
        '        ff = New clsFTP("microsoft", _
        '                        ".", _
        '                        "ftpuser", _
        '                        "password", _
        '                        21)

        '        ' Attempt to log into the FTP Server.
        '        If (ff.Login()) Then
        '            '
        '            ' Move the to Area1\Section1\Subby1\ directory.
        '            ff.ChangeDirectory("Area1")
        '            ff.ChangeDirectory("Section1")

        '            'ff.CreateDirectory("Subby1")
        '            ff.ChangeDirectory("Subby1")
        '            ff.SetBinaryMode(True)

        '            ' Upload a file.
        '            'ff.UploadFile("d:\general\secureapps.pdf")

        '            ' Download a file.
        '            'ff.DownloadFile("secureapps.pdf", "d:\general\secureapps.pdf")

        '            ' Remove a file from the FTP Site.
        '            If (ff.DeleteFile("secureapps.pdf")) Then
        '                Response.Write("File has been removed from FTP Site")
        '                'MessageBox.Show("File has been removed from FTP Site")
        '            Else
        '                Response.Write("Unable to remove file from FTP Site.  Message from server: " & ff.MessageString & "<br>")
        '                'MessageBox.Show("Unable to remove file from FTP Site")
        '            End If

        '            ' Rename a file on the FTP Site.
        '            'If (ff.RenameFile("secureapps.pdf", "newapp.pdf")) Then
        '            '    Response.Write("File has been renamed")
        '            '    MessageBox.Show("File has been renamed")
        '            'End If

        '            'ff.ChangeDirectory("..")
        '            'If (ff.RemoveDirectory("Subby1")) Then
        '            '    Response.Write("Directory has been removed<br>")
        '            '    ' MessageBox.Show("Directory has been removed")
        '            'Else
        '            '    Response.Write("Unable to remove the directory.  Message from server: " & ff.MessageString & "<br>")
        '            '    ' MessageBox.Show("Unable to remove the directory.")
        '            'End If
        '        End If

        '    Catch ex As System.Exception
        '        ' ASP.NET
        '        Response.Write(ex.Message & "<BR>")
        '        Response.Write("Message from FTP Server was: " & ff.MessageString)

        '        ' WinForms
        '        'Messagebox.Show(ex.Message)
        '        'MessageBox.show("Message from FTP Server was: " & ff.MessageString)
        '    Finally
        '        '
        '        ' Always close down the connection to ensure that
        '        '  there are no "stray" Fido's Fetching data.  In
        '        '  other words, no stray/limbo/not-in-use FTP
        '        '  connections.
        '        ff.CloseConnection()
        '    End Try
        'End Sub

#End Region

    End Class
End Namespace

