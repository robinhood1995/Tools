Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports MySql.Data.MySqlClient
Imports log4net
Imports log4net.Config
Imports System.Configuration
Imports System.Reflection

Namespace Onling
    ''' <summary>
    ''' Onling.com
    ''' Use and modify at will...
    ''' This was intended to replace all the silly functions that I had...
    ''' </summary>
    ''' <remarks></remarks>
#Region "Database Connection Class"
    Public Class clsDatabase
        ''' <summary>
        ''' A Database connection for .NET
        ''' </summary>
        ''' <remarks>
        ''' This class does not hold open a connection but 
        ''' instead is stateless: for each request it 
        ''' connects, performs the request and disconnects.
        ''' </remarks>
        Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsDatabase))

#Region "LOCAL DECLARATIONS"
        Private _Sqlhostname As String
        Private _Sqlusername As String
        Private _Sqlpassword As String
        Private _SqlDBName As String
        Private _MySqlhostname As String
        Private _MySqlusername As String
        Private _MySqlpassword As String
        Private _MySqlDBName As String
        Private _MySqlport As String
        Private _Odbchostname As String
        Private _Odbcusername As String
        Private _Odbcpassword As String
        Private _OdbcDBName As String
        Private _Odbcport As String
        Private _OdbcDSN As String
        Public MsSqlDBConn As New SqlConnection
        Public MySqlDBConn As New MySqlConnection
        Public ODBCDBConn As New OdbcConnection
#End Region

#Region "CONSTRUCTORS"
        ''' <summary>
        '''  Manual Constructor
        ''' </summary>
        ''' <remarks>All has to be passed manually</remarks>
        Sub New()
            _log.Info("Starting " & MethodBase.GetCurrentMethod().ToString())
        End Sub
#End Region

#Region "Sql Properties"
        ''' <summary>
        ''' Sql Hostname
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SqlHostname() As String
            Get
                Return _Sqlhostname
            End Get
            Set(ByVal value As String)
                _Sqlhostname = value
            End Set
        End Property

        ''' <summary>
        ''' Sql Username
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property SqlUsername() As String
            Get
                Return _Sqlusername
            End Get
            Set(ByVal value As String)
                _Sqlusername = value
            End Set
        End Property

        ''' <summary>
        ''' Sql Password
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SqlPassword() As String
            Get
                Return _Sqlpassword
            End Get
            Set(ByVal value As String)
                _Sqlpassword = value
            End Set
        End Property

        ''' <summary>
        ''' Sql DBName
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SqlDBName() As String
            Get
                Return _SqlDBName
            End Get
            Set(ByVal value As String)
                _SqlDBName = value
            End Set
        End Property

#End Region

#Region "MsSQL Connection"
        ''' <summary>
        ''' Makes a Connection to MsSql
        ''' </summary>
        ''' <param name="Trusted">True or False if you have a trusted connection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SqlConn(Optional ByVal Trusted As Boolean = True) As SqlConnection
            Try
                'Dim MsSqlDBConn As New SqlConnection
                If Trusted Then
                    MsSqlDBConn.ConnectionString = "data source=" & _Sqlhostname & ";" & _
                                            "Initial catalog=""" & _SqlDBName & """;" & _
                                            "Integrated Security=True;" & _
                                            "MultipleActiveResultSets=True;"
                    'Network Library=DBMSSOCN;"
                    _log.Info(Environment.UserName & " Opening Database Trusted IP: " & _Sqlhostname & " DBName " & _SqlDBName)
                Else
                    MsSqlDBConn.ConnectionString = "data source=" & _Sqlhostname & ";" & _
                        "Initial catalog=" & _SqlDBName & ";" & _
                        "User ID=" & _Sqlusername & ";" & _
                        "Password=" & _Sqlpassword & ";" & _
                        "MultipleActiveResultSets=True;"
                    _log.Info(Environment.UserName & " Opening Database IP: " & _Sqlhostname & " DBName " & _SqlDBName)
                    _log.Info(Environment.UserName & " Opened Database Username : " & _Sqlusername & " Password " & _Sqlpassword)
                End If

                MsSqlDBConn.Open()
                Return MsSqlDBConn

            Catch ex As SqlException
                Throw New ApplicationException("Connection to " & _Sqlhostname & " was not Successfull.")
                _log.Error("Connection to " & _Sqlhostname & " was not Successfull.")
            End Try
        End Function

#End Region

#Region "MySql Properties"
        ''' <summary>
        ''' MySql Hostname
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MySqlHostname() As String
            Get
                Return _MySqlhostname
            End Get
            Set(ByVal value As String)
                _MySqlhostname = value
            End Set
        End Property

        ''' <summary>
        ''' MySql Username
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property MySqlUsername() As String
            Get
                Return _MySqlusername
            End Get
            Set(ByVal value As String)
                _MySqlusername = value
            End Set
        End Property

        ''' <summary>
        ''' MySql Password
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MySqlPassword() As String
            Get
                Return _MySqlpassword
            End Get
            Set(ByVal value As String)
                _MySqlpassword = value
            End Set
        End Property

        ''' <summary>
        ''' MySql DBName
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MySqlDBName() As String
            Get
                Return _MySqlDBName
            End Get
            Set(ByVal value As String)
                _MySqlDBName = value
            End Set
        End Property
        ''' <summary>
        ''' MySql Port
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MySqlPort() As String
            Get
                Return _MySqlport
            End Get
            Set(ByVal value As String)
                _MySqlport = value
            End Set
        End Property

#End Region

#Region "MySQL Connection"
        Public Function MySqlConn(Optional ByVal ConnectionString As String = "", Optional ByVal MySqlport As Integer = 3306)

            Try

                If ConnectionString.Length > 0 Then
                    MySqlDBConn.ConnectionString = ConnectionString
                Else
                    'Dim OdbcDBConn As New OdbcConnection
                    MySqlDBConn.ConnectionString = "server=" & _MySqlhostname & ";" &
                                                "database=" & _MySqlDBName & ";" &
                                                "uid=" & _MySqlusername & ";" &
                                                "pwd=" & _MySqlpassword & ";" &
                                                "port=" & _MySqlport & ";"
                End If

                'oOdbcConn.ConnectionString = _
                '"User ID=uid;" & _
                '"Password=pw;" & _
                '"Host=ip;" & _
                '"Port=3306;" & _
                '"Database=db;" & _
                '"Direct=true;" & _
                '"Protocol=TCP;" & _
                '"Compress=false;" & _
                '"Pooling=true;" & _
                '"Min Pool Size=0;" & _
                '"Max Pool Size=100;" & _
                '"Connection Lifetime=0"


                'OdbcConn.ConnectionString = "Host=" & DBServer & ";" & _
                '                            "Database=" & DBName & ";" & _
                '                            "User ID=" & DBLogin & ";" & _
                '                            "Password=" & DBPassword & ";" & _
                '                            "Port=" & DBPort & ";" & _
                '                            "Direct=True;" & _
                '                            "Protocol=TCP;" & _
                '                            "Compress=false;" & _
                '                            "Min Pool Size=0;" & _
                '                            "Max Pool Size=100;" & _
                '                            "Connection Lifetime=0"

                _log.Info(Environment.UserName & " Opened MySql Database : " & _MySqlhostname & " DBName " & _MySqlDBName)
                MySqlDBConn.Open()

                Return True

            Catch ex As MySqlException
                If ex.Number = 0 Then
                    Throw New ApplicationException("Connection to " & _MySqlhostname & " was not Successfull.")
                    _log.Error("Connection to " & _MySqlhostname & " was not Successfull.")
                ElseIf ex.Number = 1045 Then
                    Throw New ApplicationException("Invalid username/password, please try again")
                    _log.Error("Connection to " & _MySqlhostname & " was not Successfull as Invalid username/password, please try again")
                End If
                Throw New ApplicationException("Connection to " & _MySqlhostname & " was not Successfull.")
                _log.Error("Connection to " & _MySqlhostname & " was not Successfull with error number: " & ex.Number.ToString)
            End Try
        End Function
#End Region

#Region "ODBC Properties"
        ''' <summary>
        ''' ODBC Hostname
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ODBCHostname() As String
            Get
                Return _Odbchostname
            End Get
            Set(ByVal value As String)
                _Odbchostname = value
            End Set
        End Property

        ''' <summary>
        ''' ODBC Username
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property ODBCUsername() As String
            Get
                Return _Odbcusername
            End Get
            Set(ByVal value As String)
                _Odbcusername = value
            End Set
        End Property

        ''' <summary>
        ''' ODBC Password
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ODBCPassword() As String
            Get
                Return _Odbcpassword
            End Get
            Set(ByVal value As String)
                _Odbcpassword = value
            End Set
        End Property

        ''' <summary>
        ''' ODBC DBName
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ODBCDBName() As String
            Get
                Return _OdbcDBName
            End Get
            Set(ByVal value As String)
                _OdbcDBName = value
            End Set
        End Property

        ''' <summary>
        ''' ODBC Port
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ODBCPort() As String
            Get
                Return _Odbcport
            End Get
            Set(ByVal value As String)
                _Odbcport = value
            End Set
        End Property

        ''' <summary>
        ''' ODBC DSN
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>DSN Name</remarks>
        Public Property ODBCDSN() As String
            Get
                Return _OdbcDSN
            End Get
            Set(ByVal value As String)
                _OdbcDSN = value
            End Set
        End Property
#End Region

#Region "ODBC Connection"
        ''' <summary>
        ''' ODBC Connection
        ''' </summary>
        ''' <param name="DBType">Database Type MySQL, TranSoft or DSN</param>
        ''' <returns></returns>
        ''' <remarks>Opens a connection via ODBC</remarks>
        Public Function ODBCConn(ByVal DBType As String) As OdbcConnection

            Try
                'UpperCase for user input
                DBType = DBType.ToUpper()

                If DBType = "MYSQL" Then
                    ODBCDBConn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
                                                   "Server=" & _Odbchostname & ";" & _
                                                   "Port=" & _Odbcport & ";" & _
                                                   "Option=3;" & _
                                                   "Stmt=;" & _
                                                   "Database=" & _OdbcDBName & ";" & _
                                                   "Uid=" & _Odbcusername & ";" & _
                                                   "Pwd=" & _Odbcpassword & ""
                    ODBCDBConn.Open()
                End If

                If DBType = "TRANSOFT" Then
                    ODBCDBConn.ConnectionString = "DRIVER={Transoft ODBC Driver};" & _
                                                    "TSDSN=" & _OdbcDSN & ";" & _
                                                    "Server=" & _Odbchostname & ";" & _
                                                "Port=" & _Odbcport & ";" & _
                                                    "Timeout=200;" & _
                                                    "Description="""
                    ODBCDBConn.Open()
                End If

                If DBType = "DSN" Then
                    ODBCDBConn.ConnectionString = "DSN=" & _OdbcDSN & ";Uid=" & _Odbcusername & ";" & _
                    "Pwd=" & _Odbcpassword & ""
                    ODBCDBConn.Open()
                End If

                Return ODBCDBConn

            Catch ex As Exception
                Throw New ArgumentException()
            End Try

            'Dim csLogon As String
            'cLogon = "DRIVER={MySQL ODBC 3.51 Driver};" & "Server=" & DBServer & ";Port=3306;Database=xxxxxx;Uid=xxxxxx;Pwd=xxxxxx;Option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            'adoConn = New System.Data.Odbc.OdbcConnection(csLogon)

        End Function
#End Region

#Region "Public Functions"
        ''' <summary>
        ''' Gets the MsSql Connection State Opened or Closed
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MsSqlState() As Boolean
            If MsSqlDBConn.State = ConnectionState.Closed Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Gets the MySql Connection State Opened or Closed
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MySqlState() As Boolean
            If MySqlDBConn.State = ConnectionState.Closed Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Gets the Odbc Connection State Opened or Closed
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function OdbcState() As Boolean
            If ODBCDBConn.State = ConnectionState.Closed Then
                Return False
            Else
                Return True
            End If
        End Function
#End Region

    End Class
#End Region

#Region " SQLCommands "
    Public Class clsSQLExecutes
        Inherits Onling.clsDatabase
        ''' <summary>
        ''' A Sql Data Provider Class
        ''' </summary>
        ''' <remarks>This class is used to ease the commands to update the database</remarks>
        Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsSQLExecutes))

#Region "LOCAL DECLARATIONS"

#End Region

#Region "CONSRUCTORS"
        ''' <summary>
        ''' Blank constructor
        ''' </summary>
        ''' <remarks></remarks>
        Sub New()
            _log.Info("Starting " & MethodBase.GetCurrentMethod().ToString())
        End Sub
#End Region

#Region "PROPERTIES"

#End Region

#Region "PRIVATE METHODS"
        Private Sub AssignParameters(ByVal cmd As SqlCommand, ByVal cmdParameters() As SqlParameter)
            If (cmdParameters Is Nothing) Then Exit Sub
            For Each p As SqlParameter In cmdParameters
                cmd.Parameters.Add(p)
                _log.Info("Adding Command Parameter " & p.Value.ToString)
            Next
        End Sub

        Private Sub AssignParameters(ByVal cmd As OdbcCommand, ByVal cmdParameters() As OdbcParameter)
            If (cmdParameters Is Nothing) Then Exit Sub
            For Each p As OdbcParameter In cmdParameters
                cmd.Parameters.Add(p)
                _log.Info("Adding Command Parameter " & p.Value.ToString)
            Next
        End Sub

        Private Sub AssignParameters(ByVal cmd As SqlCommand, ByVal parameterValues() As Object)
            If Not (cmd.Parameters.Count - 1 = parameterValues.Length) Then Throw New ApplicationException("Stored procedure's parameters and parameter values does not match.")
            Dim i As Integer
            For Each param As SqlParameter In cmd.Parameters
                If Not (param.Direction = ParameterDirection.Output) AndAlso Not (param.Direction = ParameterDirection.ReturnValue) Then
                    param.Value = parameterValues(i)
                    _log.Info("Adding Parameter " & param.Value.ToString)
                    i += 1
                End If
            Next
        End Sub
        Private Sub AssignParameters(ByVal cmd As OdbcCommand, ByVal parameterValues() As Object)
            If Not (cmd.Parameters.Count - 1 = parameterValues.Length) Then Throw New ApplicationException("Stored procedure's parameters and parameter values does not match.")
            Dim i As Integer
            For Each param As OdbcParameter In cmd.Parameters
                If Not (param.Direction = ParameterDirection.Output) AndAlso Not (param.Direction = ParameterDirection.ReturnValue) Then
                    param.Value = parameterValues(i)
                    _log.Info("Adding Parameter " & param.Value.ToString)
                    i += 1
                End If
            Next
        End Sub
#End Region

#Region " Execute Stored Procedures "
        ''' <summary>
        ''' To Execute a SQL based Stored Prcedure that cannot be sent via a Transaction.
        ''' </summary>
        ''' <param name="spname">The name of the stored procedure to execute at the data source.</param>
        ''' <param name="parameterValues">The parameter values of the stored procedure.</param>
        ''' <returns>0 for Success or 1 for Failure</returns>
        ''' <remarks></remarks>
        Public Function ExecuteSPSQL(ByVal spname As String, ByVal ParamArray parameterValues() As Object) As Integer
            Dim command As SqlCommand = Nothing
            Dim res As Integer = -1
            Try
                command = New SqlCommand(spname, MsSqlDBConn)
                command.CommandType = CommandType.StoredProcedure
                command.CommandTimeout = 0
                SqlCommandBuilder.DeriveParameters(command)
                _log.Info("Command Text : " & command.CommandText)
                Me.AssignParameters(command, parameterValues)
                _log.Info("Executing : " & command.CommandText)
                res = command.ExecuteNonQuery()

            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)

            Finally
                If Not (command Is Nothing) Then command.Dispose()
            End Try
            Return res
        End Function
#End Region

#Region " ExecuteNonQuery "
        ''' <summary>
        ''' Executes a Transact-SQL statement against the connection and returns the number of rows affected.
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.SqlClient.SqlCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>The number of rows affected.</returns>
        Public Function ExecuteNonQuerySQL_NoTan(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As SqlParameter = Nothing) As Integer
            Dim transaction As SqlTransaction = Nothing
            Dim command As SqlCommand = Nothing
            Dim res As Integer = -1
            Try
                command = New SqlCommand(cmd, MsSqlDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                _log.Info("Command Text : " & command.CommandText)
                Me.AssignParameters(command, parameters)
                'transaction = MsSqlDBConn.BeginTransaction
                'command.Transaction = transaction
                _log.Info("Executing : " & command.CommandText)
                res = command.ExecuteNonQuery()
            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                    Throw New SqlDatabaseException("Transaction Rolled Back " & ex.Message, ex.InnerException)
                    _log.Error(ex.ToString)
                    _log.Error(ex.InnerException.ToString)
                End If
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
                'Throw New SqlDatabaseException(ex.Message, ex.InnerException)
            Finally
                'transaction.Commit()
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

        ''' <summary>
        ''' Executes a Transact-SQL statement against the connection and returns the number of rows affected.
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.SqlClient.SqlCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>The number of rows affected.</returns>
        Public Function ExecuteNonQuerySQL(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As SqlParameter = Nothing) As Integer
            Dim transaction As SqlTransaction = Nothing
            Dim command As SqlCommand = Nothing
            Dim res As Integer = -1
            Try
                command = New SqlCommand(cmd, MsSqlDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                _log.Info("Command Text : " & command.CommandText)
                Me.AssignParameters(command, parameters)
                transaction = MsSqlDBConn.BeginTransaction
                command.Transaction = transaction
                _log.Info("Executing : " & command.CommandText)
                res = command.ExecuteNonQuery()
            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                    Throw New SqlDatabaseException("Transaction Rolled Back " & ex.Message, ex.InnerException)
                    _log.Error(ex.ToString)
                    _log.Error(ex.InnerException.ToString)
                End If
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                transaction.Commit()
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

        ''' <summary>
        ''' Executes a Transact-SQL statement against the connection and returns the number of rows affected.
        ''' </summary>
        ''' <param name="spname">The stored procedure to execute at the data source.</param>
        ''' <param name="returnValue">The returned value from stored procedure.</param>
        ''' <param name="parameterValues">The parameter values of the stored procedure.</param>
        ''' <returns>The number of rows affected.</returns>
        Public Function ExecuteNonQuerySQL(ByVal spname As String, ByRef returnValue As Integer, ByVal ParamArray parameterValues() As Object) As Integer
            Dim transaction As SqlTransaction = Nothing
            Dim command As SqlCommand = Nothing
            Dim res As Integer = -1
            Try
                command = New SqlCommand(spname, MsSqlDBConn)
                command.CommandType = CommandType.StoredProcedure
                command.CommandTimeout = 0
                SqlCommandBuilder.DeriveParameters(command)
                _log.Info("Command Text : " & command.CommandText)
                Me.AssignParameters(command, parameterValues)
                transaction = MsSqlDBConn.BeginTransaction()
                command.Transaction = transaction
                _log.Info("Executing : " & command.CommandText)
                res = command.ExecuteNonQuery()
                returnValue = command.Parameters(0).Value

            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                    Throw New SqlDatabaseException("Transaction Rolled Back " & ex.Message, ex.InnerException)
                    _log.Error(ex.ToString)
                    _log.Error(ex.InnerException.ToString)
                End If
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)

            Finally
                transaction.Commit()
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

        ''' <summary>
        ''' Executes a Transact-SQL statement against the connection and returns the number of rows affected.
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.OdbcClient.OdbcCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>The number of rows affected.</returns>
        Public Function ExecuteNonQueryODBC(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As OdbcParameter = Nothing) As Integer
            Dim transaction As OdbcTransaction = Nothing
            Dim command As OdbcCommand = Nothing
            Dim res As Integer = -1
            Try
                command = New OdbcCommand(cmd, ODBCDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                _log.Info("Command Text : " & command.CommandText)
                Me.AssignParameters(command, parameters)
                transaction = ODBCDBConn.BeginTransaction
                command.Transaction = transaction
                _log.Info("Executing : " & command.CommandText)
                res = command.ExecuteNonQuery()
            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                    Throw New ApplicationException("Transaction Rolled Back " & ex.Message, ex.InnerException)
                    _log.Error(ex.ToString)
                    _log.Error(ex.InnerException.ToString)
                End If
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                transaction.Commit()
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

        ''' <summary>
        ''' Executes a Transact-SQL statement against the connection and returns the number of rows affected.
        ''' </summary>
        ''' <param name="spname">The stored procedure to execute at the data source.</param>
        ''' <param name="returnValue">The returned value from stored procedure.</param>
        ''' <param name="parameterValues">The parameter values of the stored procedure.</param>
        ''' <returns>The number of rows affected.</returns>
        Public Function ExecuteNonQueryODBC(ByVal spname As String, ByRef returnValue As Integer, ByVal ParamArray parameterValues() As Object) As Integer
            Dim transaction As OdbcTransaction = Nothing
            Dim command As OdbcCommand = Nothing
            Dim res As Integer = -1
            Try
                command = New OdbcCommand(spname, ODBCDBConn)
                command.CommandType = CommandType.StoredProcedure
                command.CommandTimeout = 0
                _log.Info("Command Text : " & command.CommandText)
                OdbcCommandBuilder.DeriveParameters(command)
                Me.AssignParameters(command, parameterValues)
                transaction = ODBCDBConn.BeginTransaction()
                command.Transaction = transaction
                _log.Info("Executing : " & command.CommandText)
                res = command.ExecuteNonQuery()
                returnValue = command.Parameters(0).Value
            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                End If
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                Throw New SqlDatabaseException("Transaction Rolled Back " & ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                transaction.Commit()
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

#End Region

#Region " ExecuteScalar "

        ''' <summary>
        ''' Executes the query, and returns the first column of the first row in the result set returned by the query. Additional columns or rows are ignored.
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.SqlClient.SqlCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns>
        Public Function ExecuteScalarSQL(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As SqlParameter = Nothing) As Object
            Dim transaction As SqlTransaction = Nothing
            Dim command As SqlCommand = Nothing
            Dim res As Object = Nothing
            Try
                command = New SqlCommand(cmd, MsSqlDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                Me.AssignParameters(command, parameters)
                transaction = MsSqlDBConn.BeginTransaction()
                command.Transaction = transaction
                res = command.ExecuteScalar()
                transaction.Commit()
            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                End If
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

        ''' <summary>
        ''' Executes the query, and returns the first column of the first row in the result set returned by the query. Additional columns or rows are ignored.
        ''' </summary>
        ''' <param name="spname">The stored procedure to execute at the data source.</param>
        ''' <param name="returnValue">The returned value from stored procedure.</param>
        ''' <param name="parameterValues">The parameter values of the stored procedure.</param>
        ''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns>
        Public Function ExecuteScalarSQL(ByVal spname As String, ByRef returnValue As Integer, ByVal ParamArray parameterValues() As Object) As Object
            Dim transaction As SqlTransaction = Nothing
            Dim command As SqlCommand = Nothing
            Dim res As Object = Nothing
            Try
                command = New SqlCommand(spname, MsSqlDBConn)
                command.CommandType = CommandType.StoredProcedure
                command.CommandTimeout = 0
                SqlCommandBuilder.DeriveParameters(command)
                Me.AssignParameters(command, parameterValues)
                transaction = MsSqlDBConn.BeginTransaction()
                command.Transaction = transaction
                res = command.ExecuteScalar()
                returnValue = command.Parameters(0).Value
                transaction.Commit()
            Catch ex As Exception
                If Not (transaction Is Nothing) Then
                    transaction.Rollback()
                End If
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If Not (command Is Nothing) Then command.Dispose()
                If Not (transaction Is Nothing) Then transaction.Dispose()
            End Try
            Return res
        End Function

#End Region

#Region " ExecuteReader "

        ''' <summary>
        ''' Sends the System.Data.SqlClient.SqlCommand.CommandText to the System.Data.SqlClient.SqlCommand.Connection, and builds a System.Data.SqlClient.SqlDataReader using one of the System.Data.CommandBehavior values.
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.SqlClient.SqlCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>A System.Data.SqlClient.SqlDataReader object.</returns>
        Public Function ExecuteReader(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As SqlParameter = Nothing) As IDataReader
            Dim command As SqlCommand = Nothing
            Dim res As SqlDataReader = Nothing
            Try
                command = New SqlCommand(cmd, MsSqlDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                Me.AssignParameters(command, parameters)
                res = command.ExecuteReader(CommandBehavior.CloseConnection)
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            End Try
            Return CType(res, IDataReader)
        End Function

        ''' <summary>
        ''' Sends the System.Data.SqlClient.SqlCommand.CommandText to the System.Data.SqlClient.SqlCommand.Connection, and builds a System.Data.SqlClient.SqlDataReader using one of the System.Data.CommandBehavior values.
        ''' </summary>
        ''' <param name="spname">The stored procedure to execute at the data source.</param>
        ''' <param name="returnValue">The returned value from stored procedure.</param>
        ''' <param name="parameterValues">The parameter values of the stored procedure.</param>
        ''' <returns>A System.Data.SqlClient.SqlDataReader object.</returns>
        Public Function ExecuteReader(ByVal spname As String, ByRef returnValue As Integer, ByVal ParamArray parameterValues() As Object) As IDataReader
            Dim connection As SqlConnection = Nothing
            Dim command As SqlCommand = Nothing
            Dim res As SqlDataReader = Nothing
            Try
                command = New SqlCommand(spname, connection)
                command.CommandType = CommandType.StoredProcedure
                command.CommandTimeout = 0
                connection.Open()
                SqlCommandBuilder.DeriveParameters(command)
                Me.AssignParameters(command, parameterValues)
                res = command.ExecuteReader(CommandBehavior.CloseConnection)
                returnValue = command.Parameters(0).Value
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            End Try
            Return CType(res, IDataReader)
        End Function

#End Region

#Region " FillDataset "

        ''' <summary>
        ''' Adds or refreshes rows in the System.Data.DataSet to match those in the data source using the System.Data.DataSet name, and creates a System.Data.DataTable named "Table."
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.SqlClient.SqlCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>A System.Data.Dataset object.</returns>
        Public Function FillDatasetSQL(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As SqlParameter = Nothing) As DataSet
            Dim command As SqlCommand = Nothing
            Dim sqlda As SqlDataAdapter = Nothing
            Dim res As New DataSet
            Try
                command = New SqlCommand(cmd, MsSqlDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                AssignParameters(command, parameters)
                sqlda = New SqlDataAdapter(command)
                sqlda.Fill(res)
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If Not (command Is Nothing) Then command.Dispose()
                If Not (sqlda Is Nothing) Then sqlda.Dispose()
            End Try
            Return res
        End Function

        ''' <summary>
        ''' Adds or refreshes rows in the System.Data.DataSet to match those in the data source using the System.Data.DataSet name, and creates a System.Data.DataTable named "Table."
        ''' </summary>
        ''' <param name="cmd">The Transact-SQL statement or stored procedure to execute at the data source.</param>
        ''' <param name="cmdType">A value indicating how the System.Data.SqlClient.SqlCommand.CommandText property is to be interpreted.</param>
        ''' <param name="parameters">The parameters of the Transact-SQL statement or stored procedure.</param>
        ''' <returns>A System.Data.Dataset object.</returns>
        Public Function FillDatasetODBC(ByVal cmd As String, ByVal cmdType As CommandType, Optional ByVal parameters() As OdbcParameter = Nothing) As DataSet
            Dim command As OdbcCommand = Nothing
            Dim sqlda As OdbcDataAdapter = Nothing
            Dim res As New DataSet
            Try
                command = New OdbcCommand(cmd, ODBCDBConn)
                command.CommandType = cmdType
                command.CommandTimeout = 0
                AssignParameters(command, parameters)
                sqlda = New OdbcDataAdapter(command)
                sqlda.Fill(res)
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If Not (command Is Nothing) Then command.Dispose()
                If Not (sqlda Is Nothing) Then sqlda.Dispose()
            End Try
            Return res
        End Function
#End Region

#Region " ExecuteDataset "

        ''' <summary>
        ''' Calls the respective INSERT, UPDATE, or DELETE statements for each inserted, updated, or deleted row in the System.Data.DataSet with the specified System.Data.DataTable name.
        ''' </summary>
        ''' <param name="insertCmd">A command used to insert new records into the data source.</param>
        ''' <param name="updateCmd">A command used to update records in the data source.</param>
        ''' <param name="deleteCmd">A command for deleting records from the data set.</param>
        ''' <param name="ds">The System.Data.DataSet to use to update the data source. </param>
        ''' <param name="srcTable">The name of the source table to use for table mapping.</param>
        ''' <returns>The number of rows successfully updated from the System.Data.DataSet.</returns>
        Public Function ExecuteDataset(ByVal insertCmd As SqlCommand, ByVal updateCmd As SqlCommand, ByVal deleteCmd As SqlCommand, ByVal ds As DataSet, ByVal srcTable As String) As Integer
            Dim sqlda As SqlDataAdapter = Nothing
            Dim res As Integer = 0
            Try
                sqlda = New SqlDataAdapter
                If Not (insertCmd Is Nothing) Then insertCmd.Connection = MsSqlDBConn : sqlda.InsertCommand = insertCmd
                If Not (updateCmd Is Nothing) Then updateCmd.Connection = MsSqlDBConn : sqlda.UpdateCommand = updateCmd
                If Not (deleteCmd Is Nothing) Then deleteCmd.Connection = MsSqlDBConn : sqlda.DeleteCommand = deleteCmd
                res = sqlda.Update(ds, srcTable)
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If Not (insertCmd Is Nothing) Then insertCmd.Dispose()
                If Not (updateCmd Is Nothing) Then updateCmd.Dispose()
                If Not (deleteCmd Is Nothing) Then deleteCmd.Dispose()
                If Not (sqlda Is Nothing) Then sqlda.Dispose()
            End Try
            Return res
        End Function

#End Region

#Region " ExecuteScript "

        ''' <summary>
        ''' Executes a SQL query file against the connection.
        ''' </summary>
        ''' <param name="filename">SQL query file name.</param>
        ''' <param name="parameters">The parameters of the SQL query file.</param>
        Public Sub ExecuteScriptSQL(ByVal filename As String, Optional ByVal parameters() As SqlParameter = Nothing)
            Dim fStream As FileStream = Nothing
            Dim sReader As StreamReader = Nothing
            Dim command As SqlCommand = Nothing
            Try
                fStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                sReader = New StreamReader(fStream)
                command = MsSqlDBConn.CreateCommand()
                While (Not sReader.EndOfStream)
                    Dim sb As New StringBuilder
                    While (Not sReader.EndOfStream)
                        Dim s As String = sReader.ReadLine
                        If (Not String.IsNullOrEmpty(s)) AndAlso (s.ToUpper.Trim = "GO") Then
                            Exit While
                        End If
                        sb.AppendLine(s)
                    End While
                    command.CommandText = sb.ToString
                    command.CommandType = CommandType.Text
                    command.CommandTimeout = 0
                    AssignParameters(command, parameters)
                    command.ExecuteNonQuery()
                End While
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If (Not IsNothing(command)) Then command.Dispose()
                If (Not IsNothing(sReader)) Then sReader.Close()
                If (Not IsNothing(fStream)) Then fStream.Close()
            End Try
        End Sub

        ''' <summary>
        ''' Executes a SQL query file against the connection.
        ''' </summary>
        ''' <param name="filename">SQL query file name.</param>
        ''' <param name="parameters">The parameters of the SQL query file.</param>
        Public Function ExecuteScriptODBC(ByVal filename As String, Optional ByVal parameters() As OdbcParameter = Nothing)
            Dim fStream As FileStream = Nothing
            Dim sReader As StreamReader = Nothing
            Dim command As OdbcCommand = Nothing
            Try
                fStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                sReader = New StreamReader(fStream)
                command = ODBCDBConn.CreateCommand()
                While (Not sReader.EndOfStream)
                    Dim sb As New StringBuilder
                    While (Not sReader.EndOfStream)
                        Dim s As String = sReader.ReadLine
                        If (Not String.IsNullOrEmpty(s)) AndAlso (s.ToUpper.Trim = "GO") Then
                            Exit While
                        End If
                        sb.AppendLine(s)
                    End While
                    command.CommandText = sb.ToString
                    command.CommandType = CommandType.Text
                    command.CommandTimeout = 0
                    AssignParameters(command, parameters)
                    command.ExecuteNonQuery()
                End While
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If (Not IsNothing(command)) Then command.Dispose()
                If (Not IsNothing(sReader)) Then sReader.Close()
                If (Not IsNothing(fStream)) Then fStream.Close()
            End Try
        End Function

#End Region

#Region " Read Script"
        ''' <summary>
        ''' Reads a SQL query file
        ''' </summary>
        ''' <param name="filename">SQL query file name.</param>
        ''' <param name="parameters">The parameters of the SQL query file.</param>
        Public Function ReadsScriptSQL(ByVal filename As String, Optional ByVal parameters() As SqlParameter = Nothing) As String
            Dim fStream As FileStream = Nothing
            Dim sReader As StreamReader = Nothing
            Dim command As SqlCommand = Nothing
            Try
                fStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                sReader = New StreamReader(fStream)
                command = MsSqlDBConn.CreateCommand()
                While (Not sReader.EndOfStream)
                    Dim sb As New StringBuilder
                    While (Not sReader.EndOfStream)
                        Dim s As String = sReader.ReadLine
                        If (Not String.IsNullOrEmpty(s)) AndAlso (s.ToUpper.Trim = "GO") Then
                            Exit While
                        End If
                        sb.AppendLine(s)
                    End While
                    command.CommandText = sb.ToString
                    command.CommandType = CommandType.Text
                    command.CommandTimeout = 0
                    AssignParameters(command, parameters)
                    'command.ExecuteNonQuery()
                    Return command.CommandText
                End While
            Catch ex As Exception
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
                _log.Error(ex.ToString)
                _log.Error(ex.InnerException.ToString)
            Finally
                If (Not IsNothing(command)) Then command.Dispose()
                If (Not IsNothing(sReader)) Then sReader.Close()
                If (Not IsNothing(fStream)) Then fStream.Close()
            End Try
        End Function

        ''' <summary>
        ''' Reads a SQL query from a file
        ''' </summary>
        ''' <param name="filename">SQL query file name.</param>
        ''' <param name="parameters">The parameters of the SQL query file.</param>
        Public Function ReadScriptODBC(ByVal filename As String, Optional ByVal parameters() As OdbcParameter = Nothing) As String
            Dim fStream As FileStream = Nothing
            Dim sReader As StreamReader = Nothing
            Dim command As OdbcCommand = Nothing
            Try
                fStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                sReader = New StreamReader(fStream)
                command = ODBCDBConn.CreateCommand()
                While (Not sReader.EndOfStream)
                    Dim sb As New StringBuilder
                    While (Not sReader.EndOfStream)
                        Dim s As String = sReader.ReadLine
                        If (Not String.IsNullOrEmpty(s)) AndAlso (s.ToUpper.Trim = "GO") Then
                            Exit While
                        End If
                        sb.AppendLine(s)
                    End While
                    command.CommandText = sb.ToString
                    command.CommandType = CommandType.Text
                    command.CommandTimeout = 0
                    AssignParameters(command, parameters)
                    'command.ExecuteNonQuery()
                    Return command.CommandText
                End While
            Catch ex As Exception
                _log.Error(ex.Message)
                _log.Error(ex.InnerException)
                Throw New SqlDatabaseException(ex.Message, ex.InnerException)
            Finally
                If (Not IsNothing(command)) Then command.Dispose()
                If (Not IsNothing(sReader)) Then sReader.Close()
                If (Not IsNothing(fStream)) Then fStream.Close()
            End Try
        End Function

#End Region

    End Class
#End Region

#Region " SQL Database Exception "
    Public Class SqlDatabaseException
        Inherits Exception

#Region " CONSTRUCTORS "
        ''' <summary>
        ''' Initializes a new instance of the ADO.SqlDatabaseException class.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

        End Sub

        ''' <summary>
        ''' Initializes a new instance of the ADO.SqlDatabaseException class with a specified error message.
        ''' </summary>
        ''' <param name="message">The message that describes the error.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the ADO.SqlDatabaseException class with a specified error message 
        ''' and a reference to the inner exception that is the cause of this exception.
        ''' </summary>
        ''' <param name="message">The error message that explains the reason for the exception.</param>
        ''' <param name="innerException">The exception that is the cause of the current exception, or a null reference 
        ''' (Nothing in Visual Basic) if no inner exception is specified.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
        End Sub
#End Region

    End Class
#End Region

End Namespace


