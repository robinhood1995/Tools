Imports System.IO
Imports System.Text
Imports log4net
Imports log4net.Config
Imports System.Reflection

Public Class clsRSRGroup

    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsRSRGroup))

#Region "Local Decarations"

    Private _Path As String
    Private _Filename As String
    Private _sb As New StringBuilder

#End Region

#Region "Constructors"
    ''' <summary>
    ''' Manual Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _log.Info("Starting " & MethodBase.GetCurrentMethod().ToString())
    End Sub
#End Region

#Region "Private Propeeties"
    Private Property sb() As StringBuilder
        Get
            Return _sb
        End Get
        Set(ByVal value As StringBuilder)
            _sb = value
        End Set
    End Property
#End Region

End Class
