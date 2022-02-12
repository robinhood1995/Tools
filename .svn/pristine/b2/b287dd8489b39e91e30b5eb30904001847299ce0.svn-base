Imports log4net
Imports log4net.Config
Imports System.Configuration
Imports System.Reflection

Public Class clsFileInfo

    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsFileInfo))

    Public FileInfo As System.IO.FileInfo = Nothing

    Public Sub New(ByVal Path As String)
        Me.FileInfo = New System.IO.FileInfo(Path)
    End Sub

    Public Overrides Function ToString() As String
        Return Me.FileInfo.Name
    End Function

End Class
