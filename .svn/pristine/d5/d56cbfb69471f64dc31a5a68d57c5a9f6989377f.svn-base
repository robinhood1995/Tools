Imports log4net
Imports log4net.Config
Imports System.Windows.Forms

''' <summary>
'''  This Class will parse the SQL text to replace parameter labels with obj code 
''' </summary>
''' <remarks></remarks>

Public Class clsParseSQLtxtParam

    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsParseSQLtxtParam))

    Public Function replaceParameters(ByVal sqlTXT As String _
                                    , ByVal objDRconn As DataRow _
                                    , Optional ByVal storenumber As String = "") As String
        replaceParameters = ""

        Try
            replaceParameters = sqlTXT
            replaceParameters.Replace("@PlantCode@", objDRconn.Item("linkmysqlplantinput"))
            replaceParameters.Replace("@StoreNumber@", storenumber)

        Catch ex As Exception
            _log.Error(ex.ToString & vbCrLf & ex.StackTrace.ToString)
            MessageBox.Show(ex.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Function

End Class
