Imports System.Configuration
Imports System.Xml
''' <summary>
''' AppConfigFileSettings: This class is used to Change the 
''' AppConfigs Parameters at runtime through User Interface
''' </summary>
''' <remarks></remarks>
Public Class clsAppConfigFileSettings
    ''' <summary>
    ''' UpdateAppSettings: It will update the app.Config file AppConfig key values
    ''' </summary>
    ''' <param name="KeyName">AppConfigs KeyName</param>
    ''' <param name="KeyValue">AppConfigs KeyValue</param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateAppSettings(ByVal KeyName As String, ByVal KeyValue As String)
        '  AppDomain.CurrentDomain.SetupInformation.ConfigurationFile 
        ' This will get the app.config file path from Current application Domain
        Dim XmlDoc As New XmlDocument()
        ' Load XML Document
        XmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)
        ' Navigate Each XML Element of app.Config file
        For Each xElement As XmlElement In XmlDoc.DocumentElement
            If xElement.Name = "appSettings" Then
                ' Loop each node of appSettings Element 
                ' xNode.Attributes(0).Value , Mean First Attributes of Node , 
                ' KeyName Portion
                ' xNode.Attributes(1).Value , Mean Second Attributes of Node,
                ' KeyValue Portion
                For Each xNode As XmlNode In xElement.ChildNodes
                    If xNode.Attributes(0).Value = KeyName Then
                        xNode.Attributes(1).Value = KeyValue
                    End If
                Next
            End If
        Next
        ' Save app.config file
        XmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)
    End Sub

    ' Usage This Will set the appConfigs Paramters values to Text box Controls
    'Private Sub LoadConfigValueToControls()
    '    txtServerName.Text = System.Configuration.ConfigurationSettings.AppSettings.Get("DBServerName")
    '    txtDBName.Text = System.Configuration.ConfigurationSettings.AppSettings.Get("DatabaseName")
    '    txtDBUserID.Text = System.Configuration.ConfigurationSettings.AppSettings.Get("DatabaseUserID")
    '    txtDBPwd.Text = System.Configuration.ConfigurationSettings.AppSettings.Get("DatabasePwd")
    'End Sub

    'Private Sub btnChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChange.Click
    '    AppConfigFileSettings.UpdateAppSettings("DBServerName", txtServerName.Text)
    '    AppConfigFileSettings.UpdateAppSettings("DatabaseName", txtDBName.Text)
    '    AppConfigFileSettings.UpdateAppSettings("DatabaseUserID", txtDBUserID.Text)
    '    AppConfigFileSettings.UpdateAppSettings("DatabasePwd", txtDBPwd.Text)
    '    MsgBox("Application Settings has been Changed successfully.")
    'End Sub
End Class
