Imports System
Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class clsCommonAppFolder
    Private applicationFolder As String
    Private companyFolder As String
    Private Shared ReadOnly directory As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)

    Public Sub New(ByVal companyFolder As String, ByVal applicationFolder As String)
        Me.New(companyFolder, applicationFolder, False)
    End Sub

    Public Sub New(ByVal companyFolder As String, ByVal applicationFolder As String, ByVal allUsers As Boolean)
        Me.applicationFolder = applicationFolder
        Me.companyFolder = companyFolder
        CreateFolders(allUsers)
    End Sub

    Public ReadOnly Property ApplicationFolderPath As String
        Get
            Return Path.Combine(CompanyFolderPath, applicationFolder)
        End Get
    End Property

    Public ReadOnly Property CompanyFolderPath As String
        Get
            Return Path.Combine(directory, companyFolder)
        End Get
    End Property

    Private Sub CreateFolders(ByVal allUsers As Boolean)
        Dim directoryInfo As DirectoryInfo
        Dim directorySecurity As DirectorySecurity
        Dim rule As AccessRule
        Dim securityIdentifier As SecurityIdentifier = New SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, Nothing)

        If Not IO.Directory.Exists(CompanyFolderPath) Then
            directoryInfo = IO.Directory.CreateDirectory(CompanyFolderPath)
            Dim modified As Boolean
            directorySecurity = directoryInfo.GetAccessControl()
            rule = New FileSystemAccessRule(securityIdentifier, FileSystemRights.Write Or FileSystemRights.ReadAndExecute Or FileSystemRights.Modify, AccessControlType.Allow)
            directorySecurity.ModifyAccessRule(AccessControlModification.Add, rule, modified)
            directoryInfo.SetAccessControl(directorySecurity)
        End If

        If Not IO.Directory.Exists(ApplicationFolderPath) Then
            directoryInfo = IO.Directory.CreateDirectory(ApplicationFolderPath)

            If allUsers Then
                Dim modified As Boolean
                directorySecurity = directoryInfo.GetAccessControl()
                rule = New FileSystemAccessRule(securityIdentifier, FileSystemRights.Write Or FileSystemRights.ReadAndExecute Or FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.InheritOnly, AccessControlType.Allow)
                directorySecurity.ModifyAccessRule(AccessControlModification.Add, rule, modified)
                directoryInfo.SetAccessControl(directorySecurity)
            End If
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return ApplicationFolderPath
    End Function
End Class

