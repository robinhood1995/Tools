Option Compare Binary
Option Explicit On
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

Public Delegate Sub OnLineAddedInvoker(ByVal line As String)

Public Class clsOutput

    Public Event OnLineAdded As OnLineAddedInvoker

    Private m_Process As Process
    Private m_OutputThread As Thread
    Private m_ErrorThread As Thread


    Public Sub Start()

        ' To use the program to execute a command on the
        ' You can run the console (e.g. "netstat" or "ping")
        ' instead of "cmd" for arguments, simply use the corresponding one
        ' Insert command (e.g. "ping") and in the property 'Arguments'
        ' the corresponding arguments (for example, for "ping" the IP address
        ' address).

        m_Process = New Process
        With m_Process.StartInfo
            .FileName = "cmd"
            .UseShellExecute = False
            .CreateNoWindow = True
            .RedirectStandardOutput = True
            .RedirectStandardError = True
            .RedirectStandardInput = True
        End With
        m_Process.Start()

        ' Changing the data streams so that we are aware of changes.
        m_OutputThread = New Thread(AddressOf StreamOutput)
        m_OutputThread.IsBackground = True
        m_OutputThread.Start()
        m_ErrorThread = New Thread(AddressOf StreamError)
        m_ErrorThread.IsBackground = True
        m_ErrorThread.Start()
    End Sub

    Public Sub Send(ByVal text As String)
        StreamInput(text)
    End Sub

    Public Sub Close()
        If Not m_Process.HasExited Then
            m_Process.Kill()
        End If
        m_Process.Close()
    End Sub

    ''' <summary>
    ''' Schreibt den im Parameter <paramref name="Text"/> angebenen Text
    ''' auf den Ausgabestrom.
    ''' </summary>
    ''' <param name="Text">
    ''' Text, der auf den Ausgabestrom geschrieben werden soll.
    ''' </param>
    Private Sub StreamInput(ByVal Text As String)
        m_Process.StandardInput.WriteLine(Text)
        m_Process.StandardInput.Flush()
    End Sub

    ''' <summary>
    ''' Konvertiert den Text in <paramref name="Text"/> von der DOS-
    ''' Codepage (OEM) in die Windows-Codepage (ANSI).
    ''' </summary>
    ''' <param name="Text">Text, der konvertiert werden soll.</param>
    ''' <returns>
    ''' Der Text aus <paramref name="Text"/> in der aktuellen Windows-
    ''' Codepage.
    ''' </returns>
    Private Function ConvertFromOem(ByVal Text As String) As String
        Return _
            Encoding.GetEncoding( _
                CultureInfo.InstalledUICulture.TextInfo.OEMCodePage _
            ).GetString(Encoding.Default.GetBytes(Text))
    End Function

    ''' <summary>
    ''' Liest vom Ausgabestream und gibt die gelesenen Informationen aus.
    ''' </summary>
    Private Sub StreamOutput()
        Dim Line As String = m_Process.StandardOutput.ReadLine()
        Try
            Do While Line.Length >= 0
                If Line.Length > 0 Then
                    RaiseEvent OnLineAdded(ConvertFromOem(Line))
                End If
                Line = m_Process.StandardOutput.ReadLine()
            Loop
        Catch
            RaiseEvent OnLineAdded(String.Format("""{0}"" wurde beendet!", m_Process.StartInfo.FileName))
        End Try
    End Sub

    ''' <summary>
    ''' Liest vom Fehlerstream und gibt die gelesenen Informationen aus.
    ''' </summary>
    Private Sub StreamError()
        Dim Line As String = m_Process.StandardError.ReadLine()
        Try
            Do While Line.Length >= 0
                Line = m_Process.StandardError.ReadLine()
                If Line.Length > 0 Then
                    RaiseEvent OnLineAdded(Line)
                End If
            Loop
        Catch
            RaiseEvent OnLineAdded(String.Format("""{0}"" wurde beendet!", m_Process.StartInfo.FileName))
        End Try
    End Sub

End Class

