Namespace Onling

    '-------------------------------------------------------------------
    'Using this class
    'Private m_clsProcess As ProcessClass

    'Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
    '    'clear treeview...
    '    tvMain.Nodes.Clear()
    '    'create a new process class if we haven't already...
    '    If m_clsProcess Is Nothing Then
    '        m_clsProcess = New ProcessClass(Me, New ProcessClass.NotifyProgress(AddressOf DelegateProgress))
    '    End If
    '    'kick off the process, execution returns immediately to the next line...
    '    m_clsProcess.Start()
    'End Sub

    ''this routine was declared to handle notify messages sent via delegate...
    'Private Sub DelegateProgress(ByVal Message As String, ByVal PercentComplete As Integer)
    '    'display progress in a label and progress bar - this won't error across threads...
    '    lblProgress.Text = String.Concat(Message, " [DELEGATE]")
    '    pbMain.Value = PercentComplete
    '    'display progress in a treeview, this will raise an error if you don't properly marshal the call across threads...
    '    tvMain.Nodes.Add(String.Concat(Message, " - ", PercentComplete)).EnsureVisible()
    'End Sub

    Public Class clsProcessClass

        Private m_clsNotifyDelegate As NotifyProgress
        Private m_clsThread As System.Threading.Thread
        Private m_clsSynchronizingObject As System.ComponentModel.ISynchronizeInvoke

        Public Delegate Sub NotifyProgress(ByVal Message As String, ByVal PercentComplete As Integer)

        Public Sub New(ByVal SynchronizingObject As System.ComponentModel.ISynchronizeInvoke, _
                ByVal NotifyDelegate As NotifyProgress)
            m_clsSynchronizingObject = SynchronizingObject
            m_clsNotifyDelegate = NotifyDelegate
        End Sub

        Public Sub Start()
            m_clsThread = New System.Threading.Thread(AddressOf DoProcess)
            m_clsThread.Name = "My Background Thread"
            m_clsThread.IsBackground = True
            m_clsThread.Start()
        End Sub

        Private Sub DoProcess()
            For i As Integer = 1 To 100
                NotifyUI("Processing", i)
                m_clsThread.Sleep(100)
            Next
            NotifyUI("Processing", 100)
        End Sub

        Private Sub NotifyUI(ByVal Message As String, ByVal Value As Integer)
            'this method will fail because we're not telling the delegate which thread to run in...
            'm_clsNotifyDelegate(Message, Value)
            'build argument list...
            Dim args(1) As Object
            args(0) = Message
            args(1) = Value
            'call the delegate, specifying the context in which to run...
            m_clsSynchronizingObject.Invoke(m_clsNotifyDelegate, args)
        End Sub

    End Class

End Namespace