Imports log4net
Imports log4net.Config
Imports System.Configuration
Imports System.Reflection
Imports System.Windows.Forms

Public Class clsAutoCombo
    Sub New(ByVal cboKeyUp As ComboBox, ByVal eKeyUp As KeyEventArgs)
        AutoCompleteCombo_KeyUp(cboKeyUp, eKeyUp)
    End Sub
    Sub New(ByVal cboLeave As ComboBox)
        AutoCompleteCombo_Leave(cboLeave)
    End Sub

#Region " Properties "

#End Region

    ''' <summary>
    ''' Auto Complete Combo on Key Down
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <param name="e"></param>
    ''' <remarks>AutoCompleteCombo_KeyUp(cboName, e)</remarks>
    Public Sub AutoCompleteCombo_KeyUp(ByVal cbo As ComboBox, ByVal e As KeyEventArgs)
        Dim sTypedText As String
        Dim iFoundIndex As Integer
        Dim oFoundItem As Object
        Dim sFoundText As String
        Dim sAppendText As String

        'Allow select keys without Autocompleting

        Select Case e.KeyCode
            Case Keys.Back, Keys.Left, Keys.Right, Keys.Up, Keys.Delete, Keys.Down
                Return
        End Select

        'Get the Typed Text and Find it in the list

        sTypedText = cbo.Text
        iFoundIndex = cbo.FindString(sTypedText)

        'If we found the Typed Text in the list then Autocomplete

        If iFoundIndex >= 0 Then

            'Get the Item from the list (Return Type depends if Datasource was bound 

            ' or List Created)

            oFoundItem = cbo.Items(iFoundIndex)

            'Use the ListControl.GetItemText to resolve the Name in case the Combo 

            ' was Data bound

            sFoundText = cbo.GetItemText(oFoundItem)

            'Append then found text to the typed text to preserve case

            sAppendText = sFoundText.Substring(sTypedText.Length)
            cbo.Text = sTypedText & sAppendText

            'Select the Appended Text

            cbo.SelectionStart = sTypedText.Length
            cbo.SelectionLength = sAppendText.Length

        End If

    End Sub
    ''' <summary>
    ''' Auto Complete Combo on Leave
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <remarks>AutoCompleteCombo_Leave(cboName)</remarks>
    Public Sub AutoCompleteCombo_Leave(ByVal cbo As ComboBox)
        Dim iFoundIndex As Integer
        iFoundIndex = cbo.FindStringExact(cbo.Text)
        cbo.SelectedIndex = iFoundIndex
    End Sub
End Class
