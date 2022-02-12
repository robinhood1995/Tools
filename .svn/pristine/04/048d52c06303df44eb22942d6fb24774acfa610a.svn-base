Option Strict On
Option Explicit On

Imports System.Windows.Forms
Public Class clsAutoCompleteCombo
    Inherits ComboBox
    Private mResetOnClear As Boolean = False

    Protected Overrides Sub RefreshItem(ByVal index As Integer)
        MyBase.RefreshItem(index)
    End Sub

    Protected Overrides Sub SetItemsCore(ByVal items As System.Collections.IList)
        MyBase.SetItemsCore(items)
    End Sub

    Public Shadows Sub KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim intIndex As Integer
        Dim strEntry As String

        If Char.IsControl(e.KeyChar) Then
            If MyBase.SelectionStart <= 1 Then
                If mResetOnClear Then
                    MyBase.SelectedIndex = 0
                    MyBase.SelectAll()
                Else
                    MyBase.Text = String.Empty
                    MyBase.SelectedIndex = -1
                End If
                e.Handled = True
                Exit Sub
            End If
            If MyBase.SelectionLength = 0 Then
                strEntry = MyBase.Text.Substring(0, MyBase.Text.Length - 1)
            Else
                strEntry = MyBase.Text.Substring(0, MyBase.SelectionStart - 1)
            End If
        ElseIf (Not Char.IsLetterOrDigit(e.KeyChar)) And (Not Char.IsWhiteSpace(e.KeyChar)) Then  '< 32 Or KeyAscii > 127 Then
            Exit Sub
        Else
            If MyBase.SelectionLength = 0 Then
                strEntry = UCase(MyBase.Text & e.KeyChar)
            Else
                strEntry = MyBase.Text.Substring(0, MyBase.SelectionStart) & e.KeyChar
            End If
        End If

        intIndex = MyBase.FindString(strEntry)

        If intIndex <> -1 Then
            MyBase.SelectedIndex = intIndex
            MyBase.SelectionStart = strEntry.Length
            MyBase.SelectionLength = MyBase.Text.Length - MyBase.SelectionStart
        End If
        e.Handled = True
        Exit Sub
    End Sub

    Public Property ResetOnClear() As Boolean
        Get
            Return mResetOnClear
        End Get
        Set(ByVal Value As Boolean)
            mResetOnClear = Value
        End Set
    End Property
End Class
