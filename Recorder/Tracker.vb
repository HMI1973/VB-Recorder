Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Tracker
    Dim m_CMDs As Commands

    Public Property CMDs As Commands
        Set(value As Commands)
            m_CMDs = value
        End Set
        Get
            Return m_CMDs
        End Get
    End Property
    Public Overloads Sub Show()
        Dim ScreenSize As Rectangle = Screen.PrimaryScreen.Bounds
        Me.Location = New Drawing.Point(ScreenSize.Width - 200, ScreenSize.Height - Me.Height - 50)
        Me.Visible = True

        MyBase.Show()
    End Sub

    Public Sub updateProgressBar()
        Dim value As Integer = m_CMDs.m_CurLine
        ProgressBar.Maximum = m_CMDs.m_MaxLines
        ProgressBar.Minimum = 0

        If value > ProgressBar.Maximum Then value = ProgressBar.Maximum
        If value < ProgressBar.Minimum Then value = ProgressBar.Minimum
        ProgressBar.Value = value
    End Sub
    Private Sub UpdateText()
        Dim point As SPOINT = DLLFns.GetCursorPosition()
        Dim ActWND() As String = DLLFns.GetActiveWinTitle()

        XYPos.Text = "(" & point.X & "," & point.Y & ")"
        WinTxt.Text = ActWND(1)
        updateProgressBar()
    End Sub

    Private Sub StopBtn_Click(sender As Object, e As EventArgs) Handles StopBtn.Click
        m_CMDs.m_ImerganceBreak = True
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        UpdateText()
    End Sub
End Class