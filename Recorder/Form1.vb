Imports System.ComponentModel
Imports System.IO

Public Class Form1
    Const CheckingInterval As Integer = 1 'seconds
    Const DATA_PATH As String = "\Data\"
    Const Error_MSG1 As String = "Forced Stop at ( {0} )"
    Const Error_MSG2 As String = "Error command at ( {0} )"
    Const Error_MSG3 As String = "Start File: ( {0} )"

    Dim Bussy_Flag As Boolean = False
    Dim LastTime As Date
    Dim m_FileName As String = ""
    Dim m_CMDs As New Commands()

#Region "Private Functions"
    Private Sub SetProgStatus(ByVal Optional Status As String = "")
        Dim Template As String = "Step {1} of {2} ( {3}% )"
        Dim ProgPercent As Double = (Int(1000 * (m_CMDs.m_CurLine / m_CMDs.m_MaxLines)) / 10)

        If m_CMDs.m_MaxLines = 0 Then ProgPercent = 0
        Template = Template.Replace("{1}", m_CMDs.m_CurLine)
        Template = Template.Replace("{2}", m_CMDs.m_MaxLines)
        Template = Template.Replace("{3}", ProgPercent)
        If Status = "" Then
            StatusToolStripMenuItem.Text = "Status: " & Template
        Else
            StatusToolStripMenuItem.Text = "Status: " & Status
        End If
    End Sub
    Public Sub updateProgressBar()
        Dim value As Integer = m_CMDs.m_CurLine
        ProgressBar1.Maximum = m_CMDs.m_MaxLines
        ProgressBar1.Minimum = 0

        If value > ProgressBar1.Maximum Then value = ProgressBar1.Maximum
        If value < ProgressBar1.Minimum Then value = ProgressBar1.Minimum
        ProgressBar1.Value = value
        Application.DoEvents()
    End Sub
    'show message in the context menu where -1:IDLE,0: Running Normal, >1 Error in line,<-1 Force break at line
    Private Sub ShowMsg()
        Dim MsgNum As Integer = m_CMDs.m_Status
        Dim Msg As String = "IDLE"

        If MsgNum = 0 Then Msg = "" 'Running Normally 
        If MsgNum = -1 And Not m_FileName = "" Then Error_MSG3.Replace("{0}", m_FileName) 'Start New File
        If MsgNum > 0 Then Msg = Error_MSG2.Replace("{0}", MsgNum) 'Error in line
        If MsgNum < -1 Then Msg = Error_MSG1.Replace("{0}", -MsgNum) 'force break at line

        SetProgStatus(Msg)
    End Sub
    'Update file list menu in the App context menu
    Private Sub UpdateFilesList()
        Dim SearchFiles = Directory.EnumerateFiles(Application.StartupPath & DATA_PATH, "*.txt")

        RecordFilesToolStripMenuItem.DropDownItems.Clear()
        For Index = 0 To SearchFiles.Count - 1
            Dim FileName As String = SearchFiles(Index)
            Dim x As ToolStripMenuItem

            FileName = FileName.Substring(FileName.LastIndexOf("\") + 1)
            x = New ToolStripMenuItem(FileName, Nothing, New EventHandler(AddressOf FileListItemClickHandler))
            x.CheckOnClick = True
            If FileName = m_FileName Then x.Checked = True

            RecordFilesToolStripMenuItem.DropDownItems.Add(x)
        Next
        ShowMsg()
        updateProgressBar()
    End Sub
    'check if parameter passed to EXE
    Private Sub CheckCMDPRMT()
        Dim CmdLine() As String = Environment.CommandLine().Split("""")

        If CmdLine.Count > 2 Then
            Dim tmpFileName As String = CmdLine(2).Replace("""", "").Trim
            If tmpFileName.IndexOf(":\") = -1 Then tmpFileName = Application.StartupPath & DATA_PATH & tmpFileName
            If Path.GetExtension(tmpFileName) = "" Then tmpFileName = tmpFileName & ".txt"
            If File.Exists(tmpFileName) Then
                m_FileName = tmpFileName
                DoFun()
            End If
        End If
    End Sub
    Private Sub DoFun()
        Dim FullFilename As String = Application.StartupPath & DATA_PATH & m_FileName

        If m_FileName = "" Then Return
        If Not File.Exists(FullFilename) Then Return

        'Change App Icon and show "stop" button from info menu
        NotifyIcon1.Icon = My.Resources.StopIcn : Tracker.StopBtn.Visible = True : Application.DoEvents()

        m_CMDs.DoFile(FullFilename)
        'Change App Icon and hide "Stop" button from info menu
        NotifyIcon1.Icon = My.Resources.PlayIcn : Tracker.StopBtn.Visible = False : Application.DoEvents()

        m_FileName = ""
        UpdateFilesList()
    End Sub
#End Region

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If
        MyBase.SetVisibleCore(value)

        Tracker.CMDs = m_CMDs
        CheckCMDPRMT()
        UpdateFilesList()
        updateProgressBar()
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Application.Exit()
    End Sub
    Private Sub ExitServiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitServiceToolStripMenuItem.Click
        m_CMDs.m_ImerganceBreak = True
        Timer1.Enabled = False
        Application.Exit()
    End Sub
    Private Sub FileListItemClickHandler(sender As Object, e As EventArgs)
        Dim clickedItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If clickedItem.Checked Then
            If Not m_FileName = "" Then clickedItem.Checked = False : Return
            m_FileName = clickedItem.Text
            UpdateFilesList()
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub
    Private Sub InfoMenuItem_Click(sender As Object, e As EventArgs) Handles InfoMenuItem.Click
        If InfoMenuItem.Checked Then
            Tracker.Show()
        Else
            Tracker.Hide()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Bussy_Flag Then
            ShowMsg()
            Exit Sub
        End If

        If DateAndTime.DateAdd("s", CheckingInterval, LastTime) < Now Then
            Bussy_Flag = True
            DoFun()
            LastTime = Now
            Bussy_Flag = False
        End If
    End Sub

End Class
