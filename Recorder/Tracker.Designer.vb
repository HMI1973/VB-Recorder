<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Tracker
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.XYPos = New System.Windows.Forms.Label()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.StopBtn = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.WinTxt = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'XYPos
        '
        Me.XYPos.Font = New System.Drawing.Font("Arial Narrow", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.XYPos.Location = New System.Drawing.Point(0, 5)
        Me.XYPos.Name = "XYPos"
        Me.XYPos.Size = New System.Drawing.Size(200, 20)
        Me.XYPos.TabIndex = 0
        Me.XYPos.Text = "X,Y"
        Me.XYPos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(0, 108)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(200, 10)
        Me.ProgressBar.TabIndex = 1
        '
        'StopBtn
        '
        Me.StopBtn.BackColor = System.Drawing.Color.Red
        Me.StopBtn.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.StopBtn.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StopBtn.ForeColor = System.Drawing.Color.White
        Me.StopBtn.Location = New System.Drawing.Point(0, 76)
        Me.StopBtn.Name = "StopBtn"
        Me.StopBtn.Size = New System.Drawing.Size(200, 30)
        Me.StopBtn.TabIndex = 2
        Me.StopBtn.Text = "STOP"
        Me.StopBtn.UseVisualStyleBackColor = False
        Me.StopBtn.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'WinTxt
        '
        Me.WinTxt.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WinTxt.ForeColor = System.Drawing.Color.Black
        Me.WinTxt.Location = New System.Drawing.Point(0, 25)
        Me.WinTxt.Name = "WinTxt"
        Me.WinTxt.Size = New System.Drawing.Size(200, 45)
        Me.WinTxt.TabIndex = 3
        Me.WinTxt.Text = "X,Y"
        '
        'Tracker
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(200, 120)
        Me.ControlBox = False
        Me.Controls.Add(Me.WinTxt)
        Me.Controls.Add(Me.StopBtn)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.XYPos)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Tracker"
        Me.Opacity = 0.75R
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Tracker"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents XYPos As Label
    Friend WithEvents ProgressBar As ProgressBar
    Friend WithEvents StopBtn As Button
    Friend WithEvents Timer1 As Timer
    Friend WithEvents WinTxt As Label
End Class
