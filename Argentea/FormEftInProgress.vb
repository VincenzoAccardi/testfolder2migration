#Region "Documentation"
' ********** ********** ********** **********
' Argentea EFT
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region

Public Class FormEftInProgress
    Inherits TPDotnet.Pos.FormBase
    Public WithEvents ProgressBarRunning As System.Windows.Forms.ProgressBar
    Public WithEvents LabelArgenteaWait As System.Windows.Forms.Label
    Public WithEvents cmdCancel As TPDotnet.GUIControls.WinButtonEx

    Public Sub New()

        If Not Me.DesignMode Then
            InitializeComponent()
        End If

    End Sub


    Private Sub InitializeComponent()
        Me.cmdCancel = New TPDotnet.GUIControls.WinButtonEx()
        Me.ProgressBarRunning = New System.Windows.Forms.ProgressBar()
        Me.LabelArgenteaWait = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Visible = False
        '
        'WatchScanner
        '
        Me.WatchScanner.Enabled = True
        Me.WatchScanner.Interval = 1000
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DeactiveTextColor = System.Drawing.SystemColors.GrayText
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.cmdCancel.ForeColor = System.Drawing.Color.White
        Me.cmdCancel.Location = New System.Drawing.Point(121, 517)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(110, 48)
        Me.cmdCancel.TabIndex = 19
        Me.cmdCancel.Text = "Annulla"
        Me.cmdCancel.UseVisualStyleBackColor = False
        Me.cmdCancel.Visible = False
        '
        'ProgressBarRunning
        '
        Me.ProgressBarRunning.BackColor = System.Drawing.Color.White
        Me.ProgressBarRunning.ForeColor = System.Drawing.SystemColors.Desktop
        Me.ProgressBarRunning.Location = New System.Drawing.Point(25, 385)
        Me.ProgressBarRunning.Name = "ProgressBarRunning"
        Me.ProgressBarRunning.Size = New System.Drawing.Size(245, 23)
        Me.ProgressBarRunning.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBarRunning.TabIndex = 20
        Me.ProgressBarRunning.Visible = False
        '
        'LabelArgenteaWait
        '
        Me.LabelArgenteaWait.AutoSize = True
        Me.LabelArgenteaWait.BackColor = System.Drawing.Color.Transparent
        Me.LabelArgenteaWait.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelArgenteaWait.Location = New System.Drawing.Point(22, 359)
        Me.LabelArgenteaWait.Name = "LabelArgenteaWait"
        Me.LabelArgenteaWait.Size = New System.Drawing.Size(96, 17)
        Me.LabelArgenteaWait.TabIndex = 21
        Me.LabelArgenteaWait.Text = "Please wait..."
        Me.LabelArgenteaWait.Visible = False
        '
        'FormEftInProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 19.0!)
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.ProgressBarRunning)
        Me.Controls.Add(Me.LabelArgenteaWait)
        Me.Name = "FormEftInProgress"
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.LabelArgenteaWait, 0)
        Me.Controls.SetChildIndex(Me.ProgressBarRunning, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


    Protected Overrides Sub WatchScanner_Tick(ByVal sender As Object, ByVal e As System.EventArgs)


        Me.ProgressBarRunning.PerformStep()
        'If ProgressBarRunning.Value = ProgressBarRunning.Maximum Then
        '    Me.cmdCancel.Enabled = True
        '    Me.cmdCancel.Visible = True
        'End If

        'MyBase.WatchScanner_Tick(sender, e)

    End Sub

    Private Sub FormEftInProgress_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
    Public Property ProgBarValue As Integer
        Get
            Return ProgressBarRunning.Value
        End Get
        Set(value As Integer)
            ProgressBarRunning.Value = value
            ProgressBarRunning.Refresh()
        End Set
    End Property
    Public Property ProgBarMax As Integer
        Get
            Return ProgressBarRunning.Maximum
        End Get
        Set(value As Integer)
            ProgressBarRunning.Maximum = value
        End Set
    End Property
    Public Property EnableCancel As Boolean
        Get
            Return cmdCancel.Enabled
        End Get
        Set(value As Boolean)
            cmdCancel.Enabled = value
            cmdCancel.Visible = value
            cmdCancel.Refresh()
            Me.Refresh()
        End Set
    End Property
End Class
