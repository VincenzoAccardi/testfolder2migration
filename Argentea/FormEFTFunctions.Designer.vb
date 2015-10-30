<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormEFTFunctions
    Inherits TPDotnet.Pos.FormBase

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
        Me.lblNumberToken = New System.Windows.Forms.Label()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.lblAmountToken = New System.Windows.Forms.Label()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.lblLastValidPayment = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbGetTotals = New System.Windows.Forms.RadioButton()
        Me.rbEFTClose = New System.Windows.Forms.RadioButton()
        Me.rbVoidLT = New System.Windows.Forms.RadioButton()
        Me.GetStatus = New System.Windows.Forms.RadioButton()
        Me.cmdOK = New TPDotnet.GUIControls.WinButtonEx()
        Me.cmdCancel = New TPDotnet.GUIControls.WinButtonEx()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Visible = False
        '
        'lblNumberToken
        '
        Me.lblNumberToken.AutoSize = True
        Me.lblNumberToken.BackColor = System.Drawing.Color.Transparent
        Me.lblNumberToken.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.lblNumberToken.Location = New System.Drawing.Point(111, 294)
        Me.lblNumberToken.Name = "lblNumberToken"
        Me.lblNumberToken.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumberToken.Size = New System.Drawing.Size(32, 17)
        Me.lblNumberToken.TabIndex = 74
        Me.lblNumberToken.Text = "xxx"
        '
        'lblNumber
        '
        Me.lblNumber.AutoSize = True
        Me.lblNumber.BackColor = System.Drawing.Color.Transparent
        Me.lblNumber.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.lblNumber.Location = New System.Drawing.Point(36, 294)
        Me.lblNumber.Name = "lblNumber"
        Me.lblNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumber.Size = New System.Drawing.Size(69, 17)
        Me.lblNumber.TabIndex = 73
        Me.lblNumber.Text = "Number:"
        '
        'lblAmountToken
        '
        Me.lblAmountToken.AutoSize = True
        Me.lblAmountToken.BackColor = System.Drawing.Color.Transparent
        Me.lblAmountToken.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.lblAmountToken.Location = New System.Drawing.Point(111, 268)
        Me.lblAmountToken.Name = "lblAmountToken"
        Me.lblAmountToken.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAmountToken.Size = New System.Drawing.Size(32, 17)
        Me.lblAmountToken.TabIndex = 72
        Me.lblAmountToken.Text = "xxx"
        '
        'lblAmount
        '
        Me.lblAmount.AutoSize = True
        Me.lblAmount.BackColor = System.Drawing.Color.Transparent
        Me.lblAmount.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.lblAmount.Location = New System.Drawing.Point(36, 268)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAmount.Size = New System.Drawing.Size(69, 17)
        Me.lblAmount.TabIndex = 69
        Me.lblAmount.Text = "Amount:"
        '
        'lblLastValidPayment
        '
        Me.lblLastValidPayment.AutoSize = True
        Me.lblLastValidPayment.BackColor = System.Drawing.Color.Transparent
        Me.lblLastValidPayment.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.lblLastValidPayment.Location = New System.Drawing.Point(36, 240)
        Me.lblLastValidPayment.Name = "lblLastValidPayment"
        Me.lblLastValidPayment.Size = New System.Drawing.Size(137, 17)
        Me.lblLastValidPayment.TabIndex = 68
        Me.lblLastValidPayment.Text = "Last valid payment"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.rbGetTotals)
        Me.GroupBox1.Controls.Add(Me.rbEFTClose)
        Me.GroupBox1.Controls.Add(Me.rbVoidLT)
        Me.GroupBox1.Controls.Add(Me.GetStatus)
        Me.GroupBox1.Location = New System.Drawing.Point(33, 50)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(287, 163)
        Me.GroupBox1.TabIndex = 67
        Me.GroupBox1.TabStop = False
        '
        'rbGetTotals
        '
        Me.rbGetTotals.AutoSize = True
        Me.rbGetTotals.Checked = True
        Me.rbGetTotals.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.rbGetTotals.Location = New System.Drawing.Point(12, 126)
        Me.rbGetTotals.Name = "rbGetTotals"
        Me.rbGetTotals.Size = New System.Drawing.Size(96, 21)
        Me.rbGetTotals.TabIndex = 21
        Me.rbGetTotals.TabStop = True
        Me.rbGetTotals.Tag = "4"
        Me.rbGetTotals.Text = "Get Totals"
        Me.rbGetTotals.UseVisualStyleBackColor = True
        '
        'rbEFTClose
        '
        Me.rbEFTClose.AutoSize = True
        Me.rbEFTClose.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.rbEFTClose.Location = New System.Drawing.Point(12, 76)
        Me.rbEFTClose.Name = "rbEFTClose"
        Me.rbEFTClose.Size = New System.Drawing.Size(93, 21)
        Me.rbEFTClose.TabIndex = 20
        Me.rbEFTClose.Tag = "2"
        Me.rbEFTClose.Text = "Close EFT"
        Me.rbEFTClose.UseVisualStyleBackColor = True
        '
        'rbVoidLT
        '
        Me.rbVoidLT.AutoSize = True
        Me.rbVoidLT.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.rbVoidLT.Location = New System.Drawing.Point(12, 26)
        Me.rbVoidLT.Name = "rbVoidLT"
        Me.rbVoidLT.Size = New System.Drawing.Size(167, 21)
        Me.rbVoidLT.TabIndex = 19
        Me.rbVoidLT.Tag = "1"
        Me.rbVoidLT.Text = "Void last transaction"
        Me.rbVoidLT.UseVisualStyleBackColor = True
        '
        'GetStatus
        '
        Me.GetStatus.AutoSize = True
        Me.GetStatus.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.GetStatus.Location = New System.Drawing.Point(6, 164)
        Me.GetStatus.Name = "GetStatus"
        Me.GetStatus.Size = New System.Drawing.Size(99, 21)
        Me.GetStatus.TabIndex = 15
        Me.GetStatus.Tag = "3"
        Me.GetStatus.Text = "Get Status"
        Me.GetStatus.UseVisualStyleBackColor = True
        Me.GetStatus.Visible = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.Color.Transparent
        Me.cmdOK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.DeactiveTextColor = System.Drawing.SystemColors.GrayText
        Me.cmdOK.FlatAppearance.BorderSize = 0
        Me.cmdOK.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.cmdOK.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdOK.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.cmdOK.ForeColor = System.Drawing.Color.White
        Me.cmdOK.Location = New System.Drawing.Point(83, 464)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(207, 48)
        Me.cmdOK.TabIndex = 66
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DeactiveTextColor = System.Drawing.SystemColors.GrayText
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.cmdCancel.ForeColor = System.Drawing.Color.White
        Me.cmdCancel.Location = New System.Drawing.Point(83, 537)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(207, 48)
        Me.cmdCancel.TabIndex = 63
        Me.cmdCancel.Text = "Annulla"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'FormEFTFunctions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.Controls.Add(Me.lblNumberToken)
        Me.Controls.Add(Me.lblNumber)
        Me.Controls.Add(Me.lblAmountToken)
        Me.Controls.Add(Me.lblAmount)
        Me.Controls.Add(Me.lblLastValidPayment)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Name = "FormEFTFunctions"
        Me.Text = "FormEFTHandling"
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdOK, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.lblLastValidPayment, 0)
        Me.Controls.SetChildIndex(Me.lblAmount, 0)
        Me.Controls.SetChildIndex(Me.lblAmountToken, 0)
        Me.Controls.SetChildIndex(Me.lblNumber, 0)
        Me.Controls.SetChildIndex(Me.lblNumberToken, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblNumberToken As System.Windows.Forms.Label
    Friend WithEvents lblNumber As System.Windows.Forms.Label
    Friend WithEvents lblAmountToken As System.Windows.Forms.Label
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents lblLastValidPayment As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbGetTotals As System.Windows.Forms.RadioButton
    Friend WithEvents rbEFTClose As System.Windows.Forms.RadioButton
    Friend WithEvents rbVoidLT As System.Windows.Forms.RadioButton
    Friend WithEvents GetStatus As System.Windows.Forms.RadioButton
    Public WithEvents cmdOK As TPDotnet.GUIControls.WinButtonEx
    Public WithEvents cmdCancel As TPDotnet.GUIControls.WinButtonEx
End Class
