Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet
Imports TPDotnet.Pos

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

Public Class FormArgenteaItemInput
    Inherits TPDotnet.Pos.FormBase
    Public WithEvents txtInsertCODE128 As TPDotnet.GUIControls.WinTextBoxControl
    Public WithEvents lblInsertCODE128 As System.Windows.Forms.Label
    Public WithEvents cmdOK As TPDotnet.GUIControls.WinButtonEx
    Public WithEvents lblArticleDescription As System.Windows.Forms.Label
    Public WithEvents cmdCancel As TPDotnet.GUIControls.WinButtonEx

    Public Sub New()

        If Not Me.DesignMode Then
            InitializeComponent()
        End If

    End Sub

    Public Property ArticleDescription() As String
        Get
            If lblArticleDescription Is Nothing Then Return ""
            Return lblArticleDescription.Text
        End Get
        Set(ByVal value As String)
            If Not lblArticleDescription Is Nothing Then
                lblArticleDescription.Text = value
            End If
        End Set
    End Property


    Private Sub InitializeComponent()
        Me.cmdCancel = New TPDotnet.GUIControls.WinButtonEx()
        Me.txtInsertCODE128 = New TPDotnet.GUIControls.WinTextBoxControl()
        Me.lblInsertCODE128 = New System.Windows.Forms.Label()
        Me.cmdOK = New TPDotnet.GUIControls.WinButtonEx()
        Me.lblArticleDescription = New System.Windows.Forms.Label()
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
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.cmdCancel.ForeColor = System.Drawing.Color.White
        Me.cmdCancel.Location = New System.Drawing.Point(186, 517)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(125, 48)
        Me.cmdCancel.TabIndex = 19
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'txtInsertCODE128
        '
        Me.txtInsertCODE128.AcceptsReturn = False
        Me.txtInsertCODE128.BackColor = System.Drawing.Color.White
        Me.txtInsertCODE128.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.txtInsertCODE128.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInsertCODE128.Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Bold)
        Me.txtInsertCODE128.ForeColor = System.Drawing.Color.FromArgb(CType(CType(151, Byte), Integer), CType(CType(20, Byte), Integer), CType(CType(6, Byte), Integer))
        Me.txtInsertCODE128.LeftBorderLimit = 10
        Me.txtInsertCODE128.Location = New System.Drawing.Point(25, 168)
        Me.txtInsertCODE128.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        '
        '
        '
        Me.txtInsertCODE128.MaskedTextBox.BackColor = System.Drawing.Color.White
        Me.txtInsertCODE128.MaskedTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInsertCODE128.MaskedTextBox.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInsertCODE128.MaskedTextBox.ForeColor = System.Drawing.Color.FromArgb(CType(CType(151, Byte), Integer), CType(CType(20, Byte), Integer), CType(CType(6, Byte), Integer))
        Me.txtInsertCODE128.MaskedTextBox.Location = New System.Drawing.Point(-1, -1)
        Me.txtInsertCODE128.MaskedTextBox.Name = "TextBox"
        Me.txtInsertCODE128.MaskedTextBox.Size = New System.Drawing.Size(269, 26)
        Me.txtInsertCODE128.MaskedTextBox.TabIndex = 0
        Me.txtInsertCODE128.MaskedTextBox.TextMaskFormat = System.Windows.Forms.MaskFormat.IncludePromptAndLiterals
        '
        '
        '
        Me.txtInsertCODE128.MaskedTextBoxContainer.BackColor = System.Drawing.Color.White
        Me.txtInsertCODE128.MaskedTextBoxContainer.Controls.Add(Me.txtInsertCODE128.MaskedTextBox)
        Me.txtInsertCODE128.MaskedTextBoxContainer.Location = New System.Drawing.Point(10, 6)
        Me.txtInsertCODE128.MaskedTextBoxContainer.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.txtInsertCODE128.MaskedTextBoxContainer.Name = "TextBox"
        Me.txtInsertCODE128.MaskedTextBoxContainer.Size = New System.Drawing.Size(266, 21)
        Me.txtInsertCODE128.MaskedTextBoxContainer.TabIndex = 0
        Me.txtInsertCODE128.MaxLength = 2147483647
        Me.txtInsertCODE128.MaxValue = New Decimal(New Integer() {-1, -1, -1, 0})
        Me.txtInsertCODE128.MinValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtInsertCODE128.Name = "txtInsertCODE128"
        Me.txtInsertCODE128.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.txtInsertCODE128.PromptChar = Global.Microsoft.VisualBasic.ChrW(95)
        Me.txtInsertCODE128.SelectionLength = 0
        Me.txtInsertCODE128.SelectionStart = 0
        Me.txtInsertCODE128.Size = New System.Drawing.Size(286, 33)
        Me.txtInsertCODE128.TabIndex = 2
        Me.txtInsertCODE128.ValidationTabOnEnter = False
        '
        'lblInsertCODE128
        '
        Me.lblInsertCODE128.BackColor = System.Drawing.Color.Transparent
        Me.lblInsertCODE128.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblInsertCODE128.ForeColor = System.Drawing.Color.FromArgb(CType(CType(85, Byte), Integer), CType(CType(142, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.lblInsertCODE128.Location = New System.Drawing.Point(22, 129)
        Me.lblInsertCODE128.Name = "lblInsertCODE128"
        Me.lblInsertCODE128.Size = New System.Drawing.Size(289, 35)
        Me.lblInsertCODE128.TabIndex = 1
        Me.lblInsertCODE128.Text = "Gift card barcode"
        Me.lblInsertCODE128.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        Me.cmdOK.Location = New System.Drawing.Point(25, 517)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(125, 48)
        Me.cmdOK.TabIndex = 24
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'lblArticleDescription
        '
        Me.lblArticleDescription.BackColor = System.Drawing.Color.Transparent
        Me.lblArticleDescription.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblArticleDescription.ForeColor = System.Drawing.Color.DarkRed
        Me.lblArticleDescription.Location = New System.Drawing.Point(22, 94)
        Me.lblArticleDescription.Name = "lblArticleDescription"
        Me.lblArticleDescription.Size = New System.Drawing.Size(289, 35)
        Me.lblArticleDescription.TabIndex = 25
        Me.lblArticleDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FormArgenteaItemInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 19.0!)
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.Controls.Add(Me.lblArticleDescription)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblInsertCODE128)
        Me.Controls.Add(Me.txtInsertCODE128)
        Me.Controls.Add(Me.cmdCancel)
        Me.Name = "FormArgenteaItemInput"
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.txtInsertCODE128, 0)
        Me.Controls.SetChildIndex(Me.lblInsertCODE128, 0)
        Me.Controls.SetChildIndex(Me.cmdOK, 0)
        Me.Controls.SetChildIndex(Me.lblArticleDescription, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


    Protected Overrides Sub WatchScanner_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        System.Diagnostics.Debug.Print("WatchScanner_Tick")

        If Me.DesignMode Then
            ' We are in design mode. Important when this form is used as inherited form, 
            ' then the following code should not be passed through
            Exit Sub
        End If

        Try
            LOG_FuncStart(getLocationString("WatchScanner_Tick"))

            Me.WatchScanner.Enabled = False

            ' check, if data from scanner or msr
            If CheckScannerData(theModCntr) = True Then

                Me.txtInsertCODE128.Clear()
                Me.txtInsertCODE128.Text = m_szScannedData

                Me.cmdOK.PerformClick()

            End If

            Exit Sub

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("WatchScanner_Tick"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("WatchScanner_Tick"), InnerEx)
            End Try
        Finally
            ResetScannerInfo()
            Me.WatchScanner.Enabled = True
            LOG_FuncExit(getLocationString("WatchScanner_Tick"), "")
        End Try
    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click

        If String.IsNullOrEmpty(Me.txtInsertCODE128.Text) OrElse String.IsNullOrWhiteSpace(Me.txtInsertCODE128.Text) Then
            Return
        End If

        If Me.txtInsertCODE128.Text.Contains("=") Then
            ' consider the left part before the "="
            Me.txtInsertCODE128.Text = Me.txtInsertCODE128.Text.Split("=")(0)
        End If

        If String.IsNullOrEmpty(Me.txtInsertCODE128.Text) OrElse String.IsNullOrWhiteSpace(Me.txtInsertCODE128.Text) Then
            Me.txtInsertCODE128.Clear()
            Return
        End If

        Me.Tag = Me.txtInsertCODE128.Text
        Me.bDialogActive = False

    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        Me.Tag = String.Empty
        Me.bDialogActive = False

    End Sub

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

   
End Class
