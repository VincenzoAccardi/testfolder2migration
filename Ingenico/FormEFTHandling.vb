Option Strict Off
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos


Public Class FormEFTHandling
    Inherits FormBase

#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    'For the start-up form, the first instance created is the default instance.
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
        Try
            'This call is required by the Windows Form Designer.
            InitializeComponent()
        Catch ex As Exception
            ' Heike for testing
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Info: Error in form FormEFT.InitializeComponent")
        End Try
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    'Public WithEvents WatchScanner As System.Windows.Forms.Timer
    Friend WithEvents ProgressBarRunning As System.Windows.Forms.ProgressBar
    Friend WithEvents LabelEFTWait As System.Windows.Forms.Label
    Friend WithEvents RBInitEFT As System.Windows.Forms.RadioButton
    Friend WithEvents RBVoidLastTx As System.Windows.Forms.RadioButton
    Friend WithEvents RBCloseEFT As System.Windows.Forms.RadioButton
    Public WithEvents cmdOK As TPDotnet.GUIControls.WinButtonEx
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RBTotalizerEFT As System.Windows.Forms.RadioButton
    Friend WithEvents LabelEFTLastOKPayment As System.Windows.Forms.Label
    Friend WithEvents LabelEFTLastOKAmount As System.Windows.Forms.Label
    Friend WithEvents LabelEFTNextOpNumToken As System.Windows.Forms.Label
    Friend WithEvents LabelEFTNextOpNum As System.Windows.Forms.Label
    Friend WithEvents LabelEFTLastOKAmountToken As System.Windows.Forms.Label
    Friend WithEvents LabelEFTLastOKOpNumToken As System.Windows.Forms.Label
    Friend WithEvents LabelEFTLastOKOpNum As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents cmdCancel As TPDotnet.GUIControls.WinButtonEx
    Private Sub InitializeComponent()
        Me.cmdCancel = New TPDotnet.GUIControls.WinButtonEx
        Me.ProgressBarRunning = New System.Windows.Forms.ProgressBar
        Me.LabelEFTWait = New System.Windows.Forms.Label
        Me.RBInitEFT = New System.Windows.Forms.RadioButton
        Me.RBVoidLastTx = New System.Windows.Forms.RadioButton
        Me.RBCloseEFT = New System.Windows.Forms.RadioButton
        Me.cmdOK = New TPDotnet.GUIControls.WinButtonEx
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RBTotalizerEFT = New System.Windows.Forms.RadioButton
        Me.LabelEFTLastOKPayment = New System.Windows.Forms.Label
        Me.LabelEFTLastOKAmount = New System.Windows.Forms.Label
        Me.LabelEFTNextOpNumToken = New System.Windows.Forms.Label
        Me.LabelEFTNextOpNum = New System.Windows.Forms.Label
        Me.LabelEFTLastOKAmountToken = New System.Windows.Forms.Label
        Me.LabelEFTLastOKOpNumToken = New System.Windows.Forms.Label
        Me.LabelEFTLastOKOpNum = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Visible = False
        '
        'WatchScanner
        '
        Me.WatchScanner.Enabled = True
        Me.WatchScanner.Interval = 300
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
        Me.cmdCancel.Location = New System.Drawing.Point(179, 522)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(134, 48)
        Me.cmdCancel.TabIndex = 12
        Me.cmdCancel.Text = "Annulla"
        Me.cmdCancel.UseVisualStyleBackColor = False
        Me.cmdCancel.Visible = False
        '
        'ProgressBarRunning
        '
        Me.ProgressBarRunning.BackColor = System.Drawing.Color.White
        Me.ProgressBarRunning.ForeColor = System.Drawing.SystemColors.Desktop
        Me.ProgressBarRunning.Location = New System.Drawing.Point(26, 430)
        Me.ProgressBarRunning.Name = "ProgressBarRunning"
        Me.ProgressBarRunning.Size = New System.Drawing.Size(245, 23)
        Me.ProgressBarRunning.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBarRunning.TabIndex = 13
        Me.ProgressBarRunning.Visible = False
        '
        'LabelEFTWait
        '
        Me.LabelEFTWait.AutoSize = True
        Me.LabelEFTWait.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTWait.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTWait.Location = New System.Drawing.Point(23, 404)
        Me.LabelEFTWait.Name = "LabelEFTWait"
        Me.LabelEFTWait.Size = New System.Drawing.Size(96, 17)
        Me.LabelEFTWait.TabIndex = 14
        Me.LabelEFTWait.Text = "Please wait..."
        Me.LabelEFTWait.Visible = False
        '
        'RBInitEFT
        '
        Me.RBInitEFT.AutoSize = True
        Me.RBInitEFT.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.RBInitEFT.Location = New System.Drawing.Point(6, 119)
        Me.RBInitEFT.Name = "RBInitEFT"
        Me.RBInitEFT.Size = New System.Drawing.Size(113, 21)
        Me.RBInitEFT.TabIndex = 15
        Me.RBInitEFT.Tag = "3"
        Me.RBInitEFT.Text = "Initialize EFT"
        Me.RBInitEFT.UseVisualStyleBackColor = True
        Me.RBInitEFT.Visible = False
        '
        'RBVoidLastTx
        '
        Me.RBVoidLastTx.AutoSize = True
        Me.RBVoidLastTx.Checked = True
        Me.RBVoidLastTx.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.RBVoidLastTx.Location = New System.Drawing.Point(6, 19)
        Me.RBVoidLastTx.Name = "RBVoidLastTx"
        Me.RBVoidLastTx.Size = New System.Drawing.Size(167, 21)
        Me.RBVoidLastTx.TabIndex = 16
        Me.RBVoidLastTx.TabStop = True
        Me.RBVoidLastTx.Tag = "1"
        Me.RBVoidLastTx.Text = "Void last transaction"
        Me.RBVoidLastTx.UseVisualStyleBackColor = True
        Me.RBVoidLastTx.Visible = False
        '
        'RBCloseEFT
        '
        Me.RBCloseEFT.AutoSize = True
        Me.RBCloseEFT.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.RBCloseEFT.Location = New System.Drawing.Point(6, 69)
        Me.RBCloseEFT.Name = "RBCloseEFT"
        Me.RBCloseEFT.Size = New System.Drawing.Size(93, 21)
        Me.RBCloseEFT.TabIndex = 17
        Me.RBCloseEFT.Tag = "2"
        Me.RBCloseEFT.Text = "Close EFT"
        Me.RBCloseEFT.UseVisualStyleBackColor = True
        Me.RBCloseEFT.Visible = False
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
        Me.cmdOK.Location = New System.Drawing.Point(26, 521)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(134, 48)
        Me.cmdOK.TabIndex = 18
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        Me.cmdOK.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.RBTotalizerEFT)
        Me.GroupBox1.Controls.Add(Me.RBInitEFT)
        Me.GroupBox1.Controls.Add(Me.RBCloseEFT)
        Me.GroupBox1.Controls.Add(Me.RBVoidLastTx)
        Me.GroupBox1.Location = New System.Drawing.Point(26, 57)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(245, 201)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Visible = False
        '
        'RBTotalizerEFT
        '
        Me.RBTotalizerEFT.AutoSize = True
        Me.RBTotalizerEFT.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.RBTotalizerEFT.Location = New System.Drawing.Point(6, 165)
        Me.RBTotalizerEFT.Name = "RBTotalizerEFT"
        Me.RBTotalizerEFT.Size = New System.Drawing.Size(115, 21)
        Me.RBTotalizerEFT.TabIndex = 18
        Me.RBTotalizerEFT.Tag = "4"
        Me.RBTotalizerEFT.Text = "Totalizer EFT"
        Me.RBTotalizerEFT.UseVisualStyleBackColor = True
        Me.RBTotalizerEFT.Visible = False
        '
        'LabelEFTLastOKPayment
        '
        Me.LabelEFTLastOKPayment.AutoSize = True
        Me.LabelEFTLastOKPayment.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTLastOKPayment.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTLastOKPayment.Location = New System.Drawing.Point(23, 267)
        Me.LabelEFTLastOKPayment.Name = "LabelEFTLastOKPayment"
        Me.LabelEFTLastOKPayment.Size = New System.Drawing.Size(137, 17)
        Me.LabelEFTLastOKPayment.TabIndex = 20
        Me.LabelEFTLastOKPayment.Text = "Last valid payment"
        Me.LabelEFTLastOKPayment.Visible = False
        '
        'LabelEFTLastOKAmount
        '
        Me.LabelEFTLastOKAmount.AutoSize = True
        Me.LabelEFTLastOKAmount.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTLastOKAmount.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTLastOKAmount.Location = New System.Drawing.Point(66, 295)
        Me.LabelEFTLastOKAmount.Name = "LabelEFTLastOKAmount"
        Me.LabelEFTLastOKAmount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LabelEFTLastOKAmount.Size = New System.Drawing.Size(69, 17)
        Me.LabelEFTLastOKAmount.TabIndex = 21
        Me.LabelEFTLastOKAmount.Text = "Amount:"
        Me.LabelEFTLastOKAmount.Visible = False
        '
        'LabelEFTNextOpNumToken
        '
        Me.LabelEFTNextOpNumToken.AutoSize = True
        Me.LabelEFTNextOpNumToken.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTNextOpNumToken.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTNextOpNumToken.Location = New System.Drawing.Point(236, 352)
        Me.LabelEFTNextOpNumToken.Name = "LabelEFTNextOpNumToken"
        Me.LabelEFTNextOpNumToken.Size = New System.Drawing.Size(32, 17)
        Me.LabelEFTNextOpNumToken.TabIndex = 23
        Me.LabelEFTNextOpNumToken.Text = "xxx"
        Me.LabelEFTNextOpNumToken.Visible = False
        '
        'LabelEFTNextOpNum
        '
        Me.LabelEFTNextOpNum.AutoSize = True
        Me.LabelEFTNextOpNum.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTNextOpNum.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTNextOpNum.Location = New System.Drawing.Point(23, 352)
        Me.LabelEFTNextOpNum.Name = "LabelEFTNextOpNum"
        Me.LabelEFTNextOpNum.Size = New System.Drawing.Size(175, 17)
        Me.LabelEFTNextOpNum.TabIndex = 22
        Me.LabelEFTNextOpNum.Text = "Next operation number:"
        Me.LabelEFTNextOpNum.Visible = False
        '
        'LabelEFTLastOKAmountToken
        '
        Me.LabelEFTLastOKAmountToken.AutoSize = True
        Me.LabelEFTLastOKAmountToken.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTLastOKAmountToken.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTLastOKAmountToken.Location = New System.Drawing.Point(236, 295)
        Me.LabelEFTLastOKAmountToken.Name = "LabelEFTLastOKAmountToken"
        Me.LabelEFTLastOKAmountToken.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LabelEFTLastOKAmountToken.Size = New System.Drawing.Size(32, 17)
        Me.LabelEFTLastOKAmountToken.TabIndex = 24
        Me.LabelEFTLastOKAmountToken.Text = "xxx"
        Me.LabelEFTLastOKAmountToken.Visible = False
        '
        'LabelEFTLastOKOpNumToken
        '
        Me.LabelEFTLastOKOpNumToken.AutoSize = True
        Me.LabelEFTLastOKOpNumToken.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTLastOKOpNumToken.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTLastOKOpNumToken.Location = New System.Drawing.Point(236, 321)
        Me.LabelEFTLastOKOpNumToken.Name = "LabelEFTLastOKOpNumToken"
        Me.LabelEFTLastOKOpNumToken.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LabelEFTLastOKOpNumToken.Size = New System.Drawing.Size(32, 17)
        Me.LabelEFTLastOKOpNumToken.TabIndex = 26
        Me.LabelEFTLastOKOpNumToken.Text = "xxx"
        Me.LabelEFTLastOKOpNumToken.Visible = False
        '
        'LabelEFTLastOKOpNum
        '
        Me.LabelEFTLastOKOpNum.AutoSize = True
        Me.LabelEFTLastOKOpNum.BackColor = System.Drawing.Color.Transparent
        Me.LabelEFTLastOKOpNum.Font = New System.Drawing.Font("Tahoma", 10.5!, System.Drawing.FontStyle.Bold)
        Me.LabelEFTLastOKOpNum.Location = New System.Drawing.Point(66, 321)
        Me.LabelEFTLastOKOpNum.Name = "LabelEFTLastOKOpNum"
        Me.LabelEFTLastOKOpNum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LabelEFTLastOKOpNum.Size = New System.Drawing.Size(69, 17)
        Me.LabelEFTLastOKOpNum.TabIndex = 25
        Me.LabelEFTLastOKOpNum.Text = "Number:"
        Me.LabelEFTLastOKOpNum.Visible = False
        '
        'FormEFTHandling
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.ControlBox = False
        Me.Controls.Add(Me.LabelEFTLastOKOpNumToken)
        Me.Controls.Add(Me.LabelEFTLastOKOpNum)
        Me.Controls.Add(Me.LabelEFTLastOKAmountToken)
        Me.Controls.Add(Me.LabelEFTNextOpNumToken)
        Me.Controls.Add(Me.LabelEFTNextOpNum)
        Me.Controls.Add(Me.LabelEFTLastOKAmount)
        Me.Controls.Add(Me.LabelEFTLastOKPayment)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.ProgressBarRunning)
        Me.Controls.Add(Me.LabelEFTWait)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "FormEFTHandling"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "FormEFT"
        Me.Controls.SetChildIndex(Me.LabelEFTWait, 0)
        Me.Controls.SetChildIndex(Me.ProgressBarRunning, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOK, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTLastOKPayment, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTLastOKAmount, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTNextOpNum, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTNextOpNumToken, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTLastOKAmountToken, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTLastOKOpNum, 0)
        Me.Controls.SetChildIndex(Me.LabelEFTLastOKOpNumToken, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FormEFTHandling
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FormEFTHandling
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FormEFTHandling()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As FormEFTHandling)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
#Region "Variables"
    Public Event EFTDeviceMessage(ByRef szMessage As String)
    Public Event EFTDeviceCheckVoid(ByRef bIsVoid As Boolean)

    Public iListIndex As Short

    Protected blnCtrlVPressed As Boolean
#End Region
#Region "Control-events"

    Protected Overridable Sub cmdCancel_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdCancel.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        On Error Resume Next

        If (KeyCode = System.Windows.Forms.Keys.Return) Then Call cmdCancel_Click(cmdCancel, New System.EventArgs)
    End Sub

    Protected Overridable Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click

        On Error Resume Next
        iListIndex = -1
        RaiseEvent EFTDeviceMessage("Abort")
        'Tag = ""
        Tag = False
        theModCntr.DialogActiv = False
    End Sub

    Protected Overridable Sub FormEFT_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        Dim szActModule As String
        Dim szHelpFileName As String = ""
        Dim sKeyCode As Short = eventArgs.KeyCode
        Dim sShift As Short = eventArgs.KeyData \ &H10000
        Dim sHelpMapID As Short
        Dim bRet As Boolean

        On Error GoTo MyStdErr

        'If sKeyCode = System.Windows.Forms.Keys.F1 Then
        '    szActModule = Me.theModCntr.getPosFormValue("lblPosModel.Caption")
        '    sHelpMapID = GetHelpMapID(szHelpFileName, Me.theModCntr.contxt, szActModule, Me.Name)
        '    bRet = TestForHelp(sKeyCode, sShift, sHelpMapID, szHelpFileName)
        'End If
        Exit Sub

MyStdErr:
        'Trace "FormEFT.Form_KeyDown", "Number = " & Err.Number & " Message = " & Err.Description
        Resume Next
    End Sub

    '    Protected Overridable Sub FormEFT_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

    '        Dim bRet As Boolean

    '        On Error GoTo MyStdErr

    '        If Not theModCntr Is Nothing Then
    '            GotFocusEvent4VirtualKeyboard((Me.theModCntr), Me.Handle.ToInt32, True)
    '            ' set top,left,width and height first for reducing flicker
    '            'Call TheModCntr.setSubFormPosition(Me)
    '            'Call LoadFormEventNew(theModCntr, Me, actRectOriginal, True)
    '            'bRet = TheModCntr.setSubFormValues(Me, 1)
    '            'libo_Menu.Font = VB6.FontChangeName(libo_Menu.Font, "Arial")
    '            'libo_Menu.Font = VB6.FontChangeSize(libo_Menu.Font, 16)
    '        Else
    '            'QQQ Heike for testing
    '            MsgBox("TheModCntr not filled in form FormEFT")
    '        End If

    '        KeyPreview = True
    '        Exit Sub

    'MyStdErr:
    '        PrivLogErr(Err.Number, getLocationString("Form_Load"), Err.Description)
    '        Resume Next
    '    End Sub

    Protected Overridable Sub FormEFT_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed

        Dim bRet As Boolean

        On Error Resume Next

        theModCntr.HideVirtualKeyboard()
        bRet = theModCntr.unsetSubFormValues(Me)

    End Sub

    Protected Overridable Sub txt_Input_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        On Error Resume Next

        If Shift = VB6.ShiftConstants.CtrlMask And KeyCode = System.Windows.Forms.Keys.V Then
            blnCtrlVPressed = True
        Else
            blnCtrlVPressed = False
        End If
        'If (KeyCode = System.Windows.Forms.Keys.Return) Then cmdOK.Focus()
    End Sub

    'Protected Overridable Sub txt_Input_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
    '    Dim KeyAscii As Short = AscW(eventArgs.KeyChar)

    '    On Error Resume Next

    '    If lbl_Progress.Text = "GetDecimals" Then
    '        If Not blnCtrlVPressed Then
    '            'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
    '            KeyAscii = CheckKeyAscii4Nat(KeyAscii)
    '        Else
    '            'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
    '            KeyAscii = CheckKeyAscii4NatFromClipboard(KeyAscii)
    '        End If
    '    End If
    '    If lbl_Progress.Text = "GetAmount" Then
    '        If Not blnCtrlVPressed Then
    '            'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
    '            KeyAscii = CheckKeyAscii4Num(KeyAscii)
    '        Else
    '            'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
    '            KeyAscii = CheckKeyAscii4NumFromClipboard(KeyAscii)
    '        End If
    '    End If
    '    If KeyAscii = 0 Then
    '        eventArgs.Handled = True
    '    End If
    'End Sub

    Protected Overrides Sub WatchScanner_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        On Error Resume Next

        ' scanner input is checked here
        If theModCntr Is Nothing Then
            ' we are to early, ModCntr is not set
            Exit Sub
        End If
        If theModCntr.OposMSR_work_with = False Then
            ' No MSR activ
            Exit Sub
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object TheModCntr.thePosMSR.DataEventEnabled. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        theModCntr.thePosMSR.DataEventEnabled = True
        'UPGRADE_WARNING: Couldn't resolve default property of object TheModCntr.thePosMSR.bNewData. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If theModCntr.thePosMSR.bNewData = False Then
            ' no input received from the MSR
            ' wait longer
        Else
            ' automatic leave the mask
            ' TheModCntr.DialogActiv = False
            Me.Tag = True
            ' mark the tracks: has been read
            'UPGRADE_WARNING: Couldn't resolve default property of object TheModCntr.thePosMSR.bNewData. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            theModCntr.thePosMSR.bNewData = False
            RaiseEvent EFTDeviceMessage("Card swiped")
        End If

    End Sub
#End Region
#Region "Internal methods"
    Public Overridable Sub Payment_Succeeded()
        RaiseEvent EFTDeviceMessage("Succeeded")
    End Sub

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
#End Region

    Private Sub WatchScanner_Tick_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WatchScanner.Tick

        If ProgressBarRunning.Value >= ProgressBarRunning.Maximum Then
            ProgressBarRunning.Value = 0
        End If
        ProgressBarRunning.Value = ProgressBarRunning.Value + ProgressBarRunning.Step

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        On Error Resume Next

        If RBCloseEFT.Checked Then
            theModCntr.ModulNmbrExt = CInt(RBCloseEFT.Tag)
        ElseIf RBInitEFT.Checked Then
            theModCntr.ModulNmbrExt = CInt(RBInitEFT.Tag)
        ElseIf RBVoidLastTx.Checked Then
            theModCntr.ModulNmbrExt = CInt(RBVoidLastTx.Tag)
        ElseIf RBTotalizerEFT.CheckAlign Then
            theModCntr.ModulNmbrExt = CInt(RBTotalizerEFT.Tag)
        End If

        theModCntr.DialogActiv = False
    End Sub
End Class