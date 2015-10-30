Option Strict Off
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.Pos.UtilAllPosModuls
Imports TPDotnet.Pos.UtilAllPosModulsPart2

Public Class FormEFT
    Inherits TPDotnet.Pos.FormBase

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
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
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
    Public WithEvents lbl_Info As System.Windows.Forms.Label
    Public WithEvents lbl_Progress As System.Windows.Forms.Label
    Friend WithEvents timeout_lbl As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents cmdCancel As TPDotnet.GUIControls.WinButtonEx
    Private Sub InitializeComponent()
        Me.lbl_Info = New System.Windows.Forms.Label
        Me.lbl_Progress = New System.Windows.Forms.Label
        Me.cmdCancel = New TPDotnet.GUIControls.WinButtonEx
        Me.timeout_lbl = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Visible = False
        '
        'lbl_Info
        '
        Me.lbl_Info.BackColor = System.Drawing.Color.White
        Me.lbl_Info.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_Info.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbl_Info.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lbl_Info.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Info.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lbl_Info.Location = New System.Drawing.Point(26, 202)
        Me.lbl_Info.Name = "lbl_Info"
        Me.lbl_Info.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl_Info.Size = New System.Drawing.Size(289, 73)
        Me.lbl_Info.TabIndex = 3
        '
        'lbl_Progress
        '
        Me.lbl_Progress.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lbl_Progress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_Progress.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbl_Progress.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Progress.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lbl_Progress.Location = New System.Drawing.Point(56, 216)
        Me.lbl_Progress.Name = "lbl_Progress"
        Me.lbl_Progress.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl_Progress.Size = New System.Drawing.Size(217, 25)
        Me.lbl_Progress.TabIndex = 4
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
        Me.cmdCancel.Location = New System.Drawing.Point(26, 540)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(289, 48)
        Me.cmdCancel.TabIndex = 12
        Me.cmdCancel.Text = "Annulla"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'timeout_lbl
        '
        Me.timeout_lbl.AutoSize = True
        Me.timeout_lbl.Location = New System.Drawing.Point(23, 287)
        Me.timeout_lbl.Name = "timeout_lbl"
        Me.timeout_lbl.Size = New System.Drawing.Size(217, 14)
        Me.timeout_lbl.TabIndex = 13
        Me.timeout_lbl.Text = "Tempo rimanente per l'operazione in corso :"
        '
        'FormEFT
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.timeout_lbl)
        Me.Controls.Add(Me.lbl_Info)
        Me.Controls.Add(Me.lbl_Progress)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "FormEFT"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "FormEFT"
        Me.Controls.SetChildIndex(Me.lbl_Progress, 0)
        Me.Controls.SetChildIndex(Me.lbl_Info, 0)
        Me.Controls.SetChildIndex(Me.timeout_lbl, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FormEFT
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FormEFT
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FormEFT()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As FormEFT)
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

    Protected Overridable Sub FormEFT_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim bRet As Boolean

        On Error GoTo MyStdErr

        If Not theModCntr Is Nothing Then
            GotFocusEvent4VirtualKeyboard((Me.theModCntr), Me.Handle.ToInt32, True)
            ' set top,left,width and height first for reducing flicker
            'Call TheModCntr.setSubFormPosition(Me)
            'Call LoadFormEventNew(theModCntr, Me, actRectOriginal, True)
            'bRet = TheModCntr.setSubFormValues(Me, 1)
            'libo_Menu.Font = VB6.FontChangeName(libo_Menu.Font, "Arial")
            'libo_Menu.Font = VB6.FontChangeSize(libo_Menu.Font, 16)
        Else
            'QQQ Heike for testing
            MsgBox("TheModCntr not filled in form FormEFT")
        End If

        KeyPreview = True
        Exit Sub

MyStdErr:
        LOG_Error(getLocationString("Form_Load"), Err.Description)
        Resume Next
    End Sub

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

    Protected Overridable Sub txt_Input_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = AscW(eventArgs.KeyChar)

        On Error Resume Next

        If lbl_Progress.Text = "GetDecimals" Then
            If Not blnCtrlVPressed Then
                'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
                KeyAscii = CheckKeyAscii4Nat(KeyAscii)
            Else
                'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
                KeyAscii = CheckKeyAscii4NatFromClipboard(KeyAscii)
            End If
        End If
        If lbl_Progress.Text = "GetAmount" Then
            If Not blnCtrlVPressed Then
                'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
                KeyAscii = CheckKeyAscii4Num(KeyAscii)
            Else
                'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1058"'
                KeyAscii = CheckKeyAscii4NumFromClipboard(KeyAscii)
            End If
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

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

End Class