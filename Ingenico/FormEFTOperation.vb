Option Strict Off
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.Pos.UtilAllPosModuls
Imports TPDotnet.Pos.UtilAllPosModulsPart2

Public Class FormEFTOperation
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
    Friend WithEvents timeout_lbl As System.Windows.Forms.Label
    Friend WithEvents ListboxLog As TPDotnet.GUIControls.WinListboxEx
    Friend WithEvents ButtonListboxLogUp As TPDotnet.GUIControls.WinButtonEx
    Friend WithEvents ButtonListboxLogDown As TPDotnet.GUIControls.WinButtonEx
    Friend WithEvents Waiter1 As EmaControlLib.Waiter
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents cmdCancel As TPDotnet.GUIControls.WinButtonEx
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormEFTOperation))
        Me.cmdCancel = New TPDotnet.GUIControls.WinButtonEx
        Me.timeout_lbl = New System.Windows.Forms.Label
        Me.ListboxLog = New TPDotnet.GUIControls.WinListboxEx
        Me.ButtonListboxLogUp = New TPDotnet.GUIControls.WinButtonEx
        Me.ButtonListboxLogDown = New TPDotnet.GUIControls.WinButtonEx
        Me.Waiter1 = New EmaControlLib.Waiter
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(2, 4)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2, 4, 2, 4)
        Me.PictureBox1.Size = New System.Drawing.Size(39, 19)
        Me.PictureBox1.Visible = False
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
        Me.cmdCancel.ForeColor = System.Drawing.Color.Black
        Me.cmdCancel.Location = New System.Drawing.Point(26, 535)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(242, 43)
        Me.cmdCancel.TabIndex = 12
        Me.cmdCancel.Text = "Annulla"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'timeout_lbl
        '
        Me.timeout_lbl.AutoSize = True
        Me.timeout_lbl.Location = New System.Drawing.Point(23, 497)
        Me.timeout_lbl.Name = "timeout_lbl"
        Me.timeout_lbl.Size = New System.Drawing.Size(217, 14)
        Me.timeout_lbl.TabIndex = 13
        Me.timeout_lbl.Text = "Tempo rimanente per l'operazione in corso :"
        '
        'ListboxLog
        '
        Me.ListboxLog.AngleToleranceForSlideInDegree = 40
        Me.ListboxLog.BlockMoveOnEdges = True
        Me.ListboxLog.BottomBorderWidth = 2
        Me.ListboxLog.ColumnsSeparators = New Integer(-1) {}
        Me.ListboxLog.DataSource = CType(resources.GetObject("ListboxLog.DataSource"), System.Collections.ArrayList)
        Me.ListboxLog.FullRowSelection = True
        Me.ListboxLog.Items = CType(resources.GetObject("ListboxLog.Items"), System.Collections.ArrayList)
        Me.ListboxLog.LeftBorderWidth = 2
        Me.ListboxLog.Location = New System.Drawing.Point(26, 145)
        Me.ListboxLog.MinimalLengthForSlideInPx = 10
        Me.ListboxLog.Name = "ListboxLog"
        Me.ListboxLog.NormalTextColor = System.Drawing.Color.Black
        Me.ListboxLog.RightBorderWidth = 2
        Me.ListboxLog.RowHeight = 25
        Me.ListboxLog.RowSeparatorColor = System.Drawing.Color.Gray
        Me.ListboxLog.ScrollDirection = TPDotnet.GUIControls.WinListboxEx.EnumScrollDirection.NoScroll
        Me.ListboxLog.ScrollFrameRatePerSec = 40
        Me.ListboxLog.ScrollSlowDownInPxPerSec = 60
        Me.ListboxLog.ScrollSpeedInPxPerMs = 0
        Me.ListboxLog.ScrollStartSensibilityInMs = 3
        Me.ListboxLog.SelectedBackgroundColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.ListboxLog.SelectedColumnIndex = -1
        Me.ListboxLog.SelectedRowIndex = -1
        Me.ListboxLog.SelectedTextColor = System.Drawing.Color.Black
        Me.ListboxLog.SelectionMargeToleranceInPx = 4
        Me.ListboxLog.SelectRowOnSlide = True
        Me.ListboxLog.Size = New System.Drawing.Size(242, 219)
        Me.ListboxLog.StartRowSelect = -1
        Me.ListboxLog.TabIndex = 16
        Me.ListboxLog.Text = "WinListboxEx1"
        Me.ListboxLog.TopBorderWidth = 5
        Me.ListboxLog.WaitTimeBeforeDisplayPreselectionInMs = 250
        '
        'ButtonListboxLogUp
        '
        Me.ButtonListboxLogUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonListboxLogUp.DeactiveTextColor = System.Drawing.SystemColors.GrayText
        Me.ButtonListboxLogUp.FlatAppearance.BorderSize = 0
        Me.ButtonListboxLogUp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonListboxLogUp.Location = New System.Drawing.Point(26, 113)
        Me.ButtonListboxLogUp.Name = "ButtonListboxLogUp"
        Me.ButtonListboxLogUp.Size = New System.Drawing.Size(242, 26)
        Me.ButtonListboxLogUp.TabIndex = 17
        Me.ButtonListboxLogUp.UseVisualStyleBackColor = True
        '
        'ButtonListboxLogDown
        '
        Me.ButtonListboxLogDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonListboxLogDown.DeactiveTextColor = System.Drawing.SystemColors.GrayText
        Me.ButtonListboxLogDown.FlatAppearance.BorderSize = 0
        Me.ButtonListboxLogDown.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonListboxLogDown.Location = New System.Drawing.Point(26, 370)
        Me.ButtonListboxLogDown.Name = "ButtonListboxLogDown"
        Me.ButtonListboxLogDown.Size = New System.Drawing.Size(242, 26)
        Me.ButtonListboxLogDown.TabIndex = 18
        Me.ButtonListboxLogDown.UseVisualStyleBackColor = True
        '
        'Waiter1
        '
        Me.Waiter1.BackColor = System.Drawing.Color.Transparent
        Me.Waiter1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Waiter1.Location = New System.Drawing.Point(115, 419)
        Me.Waiter1.Margin = New System.Windows.Forms.Padding(0)
        Me.Waiter1.Name = "Waiter1"
        Me.Waiter1.Size = New System.Drawing.Size(62, 62)
        Me.Waiter1.TabIndex = 19
        Me.Waiter1.Visible = False
        '
        'FormEFTOperation
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.ControlBox = False
        Me.Controls.Add(Me.Waiter1)
        Me.Controls.Add(Me.ButtonListboxLogDown)
        Me.Controls.Add(Me.ButtonListboxLogUp)
        Me.Controls.Add(Me.timeout_lbl)
        Me.Controls.Add(Me.ListboxLog)
        Me.Controls.Add(Me.cmdCancel)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!)
        Me.Name = "FormEFTOperation"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "FormEFT"
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.ListboxLog, 0)
        Me.Controls.SetChildIndex(Me.timeout_lbl, 0)
        Me.Controls.SetChildIndex(Me.ButtonListboxLogUp, 0)
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.ButtonListboxLogDown, 0)
        Me.Controls.SetChildIndex(Me.Waiter1, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FormEFTOperation
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FormEFTOperation
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FormEFTOperation()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As FormEFTOperation)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
#Region "Variables"
    Public Event EFTDeviceMessage(ByRef szMessage As String)
    Public Event EFTDeviceCheckVoid(ByRef bIsVoid As Boolean)

    Public iListIndex As Short

#End Region
#Region "Control-events"

    Protected Overridable Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click

        On Error Resume Next
        iListIndex = -1
        RaiseEvent EFTDeviceMessage("Abort")
        'Tag = ""
        Tag = False
        'theModCntr.DialogActiv = False
        LOG_Error(getLocationString("cmdCancel_Click"), "Cancel operation button clicked!")
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

    Private Sub ButtonListboxLogUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonListboxLogUp.Click
        ListboxLog.Select()
        SendKeys.SendWait("{UP}")
        Application.DoEvents()
    End Sub

    Private Sub ButtonListboxLogDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonListboxLogDown.Click
        ListboxLog.Select()
        SendKeys.SendWait("{DOWN}")
        Application.DoEvents()
    End Sub
End Class