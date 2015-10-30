Option Strict Off
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.Italy.Common.Pos

Public Class FormEFTError
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
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Info: Error in form FormEFTError.InitializeComponent")
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
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lbl_Info As System.Windows.Forms.Label
    Public WithEvents lbl_Error As System.Windows.Forms.Label
    Public WithEvents img_EFT2 As System.Windows.Forms.PictureBox
    Public WithEvents img_EFT1 As System.Windows.Forms.PictureBox
    Public WithEvents img_EFT3 As System.Windows.Forms.PictureBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents cmdOK As TPDotnet.GUIControls.WinButtonEx
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lbl_Info = New System.Windows.Forms.Label
        Me.lbl_Error = New System.Windows.Forms.Label
        Me.cmdOK = New TPDotnet.GUIControls.WinButtonEx
        Me.img_EFT2 = New System.Windows.Forms.PictureBox
        Me.img_EFT1 = New System.Windows.Forms.PictureBox
        Me.img_EFT3 = New System.Windows.Forms.PictureBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.img_EFT2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.img_EFT1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.img_EFT3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(2, 4)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2, 4, 2, 4)
        Me.PictureBox1.Size = New System.Drawing.Size(39, 19)
        Me.PictureBox1.Visible = False
        '
        'lbl_Info
        '
        Me.lbl_Info.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lbl_Info.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_Info.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbl_Info.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Info.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lbl_Info.Location = New System.Drawing.Point(55, 203)
        Me.lbl_Info.Name = "lbl_Info"
        Me.lbl_Info.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl_Info.Size = New System.Drawing.Size(217, 94)
        Me.lbl_Info.TabIndex = 1
        '
        'lbl_Error
        '
        Me.lbl_Error.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lbl_Error.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_Error.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbl_Error.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Error.ForeColor = System.Drawing.Color.Red
        Me.lbl_Error.Location = New System.Drawing.Point(55, 163)
        Me.lbl_Error.Name = "lbl_Error"
        Me.lbl_Error.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl_Error.Size = New System.Drawing.Size(217, 41)
        Me.lbl_Error.TabIndex = 0
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
        Me.cmdOK.Location = New System.Drawing.Point(85, 540)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(136, 48)
        Me.cmdOK.TabIndex = 8
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'img_EFT2
        '
        Me.img_EFT2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_EFT2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.img_EFT2.Cursor = System.Windows.Forms.Cursors.Default
        Me.img_EFT2.ErrorImage = Nothing
        Me.img_EFT2.InitialImage = Nothing
        Me.img_EFT2.Location = New System.Drawing.Point(151, 46)
        Me.img_EFT2.Name = "img_EFT2"
        Me.img_EFT2.Size = New System.Drawing.Size(121, 73)
        Me.img_EFT2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.img_EFT2.TabIndex = 11
        Me.img_EFT2.TabStop = False
        '
        'img_EFT1
        '
        Me.img_EFT1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_EFT1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.img_EFT1.Cursor = System.Windows.Forms.Cursors.Default
        Me.img_EFT1.ErrorImage = Nothing
        Me.img_EFT1.InitialImage = Nothing
        Me.img_EFT1.Location = New System.Drawing.Point(56, 46)
        Me.img_EFT1.Name = "img_EFT1"
        Me.img_EFT1.Size = New System.Drawing.Size(89, 73)
        Me.img_EFT1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.img_EFT1.TabIndex = 12
        Me.img_EFT1.TabStop = False
        '
        'img_EFT3
        '
        Me.img_EFT3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_EFT3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.img_EFT3.Cursor = System.Windows.Forms.Cursors.Default
        Me.img_EFT3.Enabled = False
        Me.img_EFT3.ErrorImage = Nothing
        Me.img_EFT3.InitialImage = Nothing
        Me.img_EFT3.Location = New System.Drawing.Point(56, 46)
        Me.img_EFT3.Name = "img_EFT3"
        Me.img_EFT3.Size = New System.Drawing.Size(217, 105)
        Me.img_EFT3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.img_EFT3.TabIndex = 13
        Me.img_EFT3.TabStop = False
        Me.img_EFT3.Visible = False
        '
        'FormEFTError
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(340, 600)
        Me.ControlBox = False
        Me.Controls.Add(Me.img_EFT2)
        Me.Controls.Add(Me.img_EFT1)
        Me.Controls.Add(Me.img_EFT3)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lbl_Info)
        Me.Controls.Add(Me.lbl_Error)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(16, 0)
        Me.Name = "FormEFTError"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "FormEFTError"
        Me.Controls.SetChildIndex(Me.lbl_Error, 0)
        Me.Controls.SetChildIndex(Me.lbl_Info, 0)
        Me.Controls.SetChildIndex(Me.cmdOK, 0)
        Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.img_EFT3, 0)
        Me.Controls.SetChildIndex(Me.img_EFT1, 0)
        Me.Controls.SetChildIndex(Me.img_EFT2, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.img_EFT2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.img_EFT1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.img_EFT3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FormEFTError
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FormEFTError
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FormEFTError()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As FormEFTError)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
#Region "Control-events"
    Protected Overridable Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Dim strTemp As String

        On Error Resume Next

        Tag = True

        theModCntr.DialogActiv = False
    End Sub

    Protected Overridable Sub FormEFTError_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        If Me.DesignMode Then
            ' We are in design mode. Important when this form is used as inherited form, 
            ' then the following code should not be passed through
            Exit Sub
        End If

        On Error GoTo MyStdErr

        If Not theModCntr Is Nothing Then
            GotFocusEvent4VirtualKeyboard((Me.theModCntr), Me.Handle.ToInt32, True)
        Else
            LOG_Info(getLocationString("FormEFTError_Load"), "theModCntr not filled")
        End If

        KeyPreview = True

        Exit Sub

MyStdErr:
        LOG_Error(getLocationString("FormEFTError_Load"), Err.Description)
        Resume Next
    End Sub

    Protected Overridable Sub FormEFTError_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        Dim bRet As Boolean

        On Error Resume Next

        bRet = theModCntr.unsetSubFormValues(Me)

    End Sub
#End Region
#Region "Internal methods"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
#End Region
End Class