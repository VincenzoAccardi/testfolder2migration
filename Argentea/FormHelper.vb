Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet
Imports TPDotnet.Pos

Public Class FormHelper

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

    Protected eftForm As FormEftInProgress

    ' open the eft form and disable the others function
    Public Overridable Sub OpenEftForm(ByRef TheModCntr As TPDotnet.Pos.ModCntr, _
                                       ByVal state As ArgenteaFormStates, _
                                       ByVal timeout As Integer)
        Try
            If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
                eftForm = TheModCntr.GetCustomizedForm(GetType(FormEftInProgress), STRETCH_TO_SMALL_WINDOW)
                eftForm.theModCntr = TheModCntr
                eftForm.LabelArgenteaWait.Visible = True
                eftForm.ProgressBarRunning.Visible = True
                eftForm.ProgressBarRunning.Step = 1
                eftForm.ProgressBarRunning.Maximum = timeout
                eftForm.WatchScanner.Start()
                TheModCntr.DialogFormName = eftForm.Text
                TheModCntr.SetFuncKeys(False)
                TheModCntr.DialogActiv = True
                eftForm.Show()
            Else
                ' simulated the dialog active 
                TheModCntr.DialogActiv = True
            End If
            
        Catch ex As Exception

        End Try
        
    End Sub

    ' close the eft form and enable the others function
    Public Overridable Sub CloseEftForm(ByRef TheModCntr As TPDotnet.Pos.ModCntr)
        TheModCntr.DialogActiv = False
        If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
            TheModCntr.DialogFormName = ""
            TheModCntr.SetFuncKeys(True)
            eftForm.Close()
            TheModCntr.EndForm()
        End If
    End Sub

    Public Sub update()

        eftForm.ProgressBarRunning.PerformStep()
        'If eftForm.ProgressBarRunning.Value = eftForm.ProgressBarRunning.Maximum Then
        '    eftForm.cmdCancel.Enabled = True
        '    eftForm.cmdCancel.Visible = True
        'End If

    End Sub
    Public Property ProgBarValue As Integer
        Get
            Return eftForm.ProgBarValue
        End Get
        Set(value As Integer)
            eftForm.ProgBarValue = value
        End Set
    End Property
    Public Property ProgBarMax As Integer
        Get
            Return eftForm.ProgBarMax
        End Get
        Set(value As Integer)
            eftForm.ProgBarMax = value
        End Set
    End Property
    Public Property EnableCancel As Boolean
        Get
            Return eftForm.EnableCancel
        End Get
        Set(value As Boolean)
            eftForm.EnableCancel = value
        End Set
    End Property

    Public Shared Sub ShowWaitScreen(ByRef TheModCntr As ModCntr, ByVal bClear As Boolean, ByRef form As System.Windows.Forms.Form, Optional ByVal customMsg As String = "", Optional ByVal addCustomMsg As String = "")

        Try
            If TheModCntr.bCalledFromWebService OrElse TheModCntr.bExternalDialog Then
                Exit Sub
            End If
            Dim i As Integer = -1
            Dim resolution As String = String.Empty
            Dim NameProcess As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "PROCESS_NAME").Trim
            Dim CrashBlock As String = System.IO.Path.Combine(getBinPath(), NameProcess + ".exe")
            If bClear Then
                Dim procs() As System.Diagnostics.Process
                procs = System.Diagnostics.Process.GetProcessesByName(NameProcess)
                For Each proc As System.Diagnostics.Process In procs
                    proc.Kill()
                Next
                TheModCntr.GUICntr.ThePosForm.BringToFront()
                TheModCntr.GUICntr.ThePosForm.Focus()
            Else
                Dim proc As New System.Diagnostics.Process()
                proc = System.Diagnostics.Process.Start(CrashBlock)
            End If

            'Try
            '    If TheModCntr.bCalledFromWebService OrElse TheModCntr.bExternalDialog Then
            '        Exit Sub
            '    End If

            '    If bClear Then

            '        If form IsNot Nothing Then
            '            form.Close()
            '            If Not form Is Nothing Then
            '                If TheModCntr IsNot Nothing Then
            '                    TheModCntr.EndForm()
            '                End If
            '            End If
            '            form = Nothing
            '        End If

            '    Else

            '        Dim msg As String = IIf(Not String.IsNullOrEmpty(customMsg), customMsg, getPosTxtNew((TheModCntr.contxt), "Message", TEXT_PLEASE_WAIT))
            '        msg &= customMsg

            '        form = TPMsg(msg, TEXT_PLEASE_WAIT, TheModCntr, "Message")
            '        Dim lx As Integer = (TheModCntr.GUICntr.ThePosForm.Width / 2) - (form.Width / 2)
            '        Dim ly As Integer = (TheModCntr.GUICntr.ThePosForm.Height / 2) - (form.Height / 2)
            '        form.Location = New System.Drawing.Point(lx, ly)

            '        System.Windows.Forms.Application.DoEvents()

            '    End If

            Exit Sub

        Catch ex As Exception
        Finally
        End Try
    End Sub

End Class
