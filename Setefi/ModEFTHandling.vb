Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class ModEFTHandling
    Inherits ModBase

#Region "Documentation"
    ' ********** ********** ********** **********
    ' ModArgentea
    ' ---------- ---------- ---------- ----------
    ' Input form for advanced Agentea EMV operations
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Overridable"

    Protected myForm As FormEFTHandling = Nothing
    Protected LastOkAmount As Integer = 0
    Protected LastOkNumber As Integer = 0
    Protected NextNumber As Integer = 0

    ' open the eft form and disable the others function
    Public Overridable Sub OpenEftForm(ByRef TheModCntr As TPDotnet.Pos.ModCntr)
        myForm = TheModCntr.GetCustomizedForm(GetType(FormEFTHandling), STRETCH_TO_SMALL_WINDOW)
        myForm.theModCntr = TheModCntr
        myForm.LabelEFTLastOKAmountToken.Text = Format(CType(LastOkAmount, Double) / 100, "#0.00")
        myForm.LabelEFTLastOKOpNumToken.Text = LastOkNumber.ToString
        myForm.LabelEFTNextOpNumToken.Text = NextNumber.ToString
        myForm.GroupBox1.Visible = True
        myForm.RBCloseEFT.Visible = True
        myForm.RBInitEFT.Visible = True
        myForm.RBVoidLastTx.Visible = True
        myForm.RBTotalizerEFT.Visible = True
        myForm.cmdOK.Visible = True
        myForm.cmdCancel.Visible = True
        myForm.LabelEFTLastOKPayment.Visible = True
        myForm.LabelEFTLastOKAmount.Visible = True
        myForm.LabelEFTLastOKAmountToken.Visible = True
        myForm.LabelEFTLastOKOpNum.Visible = True
        myForm.LabelEFTLastOKOpNumToken.Visible = True
        'myForm.LabelEFTNextOpNum.Visible = True
        'myForm.LabelEFTNextOpNumToken.Visible = True
        TheModCntr.DialogFormName = myForm.Text
        TheModCntr.SetFuncKeys(False)
        TheModCntr.DialogActiv = True
        myForm.Show()

    End Sub

    ' close the eft form and enable the others function
    Public Overridable Sub CloseEftForm(ByRef TheModCntr As TPDotnet.Pos.ModCntr)
        TheModCntr.DialogActiv = False
        TheModCntr.DialogFormName = ""
        TheModCntr.SetFuncKeys(True)
        myForm.LabelEFTWait.Visible = False
        myForm.Close()
        TheModCntr.EndForm()
    End Sub

    Public Overrides Function ModBase_run(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Short

        Dim SelectedModulNmbrExt As Short = 0
        Dim Ret As Short = -1

        Try

            LOG_FuncStart("ModBase_run", "starting")

            'set the ModCtrl to thePosEFT
            'set the taobj to thePosEFT
            TheModCntr.thePosEFT.Initialize(taobj, TheModCntr)

            ' if ModulNmbrExt is 0 shown available options via ui
            If TheModCntr.ModulNmbrExt = 0 Then

                LastOkAmount = TheModCntr.thePosEFT.DoServiceRequest(taobj, "GetLastPaymentTransactionAmount")
                LastOkNumber = TheModCntr.thePosEFT.DoServiceRequest(taobj, "GetLastPaymentTransactionNumber")
                NextNumber = TheModCntr.thePosEFT.DoServiceRequest(taobj, "GetOperationNumber")

                OpenEftForm(TheModCntr)
                Do While TheModCntr.DialogActiv = True
                    Debug.Sleep(100)
                    System.Windows.Forms.Application.DoEvents()
                Loop
                CloseEftForm(TheModCntr)

            End If

            Select Case TheModCntr.ModulNmbrExt

                Case 0
                    ' no selection made
                    Exit Select

                Case 1
                    ' void last transaction
                    Ret = TheModCntr.thePosEFT.VoidEftDevice(taobj, Nothing)
                    Exit Select

                Case 2
                    ' close EFT and print totals
                    TheModCntr.thePosEFT.SilentLogoff(taobj)
                    Exit Select

                Case 3
                    ' init EFT
                    TheModCntr.thePosEFT.Initialize(taobj, TheModCntr)
                    Exit Select

                Case 4
                    ' print totals
                    TheModCntr.thePosEFT.Logoff(taobj)
                    Exit Select

            End Select

            'If TheModCntr.ModulNmbrExt <> 0 Then
            '    taobj.bTAtoFile = False
            '    taobj.bPrintReceipt = False
            '    taobj.bDelete = True ' ok , we will delete this TA
            '    taobj.TARefresh()
            'End If

            Exit Function

        Catch ex As Exception
            LOG_Error("ModBase_run", ex)
        Finally
            LOG_FuncExit("ModBase_run", "exiting")
        End Try

    End Function

#End Region

End Class
