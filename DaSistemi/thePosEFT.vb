Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.Italy.Common.Pos

Public Class thePosEFT

#Region "Documentation"
    ' ********** ********** ********** **********
    ' E F T - DaSistemi protocol implementation
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region
#Region "Global member variables"

    Protected Structure PropertyType
        Dim m_ModCntrl As ModCntr

        ' collection of the customized ta Objects, will be given from outside
        Dim m_colObjects As Collections.Hashtable
    End Structure
    Protected Property_Renamed As PropertyType
#End Region
#Region "Interface methods"

    Dim myForm As FormEFT = Nothing

    ' open the eft form and disable the others function
    Public Overridable Sub OpenEftForm()
        myForm = Me.TheModCntr.GetCustomizedForm(GetType(FormEFT), STRETCH_TO_SMALL_WINDOW)
        myForm.theModCntr = Me.TheModCntr
        TheModCntr.DialogFormName = myForm.Text
        TheModCntr.SetFuncKeys(False)
        TheModCntr.DialogActiv = True
        myForm.Show()
    End Sub

    Public Overridable Sub CloseEftForm()
        TheModCntr.DialogActiv = False
        TheModCntr.DialogFormName = ""
        TheModCntr.SetFuncKeys(True)
        myForm.Close()
        TheModCntr.EndForm()
    End Sub

    ' ask if the transaction was closed directly on the eft
    Public Overridable Function AskForManualPayment() As Short
        Dim iRet As Short
        iRet = TPMsgBoxRet(PosDef.TARMessageTypes.TPQUESTION, getPosTxtNew((TheModCntr.contxt), "UserMessage", TXT_EFT_MANUAL_PAYMENT), MsgBoxStyle.YesNo, TheModCntr, 777, "Eseguito pagamento manuale?")
        If iRet = MsgBoxResult.No Then
            Return 110
        Else
            Return 0
        End If
    End Function

    Private WithEvents DA As DaSistemi = New DaSistemi
    Public Overridable Function PostEFTDevice(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec) As Short

        Dim szRequestType As String = ""

        PostEFTDevice = 110 ' set a bad return code as default

        Try
            OpenEftForm()

            If DA.PayDaSistemi(MyTaMediaRec.dTaPaidTotal, Me.TheModCntr) = 0 Then
                Do While TheModCntr.DialogActiv = True
                    Sleep(100)
                    System.Windows.Forms.Application.DoEvents()
                Loop
            Else
                CloseEftForm()
                TPMsgBox(PosDef.TARMessageTypes.TPERROR, DA.ErrorMessage, 0, TheModCntr, DA.ErrorMessage)
                PostEFTDevice = AskForManualPayment()
                'Exit Function
            End If

            If DA.State = DaSistemi.States.DA_SISTEMI_ERROR Then
                CloseEftForm()
                TPMsgBox(PosDef.TARMessageTypes.TPERROR, DA.ErrorMessage, 0, TheModCntr, DA.ErrorMessage)
                PostEFTDevice = AskForManualPayment()
                'Exit Function
            ElseIf DA.State = DaSistemi.States.DA_SISTEMI_SUCCESS Then
                CloseEftForm()
                PostEFTDevice = 0
                TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION, getPosTxtNew((TheModCntr.contxt), "UserMessage", TXT_EFT_PAYMENT_OK), 0, TheModCntr, "Message")
            ElseIf DA.State = DaSistemi.States.DA_SISTEMI_IN_PROGRESS Then
                DA.AbortPayment() ' abort the eft thread
                CloseEftForm()
                TPMsgBox(PosDef.TARMessageTypes.TPERROR, getPosTxtNew((TheModCntr.contxt), "UserMessage", TXT_EFT_PAYMENT_ABORT), 0, TheModCntr, "Message")
                PostEFTDevice = AskForManualPayment()
                'Exit Function
            End If

        Catch ex As Exception

        Finally

        End Try
    End Function

#End Region

#Region "Italy"

    Private Delegate Sub InfoUpdateEventHandler(ByVal message As String)
    Public Sub WriteInfo(ByVal value As String)

        Try
            If (myForm.lbl_Info.InvokeRequired) Then
                myForm.lbl_Info.Invoke(New InfoUpdateEventHandler(AddressOf WriteInfo), value)
            Else
                myForm.lbl_Info.Text = value
            End If
        Catch ex As ObjectDisposedException
            Console.WriteLine(ex.Message)
        End Try

    End Sub


    Public Sub DaSistemiStatusChanged(ByVal state As Integer, ByVal message As String) Handles DA.DaSistemiStatusChanged

        If state = DaSistemi.States.DA_SISTEMI_ERROR Then
            WriteInfo(String.Format("{0}", message))

        ElseIf state = DaSistemi.States.DA_SISTEMI_IN_PROGRESS Then
            WriteInfo(String.Format("{0}", message))

        ElseIf state = DaSistemi.States.DA_SISTEMI_SUCCESS Then
            WriteInfo(String.Format("{0}", message))
        End If

    End Sub

    Private Delegate Sub RemainingSecondsUpdateEventHandler(ByVal seconds As Integer)
    Public Sub WriteRemainingSeconds(ByVal s As Integer)
        Try
            If myForm.timeout_lbl.InvokeRequired Then
                myForm.timeout_lbl.Invoke(New RemainingSecondsUpdateEventHandler(AddressOf WriteRemainingSeconds), s)
            Else
                myForm.timeout_lbl.Text = String.Format("Tempo rimanente per l'operazione in corso : {0}", s)
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Public Sub DaSistemiRemainindSecondsChanged(ByVal s As Integer) Handles DA.DaSistemiRemainingSecond
        WriteRemainingSeconds(s)
    End Sub

#End Region

#Region "Properties"

    Public Overridable Property TheModCntr() As ModCntr
        Get
            TheModCntr = Property_Renamed.m_ModCntrl
        End Get
        Set(ByVal Value As ModCntr)
            Property_Renamed.m_ModCntrl = Value
        End Set
    End Property

#End Region

End Class