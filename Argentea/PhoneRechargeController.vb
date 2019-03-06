Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing
Imports System.Collections.Generic

Public Class PhoneRechargeController

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

#Region "IPhoneRechargeActivationPreCheck"

    Public Function CheckPhoneRecharge(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "CheckPhoneRecharge"

        'Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As PhoneRechargeActivationParameters = New PhoneRechargeActivationParameters
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea CheckPhoneRecharge function")
            p.Controller = TheModCntr
            p.Transaction = taobj
            ' call in check mode
            Dim MyTaArticleRec As TPDotnet.Pos.TaArtSaleRec = CType(MyCurrentRecord, TPDotnet.Pos.TaArtSaleRec)
            Dim CSV As String = String.Empty
            Dim pCounter As Integer = p.GetNextPINCounter
            Dim szErrorMessage As String = String.Empty
            Dim szBarcode As String = MyTaArticleRec.szInputString
            Dim lValue As String = Math.Abs(CInt(MyTaArticleRec.dTaTotal * 100))
            response.ReturnCode = ArgenteaCOMObject.RichiestaPIN(pCounter, szBarcode, lValue, szErrorMessage, response.TransactionID)
            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                response.MessageOut = "KO" & ";;" & szErrorMessage
                response.SetProperty("szErrorMessage", szErrorMessage)
            Else
                response.MessageOut = "OK" & ";;" & szErrorMessage
            End If
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.PhoneRechargeCheck
            response.SetProperty("lPinCounter", pCounter)
            response.SetProperty("szBarcode", szBarcode)
            response.SetProperty("lAmount", lValue)

            paramArg.Copies = paramArg.PhoneRechargeCheckCopies
            paramArg.PrintWithinTA = paramArg.PhoneRechargeCheckPrintWithinTa


        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO

        Finally

        End Try
        Return response
    End Function

#End Region

#Region "IPhoneRechargeActivation"

    Public Function ActivatePhoneRecharge(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "ActivatePhoneRecharge"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim response As New ArgenteaResponse

        Try

            Dim MyExternalServiceRec As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(MyCurrentRecord, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
            Dim szTransactionID As String = MyExternalServiceRec.GetPropertybyName("szTransactionID")
            Dim lPinCounter As Integer = CInt(MyExternalServiceRec.GetPropertybyName("lPinCounter"))
            Dim lValue As Integer = MyExternalServiceRec.GetPropertybyName("lAmount")
            Dim szBarcode As String = MyExternalServiceRec.GetPropertybyName("szBarcode")
            Dim szPinID As String = String.Empty
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            response.ReturnCode = ArgenteaCOMObject.ConfermaPIN(szTransactionID, lPinCounter, lValue, szBarcode, szPinID, szMessageOut, szErrorMessage)

            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                ' Show an error for each gift card that cannot be definitely activated
                LOG_Debug(getLocationString(funcName), "Activation phone recharge  " & szBarcode & " returns error: " & szErrorMessage & " RetCode:" & response.ReturnCode)
                response.MessageOut = "KO" & ";" & szMessageOut & vbCrLf &
                        "ERRORE EMISSIONE RICARICA" & vbCrLf &
                        szErrorMessage & vbCrLf &
                        "Tran. ID: " & szTransactionID & vbCrLf &
                        "Barcode: " & szBarcode & vbCrLf &
                        "Value: " & Math.Round(lValue / 100, 2) & vbCrLf & vbCrLf & " " & ";" & szErrorMessage
            Else
                LOG_Debug(getLocationString(funcName), "Phone recharge number " & szBarcode & " successfuly activated")
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If

            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.PhoneRechargeActivation

            paramArg.Copies = paramArg.PhoneRechargeActivationCopies
            paramArg.PrintWithinTA = paramArg.PhoneRechargeActivationPrintWithinTa



        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO

        Finally

        End Try
        Return response
    End Function

#End Region

#Region "Overridable"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region


End Class
