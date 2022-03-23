Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing

Public Class GiftCardController

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

#Region "Instance related functions"

    Public Sub New()
        Parameters = New ArgenteaParameters()
    End Sub

#End Region

#Region "Parameters"

    Protected Parameters As ArgenteaParameters

#End Region

#Region "Public functions"

#End Region

#Region "IGiftCardActivationPreCheck"

    Public Function CheckGiftCard(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "CheckGiftCard"
        Dim response As New ArgenteaResponse


        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea CheckGiftCard function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            Dim MyTaArtSaleRec As TPDotnet.Pos.TaArtSaleRec = CType(MyCurrentRecord, TPDotnet.Pos.TaArtSaleRec)
            'p.LoadCommonFunctionParameter(Parameters)
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            Dim szBarcode As String = MyCurrentRecord.GetPropertybyName("szITGiftCardEAN")
            Dim lAmount As Integer = CInt(MyTaArtSaleRec.dTaTotal * 100)
            response.ReturnCode = ArgenteaCOMObject.GiftCardActivation(lAmount, 1, szBarcode, szBarcode, szMessageOut, szErrorMessage)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)


            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                response.MessageOut = "KO" & ";" & szMessageOut & ";" & szErrorMessage
                response.SetProperty("szErrorMessage", szErrorMessage)
            Else
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If
            response.SetProperty("szBarcode", szBarcode)
            response.SetProperty("lAmount", lAmount)
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.GiftCardActivationPreCheck

            paramArg.Copies = paramArg.GiftCardActivationCheckCopies
            paramArg.PrintWithinTA = False


        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally

        End Try
        Return response
    End Function

#End Region

#Region "IGiftCardActivation"
    Public Function ActivateGiftCard(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "ActivateGiftCard"
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IGiftCardActivation function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            Dim MyExternalServiceRec As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(MyCurrentRecord, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
            Dim lAmount As Integer = MyExternalServiceRec.GetPropertybyName("lAmount")
            Dim szBarcode As String = MyExternalServiceRec.GetPropertybyName("szBarcode")
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            response.ReturnCode = ArgenteaCOMObject.GiftCardActivation(lAmount, 0, szBarcode, szBarcode, szMessageOut, szErrorMessage)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)

            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                ' Show an error for each gift card that cannot be definitely activated
                LOG_Error(getLocationString(funcName), "Activation for giftcard  " & szBarcode & " returns error: " & szErrorMessage)
                response.MessageOut = "KO" & ";" & szMessageOut & vbCrLf &
                                                                "!!!ERRORE DI ATTIVAZIONE!!!" & vbCrLf &
                                                                 szErrorMessage & vbCrLf &
                                                                "Giftcard Serial: " & szBarcode & vbCrLf &
                                                                "Value: " & Math.Round(lAmount / 100, 2) & vbCrLf & vbCrLf & " " & ";" & szErrorMessage
                response.SetProperty("szErrorMessage", szErrorMessage)
            Else
                LOG_Debug(getLocationString(funcName), "Gift card number " & szBarcode & " successfuly activated")
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.GiftCardActivation

            paramArg.Copies = paramArg.GiftCardActivationCopies
            paramArg.PrintWithinTA = paramArg.GiftCardActivationPrintWtihinTa

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally

        End Try
        Return response
    End Function


#End Region

#Region "IGiftCardBalanceInquiry"

    'Public Function GiftCardBalanceInquiry(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardBalanceInquiry.GiftCardBalanceInquiry
    '    GiftCardBalanceInquiry = IGiftCardReturnCode.KO
    '    Dim funcName As String = "GiftCardBalanceInquiry"

    '    Dim frm As System.Windows.Forms.Form = Nothing
    '    Dim p As BalanceParameters = New BalanceParameters
    '    Dim CSV As String = String.Empty
    '    Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

    '    Try
    '        LOG_Debug(getLocationString(funcName), "We are entered in Argentea IGiftCardBalanceInquiry function")
    '        ' collect the input parameters
    '        LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
    '        ' collect the input parameters
    '        p.LoadCommonFunctionParameter(parameters)

    '        Me.Parameters.LoadParametersByReflection(p.Controller)
    '        p.GiftCardBalanceLineIdentifier = Me.Parameters.GiftCardBalanceLineIdentifier
    '        Me.Parameters.GiftCardBalanceInternalInquiry = p.GiftCardBalanceInternalInquiry
    '        ' check the balance
    '        FormHelper.ShowWaitScreen(p.Controller, False, frm)

    '        ArgenteaCOMObject = Nothing
    '        ArgenteaCOMObject = New ARGLIB.argpay()
    '        retCode = ArgenteaCOMObject.GiftCardBalance(p.Barcode, p.ErrorMessage, p.MessageOut)
    '        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". Giftcard: " & p.Barcode & ". Error: " & p.ErrorMessage & ". Output: " & p.MessageOut)

    '        If retCode <> ArgenteaFunctionsReturnCode.OK Then
    '            LOG_Error(getLocationString(funcName), "Balance for giftcard " & p.Barcode & " returns error: " & p.ErrorMessage & ". The message output is: " & p.MessageOut)
    '            CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
    '        Else
    '            GiftCardBalanceInquiry = IGiftCardReturnCode.OK
    '            CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
    '        End If

    '        Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
    '        objTPTAHelperArgentea.HandleReturnString(p.Transaction,
    '                                                 p.Controller,
    '                                                 CSV,
    '                                                 InternalArgenteaFunctionTypes.GiftCardBalance,
    '                                                 Me.Parameters)

    '        If String.IsNullOrEmpty(p.MessageOut) Then
    '            ' strange situation: function returns ok but without receipt
    '            ' log log log
    '            Exit Function
    '        End If

    '        ' copy back the values for value type fields
    '        If parameters.ContainsKey("Value") Then parameters("Value") = p.Value
    '        If parameters.ContainsKey("Receipt") Then parameters("Receipt") = p.Receipt

    '        GiftCardBalanceInquiry = IGiftCardReturnCode.OK

    '    Catch ex As Exception
    '        LOG_Error(getLocationString(funcName), ex.Message)
    '    Finally
    '        FormHelper.ShowWaitScreen(p.Controller, True, frm)
    '        ShowError(p)
    '    End Try
    'End Function


    Public Function GiftCardBalanceInquiry(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "GiftCardBalanceInquiry"
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IGiftCardBalanceInquiry function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            ' collect the input parameters
            'p.LoadCommonFunctionParameter(Parameters)

            Me.Parameters.LoadParametersByReflection(TheModCntr)
            ' check the balance
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            Dim szBarcode As String = MyCurrentRecord.GetPropertybyName("szITGiftCardEAN")

            response.ReturnCode = ArgenteaCOMObject.GiftCardBalance(szBarcode, szErrorMessage, szMessageOut)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)


            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                LOG_Error(getLocationString(funcName), "Balance for giftcard " & szBarcode & " returns error: " & szErrorMessage & ". The message output is: " & szMessageOut)
                response.SetProperty("szErrorMessage", szErrorMessage)

                response.MessageOut = "KO" & ";" & szMessageOut & ";" & szErrorMessage
            Else
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If

            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.GiftCardBalance

            paramArg.Copies = paramArg.GiftCardActivationCopies
            paramArg.PrintWithinTA = paramArg.GiftCardActivationPrintWtihinTa
            paramArg.GiftCardBalanceLineIdentifier = Me.Parameters.GiftCardBalanceLineIdentifier
            paramArg.GiftCardBalanceInternalInquiry = IIf((MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA), True, False)
            If paramArg.GiftCardBalanceInternalInquiry Then
                paramArg.Copies = 0
                paramArg.PrintWithinTA = False
                response.SetProperty("lAmount", GetValue(szMessageOut, paramArg.GiftCardBalanceLineIdentifier).ToString())
            End If
        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally

        End Try
        Return response
    End Function
    Private Function GetValue(ByVal szMessageOut As String, ByVal szGiftCardBalanceLineIdentifier As String) As Decimal
        If Not String.IsNullOrEmpty(szMessageOut) Then
            Dim linees As String() = szMessageOut.Split(vbCrLf.ToCharArray)
            For Each l As String In linees
                If l.ToUpper.StartsWith(szGiftCardBalanceLineIdentifier.ToUpper) Then
                    GetValue = Convert.ToDecimal(l.ToUpper.Substring(szGiftCardBalanceLineIdentifier.Length).Trim)
                    Exit For
                End If
            Next l
        End If
    End Function

#End Region

#Region "IGiftCardRedeemPreCheck"

    'Public Function CheckRedeemGiftCard(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardRedeemPreCheck.CheckRedeemGiftCard
    '    CheckRedeemGiftCard = IGiftCardReturnCode.KO
    '    Dim funcName As String = "CheckRedeemGiftCard"
    '    LOG_FuncStart(funcName)
    '    Dim frm As System.Windows.Forms.Form = Nothing
    '    Dim p As GiftCardRedeemParameters = New GiftCardRedeemParameters
    '    Dim CSV As String = String.Empty
    '    Dim taArgenteaEMVRec As TaExternalServiceRec = Nothing
    '    Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

    '    Try
    '        p.LoadCommonFunctionParameter(parameters)

    '        FormHelper.ShowWaitScreen(p.Controller, False, frm)

    '        ArgenteaCOMObject = Nothing
    '        ArgenteaCOMObject = New ARGLIB.argpay()
    '        LOG_Error(funcName, "Argentea dll GiftCardRedeem  function")
    '        LOG_Error(funcName, "Input : IntValue=" + p.IntValue.ToString)
    '        LOG_Error(funcName, "Input : Barcode=" + p.Barcode.ToString)
    '        LOG_Error(funcName, "Input : TransactionID=" + p.TransactionID.ToString)

    '        retCode = ArgenteaCOMObject.GiftCardRedeem(p.IntValue, 0, p.Barcode, p.TransactionID, p.ErrorMessage, p.MessageOut)
    '        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". Giftcard: " & p.Barcode & ". Error: " & p.ErrorMessage & ". Output: " & p.MessageOut)

    '        If retCode <> ArgenteaFunctionsReturnCode.OK OrElse String.IsNullOrEmpty(p.MessageOut) Then
    '            CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
    '        Else
    '            CheckRedeemGiftCard = IGiftCardReturnCode.OK
    '            CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
    '        End If
    '        LOG_Error(funcName, "Return : " + CSV)

    '        Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
    '        objTPTAHelperArgentea.HandleReturnString(p.Transaction,
    '                                                 p.Controller,
    '                                                 CSV,
    '                                                 InternalArgenteaFunctionTypes.GiftCardRedeemPreCkeck,
    '                                                 Me.Parameters,
    '                                                 taArgenteaEMVRec)
    '        taArgenteaEMVRec.theHdr.lTaRefToCreateNmbr = p.MediaRecord.theHdr.lTaCreateNmbr

    '        p.Status = ArgenteaGiftCardStatus.RedeemWithCheckMode.ToString

    '    Catch ex As Exception
    '        LOG_Error(getLocationString(funcName), ex.Message)
    '    Finally
    '        LOG_FuncExit(funcName, " returns " & CheckRedeemGiftCard.ToString())
    '        FormHelper.ShowWaitScreen(p.Controller, True, frm)
    '        ShowError(p)
    '    End Try

    'End Function

#End Region

#Region "IGiftCardRedeem"

    Public Function RedeemGiftCard(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "RedeemGiftCard"
        LOG_FuncStart(funcName)
        Dim frm As System.Windows.Forms.Form = Nothing
        Dim response As New ArgenteaResponse

        Try
            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)

            Dim szBarcode As String = MyCurrentRecord.GetPropertybyName("szITGiftCardEAN")
            Dim lValue As Integer = CInt(MyTaMediaRec.dTaPaid * 100) '[CO 20220307] From MyTaMediaRec.dTaPaidTotal To MyTaMediaRec.dTaPaid
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            response.ReturnCode = ArgenteaCOMObject.GiftCardRedeem(lValue, 0, szBarcode, response.TransactionID, szErrorMessage, szMessageOut)

            szBarcode = MyCurrentRecord.GetPropertybyName("szITGiftCardEAN")

            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)
            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK OrElse String.IsNullOrEmpty(szMessageOut) Then
                response.MessageOut = "KO" & ";" & szMessageOut & ";" & szErrorMessage
            Else
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If

            response.SetProperty("szBarcode", szBarcode)
            response.SetProperty("lAmount", lValue)
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.GiftCardRedeemPreCkeck

            paramArg.Copies = paramArg.GiftCardRedeemCheckCopies
            paramArg.PrintWithinTA = paramArg.GiftCardRedeemCheckPrintWithinTa

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
            LOG_FuncExit(funcName, " returns " & response.ReturnCode.ToString())

        End Try
        Return response
    End Function

#End Region

#Region "IGiftCardCancellationPayment"

    'Public Function GiftCardCancellation(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardCancellationPayment.GiftCardCancellation
    '    GiftCardCancellation = IGiftCardReturnCode.KO
    '    Dim funcName As String = "IGiftCardCancellationPayment"

    '    Dim frm As System.Windows.Forms.Form = Nothing
    '    Dim p As GiftCardRedeemParameters = New GiftCardRedeemParameters
    '    Dim CSV As String = String.Empty
    '    Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

    '    Try
    '        p.LoadCommonFunctionParameter(parameters)

    '        FormHelper.ShowWaitScreen(p.Controller, False, frm)

    '        ArgenteaCOMObject = Nothing
    '        ArgenteaCOMObject = New ARGLIB.argpay()

    '        retCode = ArgenteaCOMObject.GiftCardCancellation(p.IntValue, p.TransactionID, p.Barcode, p.ErrorMessage, p.MessageOut)
    '        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". Giftcard: " & p.Barcode & ". Error: " & p.ErrorMessage & ". Output: " & p.MessageOut)

    '        If retCode <> ArgenteaFunctionsReturnCode.OK Then
    '            CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
    '        Else
    '            GiftCardCancellation = IGiftCardReturnCode.OK
    '            CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
    '        End If

    '        Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
    '        objTPTAHelperArgentea.HandleReturnString(p.Transaction,
    '                                                 p.Controller,
    '                                                 CSV,
    '                                                 InternalArgenteaFunctionTypes.GiftCardRedeemCancel,
    '                                                 Me.Parameters)

    '        p.Status = ArgenteaGiftCardStatus.RedeemCanceled.ToString

    '        GiftCardCancellation = IGiftCardReturnCode.OK

    '    Catch ex As Exception
    '        LOG_Error(getLocationString(funcName), ex.Message)
    '    Finally
    '        FormHelper.ShowWaitScreen(p.Controller, True, frm)
    '        ShowError(p)
    '    End Try

    'End Function

    Public Function GiftCardCancellation(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                         ByRef taobj As TPDotnet.Pos.TA,
                         ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                         ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                         ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                         ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "IGiftCardCancellationPayment"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim CSV As String = String.Empty
        Dim myTaExternalService As TaExternalServiceRec = Nothing
        Dim response As New ArgenteaResponse
        Dim szTransactionID As String = String.Empty
        Dim lAmount As Integer = 0
        Dim szBarcode As String = String.Empty
        Dim szErrorMessage As String = String.Empty
        Dim szMessageOut As String = String.Empty
        Try

            LOG_Debug(getLocationString(funcName), "We are entered in Argentea void function")
            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)

            For i As Integer = taobj.taCollection.Count To 1 Step -1
                Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = taobj.GetTALine(i)
                If MyTaBaseRec.sid = TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE AndAlso
                    MyTaBaseRec.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr AndAlso
                    (MyTaBaseRec.ExistField("szStatus") AndAlso
                    MyTaBaseRec.GetPropertybyName("szStatus") = TPDotnet.IT.Common.Pos.TaExternalServiceRec.ExternalServiceStatus.Activated.ToString) Then
                    myTaExternalService = CType(MyTaBaseRec, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                    Exit For
                End If
            Next
            myTaExternalService.lCopies = 0
            myTaExternalService.bPrintReceipt = False

            szTransactionID = IIf(myTaExternalService.ExistField("szTransactionID"), myTaExternalService.GetPropertybyName("szTransactionID"), String.Empty)
            lAmount = IIf(myTaExternalService.ExistField("lAmount"), CInt(myTaExternalService.GetPropertybyName("lAmount")), 0)
            szBarcode = IIf(myTaExternalService.ExistField("szBarcode"), myTaExternalService.GetPropertybyName("szBarcode"), String.Empty)
            If szTransactionID = String.Empty Or lAmount < 0 Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            response.ReturnCode = ArgenteaCOMObject.GiftCardCancellation(lAmount, szTransactionID, szBarcode, szErrorMessage, szMessageOut)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)

            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                CSV = "KO" & ";" & szMessageOut & ";" & szErrorMessage
                response.SetProperty("szErrorMessage", szErrorMessage)
            Else
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If

            response.SetProperty("szBarcode", szBarcode)
            response.SetProperty("lAmount", Decimal.Negate(lAmount))
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.GiftCardRedeemCancel

            paramArg.Copies = paramArg.GiftCardRedeemCancelCopies
            paramArg.PrintWithinTA = paramArg.GiftCardRedeemSave


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
