﻿Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing

Public Class GiftCardController
    Implements IGiftCardActivationPreCheck
    Implements IGiftCardActivation
    Implements IGiftCardBalanceInquiry
    Implements IGiftCardRedeemPreCheck
    Implements IGiftCardRedeem
    Implements IGiftCardCancellationPayment

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

#Region "Argentea specific"

    Protected ArgenteaCOMObject As ARGLIB.argpay

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

    Public Function CheckGiftCard(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardActivationPreCheck.CheckGiftCard
        CheckGiftCard = IGiftCardReturnCode.KO
        Dim funcName As String = "CheckGiftCard"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As GiftCardActivationParameters = New GiftCardActivationParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea CheckGiftCard function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            p.LoadCommonFunctionParameter(parameters)

            ' show the form in order to get the CODE128
            Dim giftForm As FormArgenteaItemInput = p.Controller.GetCustomizedForm(GetType(FormArgenteaItemInput), STRETCH_TO_SMALL_WINDOW)
            giftForm.ArticleDescription = p.ArticleRecord.ARTinArtSale.szDesc
            p.Barcode = giftForm.DisplayMe(p.Transaction, p.Controller, FormRoot.DialogActive.d1_DialogActive)
            giftForm.Close()
            If String.IsNullOrEmpty(p.Barcode) Then
                Exit Function
            End If

            ' call in check mode
            FormHelper.ShowWaitScreen(p.Controller, False, frm)
            ArgenteaCOMObject = New ARGLIB.argpay()

            If ArgenteaCOMObject.GiftCardActivation(p.IntValue, 1, p.Barcode, p.Barcode, p.MessageOut, p.ErrorMessage) <> ArgenteaFunctionsReturnCode.OK Then
                LOG_Debug(getLocationString(funcName), "Activation check for giftcard  " & p.Barcode & " returns error: " & p.ErrorMessage)
                Exit Function
            Else
                LOG_Debug(getLocationString(funcName), "Gift card number " & p.Barcode & " successfuly checked for activation")
            End If

            p.Status = ArgenteaGiftCardStatus.ActivatedWithCheckMode.ToString

            CheckGiftCard = IGiftCardReturnCode.OK

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try
    End Function

#End Region

#Region "IGiftCardActivation"

    Public Function ActivateGiftCard(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardActivation.ActivateGiftCard
        ActivateGiftCard = IGiftCardReturnCode.OK
        Dim funcName As String = "ActivateGiftCard"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As GiftCardActivationParameters = New GiftCardActivationParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IGiftCardActivation function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            p.LoadCommonFunctionParameter(parameters)

            ' call in check mode
            ArgenteaCOMObject = New ARGLIB.argpay()
            For i As Integer = 1 To p.Transaction.taCollection.Count
                Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = p.Transaction.GetTALine(i)

                Select Case MyTaBaseRec.sid
                    Case TPDotnet.Pos.PosDef.TARecTypes.iTA_ART_SALE
                        Try
                            p.ArticleRecord = MyTaBaseRec
                            'Dim MyTaArtSaleRec As TaArtSaleRec = MyTaBaseRec
                            If p.ArticleRecord.theHdr.bIsVoided = 0 AndAlso TypeOf p.ArticleRecord.ARTinArtSale Is TPDotnet.IT.Common.Pos.ART Then
                                Dim ITART As TPDotnet.IT.Common.Pos.ART = p.ArticleRecord.ARTinArtSale
                                If ITART.szITSpecialItemType = TPDotnet.IT.Common.Pos.GiftCardItem Then

                                    Try
                                        Dim CSV As String = String.Empty
                                        FormHelper.ShowWaitScreen(p.Controller, False, frm)
                                        If ArgenteaCOMObject.GiftCardActivation(p.IntValue, 0, p.Barcode, p.Barcode, p.MessageOut, p.ErrorMessage) <> ArgenteaFunctionsReturnCode.OK Then
                                            ActivateGiftCard = IGiftCardReturnCode.KO
                                            ' Show an error for each gift card that cannot be definitely activated
                                            LOG_Debug(getLocationString(funcName), "Activation for giftcard  " & p.Barcode & " returns error: " & p.ErrorMessage)
                                            CSV = "KO" & ";" & p.MessageOut & vbCrLf & _
                                                                "!!!ERRORE DI ATTIVAZIONE!!!" & vbCrLf & _
                                                                 p.ErrorMessage & vbCrLf & _
                                                                "Giftcard Serial: " & p.Barcode & vbCrLf & _
                                                                "Value: " & Math.Round(p.IntValue / 100, 2) & vbCrLf & vbCrLf & " " & ";" & p.ErrorMessage
                                        Else
                                            LOG_Debug(getLocationString(funcName), "Gift card number " & p.Barcode & " successfuly activated")
                                            CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
                                        End If

                                        Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
                                        objTPTAHelperArgentea.HandleReturnString(p.Transaction, _
                                                                                 p.Controller, _
                                                                                 CSV, _
                                                                                 InternalArgenteaFunctionTypes.GiftCardActivation, _
                                                                                 Me.Parameters)
                                        p.Status = ArgenteaGiftCardStatus.ActivatedDefinitively.ToString

                                    Catch ex As Exception
                                        LOG_Error(getLocationString(funcName), ex.Message)
                                    Finally
                                        FormHelper.ShowWaitScreen(p.Controller, True, frm)
                                        ShowError(p)
                                    End Try

                                End If

                            End If

                        Catch ex As Exception
                            LOG_Error(getLocationString(funcName), ex.Message)
                        End Try
                        Exit Select

                End Select
            Next i

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally

        End Try
    End Function

#End Region

#Region "IGiftCardBalanceInquiry"

    Public Function GiftCardBalanceInquiry(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardBalanceInquiry.GiftCardBalanceInquiry
        GiftCardBalanceInquiry = IGiftCardReturnCode.KO
        Dim funcName As String = "GiftCardBalanceInquiry"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As BalanceParameters = New BalanceParameters
        Dim CSV As String = String.Empty

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IGiftCardBalanceInquiry function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            ' collect the input parameters
            p.LoadCommonFunctionParameter(parameters)

            Me.Parameters.LoadParametersByReflection(p.Controller)
            p.GiftCardBalanceLineIdentifier = Me.Parameters.GiftCardBalanceLineIdentifier
            Me.Parameters.GiftCardBalanceInternalInquiry = p.GiftCardBalanceInternalInquiry
            ' check the balance
            FormHelper.ShowWaitScreen(p.Controller, False, frm)
            ArgenteaCOMObject = New ARGLIB.argpay()
            If ArgenteaCOMObject.GiftCardBalance(p.Barcode, p.ErrorMessage, p.MessageOut) <> ArgenteaFunctionsReturnCode.OK Then
                LOG_Debug(getLocationString(funcName), "Balance for giftcard  " & p.Barcode & " returns error: " & p.ErrorMessage)
                CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
            Else
                GiftCardBalanceInquiry = IGiftCardReturnCode.OK
                CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction, _
                                                     p.Controller, _
                                                     CSV, _
                                                     InternalArgenteaFunctionTypes.GiftCardBalance, _
                                                     Me.Parameters)

            If String.IsNullOrEmpty(p.MessageOut) Then
                ' strange situation: function returns ok but without receipt
                ' log log log
                Exit Function
            End If

            ' copy back the values for value type fields
            If parameters.ContainsKey("Value") Then parameters("Value") = p.Value
            If parameters.ContainsKey("Receipt") Then parameters("Receipt") = p.Receipt

            GiftCardBalanceInquiry = IGiftCardReturnCode.OK

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try
    End Function

#End Region

#Region "IGiftCardRedeemPreCheck"

    Public Function CheckRedeemGiftCard(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardRedeemPreCheck.CheckRedeemGiftCard
        CheckRedeemGiftCard = IGiftCardReturnCode.KO
        Dim funcName As String = "CheckRedeemGiftCard"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As GiftCardRedeemParameters = New GiftCardRedeemParameters
        Dim CSV As String = String.Empty
        Dim taArgenteaEMVRec As TaArgenteaEMVRec = Nothing

        Try
            p.LoadCommonFunctionParameter(parameters)

            FormHelper.ShowWaitScreen(p.Controller, False, frm)
            ArgenteaCOMObject = New ARGLIB.argpay()
            If ArgenteaCOMObject.GiftCardRedeem(p.IntValue, 0, p.Barcode, p.TransactionID, p.ErrorMessage, p.MessageOut) <> ArgenteaFunctionsReturnCode.OK Then
                CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
            Else
                CheckRedeemGiftCard = IGiftCardReturnCode.OK
                CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction, _
                                                     p.Controller, _
                                                     CSV, _
                                                     InternalArgenteaFunctionTypes.GiftCardRedeemPreCkeck, _
                                                     Me.Parameters, _
                                                     taArgenteaEMVRec)
            taArgenteaEMVRec.theHdr.lTaRefToCreateNmbr = p.MediaRecord.theHdr.lTaCreateNmbr

            p.Status = ArgenteaGiftCardStatus.RedeemWithCheckMode.ToString

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try

    End Function

#End Region

#Region "IGiftCardRedeem"

    Public Function RedeemGiftCard(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardRedeem.RedeemGiftCard
        Dim funcName As String = "RedeemGiftCard"

        Try
            ' the argentea gift card is activated immediatly by the CheckRedeemGiftCard function
            RedeemGiftCard = IGiftCardReturnCode.OK

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Function

#End Region

#Region "IGiftCardCancellationPayment"

    Public Function GiftCardCancellation(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardCancellationPayment.GiftCardCancellation
        GiftCardCancellation = IGiftCardReturnCode.KO
        Dim funcName As String = "IGiftCardCancellationPayment"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As GiftCardRedeemParameters = New GiftCardRedeemParameters
        Dim CSV As String = String.Empty

        Try
            p.LoadCommonFunctionParameter(parameters)

            FormHelper.ShowWaitScreen(p.Controller, False, frm)
            ArgenteaCOMObject = New ARGLIB.argpay()
            If ArgenteaCOMObject.GiftCardCancellation(p.IntValue, p.TransactionID, p.Barcode, p.ErrorMessage, p.MessageOut) <> ArgenteaFunctionsReturnCode.OK Then
                CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
            Else
                GiftCardCancellation = IGiftCardReturnCode.OK
                CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction, _
                                                     p.Controller, _
                                                     CSV, _
                                                     InternalArgenteaFunctionTypes.GiftCardRedeemCancel, _
                                                     Me.Parameters)

            p.Status = ArgenteaGiftCardStatus.RedeemCanceled.ToString

            GiftCardCancellation = IGiftCardReturnCode.OK

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try

    End Function

#End Region

#Region "Overridable"

    Protected Overridable Sub ShowError(ByRef TheModCntr As TPDotnet.Pos.ModCntr, _
                                        ByRef err As String)
        Dim funcName As String = "ShowError"
        Dim szTranslatedError As String = err

        Try
            If Not TheModCntr Is Nothing AndAlso Not String.IsNullOrEmpty(err) Then

                LOG_Debug(getLocationString(funcName), err)

                szTranslatedError = getPosTxtNew(TheModCntr.contxt, "LevelITCommonArgentea" & err, 0)
                If String.Equals(szTranslatedError, "message  0 not found", StringComparison.InvariantCultureIgnoreCase) Then
                    LOG_Error(getLocationString(funcName), "Message does not exists:" & err & ". Use the original one.")
                    szTranslatedError = err
                End If

                ' not nice but we don't have a list of error codes
                TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                             szTranslatedError,
                             Integer.MaxValue - 1,
                             TheModCntr,
                             "LevelITCommonArgentea" & err)

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Sub

    Protected Overridable Sub ShowError(ByRef p As CommonParameters)
        Dim funcName As String = "ShowError"

        Try
            If Not p Is Nothing AndAlso Not String.IsNullOrEmpty(p.ErrorMessage) Then

                LOG_Debug(getLocationString(funcName), p.ErrorMessage)

                ShowError(p.Controller, p.ErrorMessage)

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Sub

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

    'Protected Overridable Sub ShowWaitScreen(ByRef TheModCntr As ModCntr, ByVal bClear As Boolean, ByRef form As System.Windows.Forms.Form, Optional ByVal customMsg As String = "", Optional ByVal addCustomMsg As String = "")

    '    Dim i As Integer = -1
    '    Dim resolution As String = String.Empty

    '    Try
    '        LOG_FuncStart(getLocationString("ShowInfo"), "function started")

    '        If TheModCntr.bCalledFromWebService OrElse TheModCntr.bExternalDialog Then
    '            Exit Sub
    '        End If

    '        If bClear Then

    '            If form IsNot Nothing Then
    '                form.Close()
    '                If Not form Is Nothing Then
    '                    If TheModCntr IsNot Nothing Then
    '                        TheModCntr.EndForm()
    '                    End If
    '                End If
    '                form = Nothing
    '            End If

    '        Else

    '            Dim msg As String = IIf(Not String.IsNullOrEmpty(customMsg), customMsg, getPosTxtNew((TheModCntr.contxt), "Message", TEXT_PLEASE_WAIT))
    '            msg &= customMsg

    '            form = TPMsg(msg, TEXT_PLEASE_WAIT, TheModCntr, "Message")
    '            Dim lx As Integer = (TheModCntr.GUICntr.ThePosForm.Width / 2) - (form.Width / 2)
    '            Dim ly As Integer = (TheModCntr.GUICntr.ThePosForm.Height / 2) - (form.Height / 2)
    '            form.Location = New System.Drawing.Point(lx, ly)

    '            'Try
    '            '    resolution = TheModCntr.GUICntr.ThePosForm.Width & "x" & TheModCntr.GUICntr.ThePosForm.Height
    '            '    i = Array.FindIndex(TheModCntr.GUICntr.POSGUIConfig.SubFormSizes, _
    '            '                                       Function(x As TPDotnet.Pos.SubFormSize) _
    '            '                                           x.Type = NO_STRETCH.ToString _
    '            '                                           AndAlso _
    '            '                                           x.Resolution = resolution)
    '            '    If i >= 0 Then




    '            '        form.BackgroundImageLayout = Windows.Forms.ImageLayout.Stretch

    '            '        form.SetBounds(TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.X, _
    '            '                       TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.Y, _
    '            '                       TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.Width, _
    '            '                       TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.Height)
    '            '    End If
    '            'Catch ex As Exception

    '            'End Try
    '            System.Windows.Forms.Application.DoEvents()

    '        End If

    '        Exit Sub

    '    Catch ex As Exception
    '        Try
    '            LOG_Error(getLocationString("ShowInfo"), ex)

    '        Catch InnerEx As Exception
    '            LOG_ErrorInTry(getLocationString("ShowInfo"), InnerEx)
    '        End Try
    '    Finally
    '        LOG_FuncExit(getLocationString("ShowInfo"), "end of function")
    '    End Try
    'End Sub

#End Region

End Class
