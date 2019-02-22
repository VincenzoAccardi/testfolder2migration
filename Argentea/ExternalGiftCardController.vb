Imports System
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports Microsoft.VisualBasic
Imports System.Xml.Linq
Imports System.Xml.XPath
Imports System.Linq




Public Class ExternalGiftCardController
    Implements IExternalGiftCardActivation
    Implements IExternalGiftCardDeActivation
    Implements IExternalGiftCardConfirm


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

#Region "Public Function"

#Region "IExternalGiftCardActivation"
    Public Function ActivationExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardActivation.ActivationExternalGiftCard
        ActivationExternalGiftCard = IExternalGiftCardReturnCode.KO
        Dim funcName As String = "ActivationExternalGiftCard"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As ExternalGiftCardActivationParameters = New ExternalGiftCardActivationParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IExternalGiftCardActivation function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            p.LoadCommonFunctionParameter(Parameters)

            ' call in check mode
            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()
            'For i As Integer = 1 To p.Transaction.taCollection.Count
            '    Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = p.Transaction.GetTALine(i)

            '    Select Case MyTaBaseRec.sid
            '        Case TPDotnet.Pos.PosDef.TARecTypes.iTA_ART_SALE
            Try
                '                p.ArticleRecord = MyTaBaseRec
                'Dim MyTaArtSaleRec As TaArtSaleRec = MyTaBaseRec
                If p.ArticleRecord.theHdr.bIsVoided = 0 AndAlso TypeOf p.ArticleRecord.ARTinArtSale Is TPDotnet.IT.Common.Pos.ART Then
                    Dim ITART As TPDotnet.IT.Common.Pos.ART = p.ArticleRecord.ARTinArtSale
                    If ITART.szITSpecialItemType = TPDotnet.IT.Common.Pos.ExternalGiftCardItem Then

                        Try
                            Dim CSV As String = String.Empty
                            Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO
                            p.Barcode = GetBarcodeFromTemplate(p.Transaction, p.Controller, p.ArticleRecord.szInputString)
                            If String.IsNullOrEmpty(p.Barcode.Trim) Then
                                Dim giftForm As FormArgenteaItemInput = p.Controller.GetCustomizedForm(GetType(FormArgenteaItemInput), STRETCH_TO_SMALL_WINDOW)
                                giftForm.ArticleDescription = p.ArticleRecord.ARTinArtSale.szDesc
                                p.Barcode = giftForm.DisplayMe(p.Transaction, p.Controller, FormRoot.DialogActive.d1_DialogActive)
                                giftForm.Close()
                                If String.IsNullOrEmpty(p.Barcode) Then
                                    Exit Function
                                End If
                            End If

                            Dim szEanCode As String = IIf(String.IsNullOrEmpty(p.ArticleRecord.szItemLookupCode), p.ArticleRecord.ARTinArtSale.szPOSItemID, p.ArticleRecord.szItemLookupCode)
                            FormHelper.ShowWaitScreen(p.Controller, False, frm)
                            retCode = ArgenteaCOMObject.Attivazione(p.IntValue, 0, szEanCode, p.Barcode, p.MessageOut, p.ErrorMessage)
                            LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". Giftcard: " & p.Barcode & ". Error: " & p.ErrorMessage & ". Output: " & p.MessageOut)

                            If retCode <> ArgenteaFunctionsReturnCode.OK Then
                                ActivationExternalGiftCard = IExternalGiftCardReturnCode.KO
                                ' Show an error for each gift card that cannot be definitely activated
                                LOG_Error(getLocationString(funcName), "Activation for giftcard  " & p.Barcode & " returns error: " & p.ErrorMessage)
                                CSV = "KO" & ";" & p.MessageOut & vbCrLf &
                                                    "!!!ERRORE DI ATTIVAZIONE!!!" & vbCrLf &
                                                     p.ErrorMessage & vbCrLf &
                                                    "Giftcard Serial: " & p.Barcode & vbCrLf &
                                                    "Value: " & Math.Round(p.IntValue / 100, 2) & vbCrLf & vbCrLf & " " & ";" & p.ErrorMessage
                                Exit Function
                            Else
                                LOG_Debug(getLocationString(funcName), "Gift card number " & p.Barcode & " successfuly activated")
                                CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
                                ActivationExternalGiftCard = IExternalGiftCardReturnCode.OK

                            End If

                            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
                            objTPTAHelperArgentea.HandleReturnString(p.Transaction,
                                                                     p.Controller,
                                                                     CSV,
                                                                     InternalArgenteaFunctionTypes.ExternalGiftCardActivation,
                                                                     Me.Parameters)
                            p.Status = ArgenteaExternalGiftCardStatus.ActivatedDefinitively.ToString
                        Catch ex As Exception
                            LOG_Error(getLocationString(funcName), ex.Message)
                        Finally
                            FormHelper.ShowWaitScreen(p.Controller, True, frm)
                            ShowError(p)
                        End Try

                        Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = p.Transaction.GetTALine(p.Transaction.taCollection.Count)
                        If MyTaBaseRec.sid = TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE Then
                            Dim myTaArgEmv As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(MyTaBaseRec, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                            If myTaArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardActivation.ToString OrElse myTaArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation.ToString Then
                                If Not myTaArgEmv.ExistField("szITSerialCode") Then myTaArgEmv.AddField("szITSerialCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                                myTaArgEmv.setPropertybyName("szITSerialCode", p.Barcode.ToString)
                            End If
                        End If
                    End If

                End If

            Catch ex As Exception
                LOG_Error(getLocationString(funcName), ex.Message)
            End Try
            ' Exit Select

            '        End Select
            ' Next i

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally

        End Try
    End Function

#End Region

#Region "IExternalGiftCardDeActivation"
    Public Function DeActivationExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardDeActivation.DeActivationExternalGiftCard
        DeActivationExternalGiftCard = IExternalGiftCardReturnCode.KO
        Dim funcName As String = "IExternalGiftCardDeActivation"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As ExternalGiftCardDeActivationParameters = New ExternalGiftCardDeActivationParameters
        Dim CSV As String = String.Empty
        Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

        Try
            p.LoadCommonFunctionParameter(Parameters)

            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()
            If p.ArticleRecord IsNot Nothing Then
                p.Barcode = GetBarcodeFromTemplate(p.Transaction, p.Controller, p.ArticleRecord.szInputString)
            End If

            If String.IsNullOrEmpty(p.Barcode.Trim) Then
                Dim giftForm As FormArgenteaItemInput = p.Controller.GetCustomizedForm(GetType(FormArgenteaItemInput), STRETCH_TO_SMALL_WINDOW)
                If p.ArticleRecord IsNot Nothing Then
                    giftForm.ArticleDescription = p.ArticleRecord.ARTinArtSale.szDesc
                Else
                    giftForm.ArticleDescription = p.ArticleReturnRecord.ARTinArtReturn.szDesc
                End If
                p.Barcode = giftForm.DisplayMe(p.Transaction, p.Controller, FormRoot.DialogActive.d1_DialogActive)
                giftForm.Close()
                If String.IsNullOrEmpty(p.Barcode) Then
                    Exit Function
                End If
            End If
            FormHelper.ShowWaitScreen(p.Controller, False, frm)

            Dim szEanCode As String = String.Empty

            If p.ArticleRecord IsNot Nothing Then
                szEanCode = IIf(String.IsNullOrEmpty(p.ArticleRecord.szItemLookupCode), p.ArticleRecord.ARTinArtSale.szPOSItemID, p.ArticleRecord.szItemLookupCode)
            Else
                szEanCode = IIf(String.IsNullOrEmpty(p.ArticleReturnRecord.szItemLookupCode), p.ArticleReturnRecord.ARTinArtReturn.szPOSItemID, p.ArticleReturnRecord.szItemLookupCode)
            End If

            retCode = ArgenteaCOMObject.DisAttivazione(p.IntValue, szEanCode, p.Barcode, p.MessageOut, p.ErrorMessage)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". Giftcard: " & p.Barcode & ". Error: " & p.ErrorMessage & ". Output: " & p.MessageOut)

            If retCode <> ArgenteaFunctionsReturnCode.OK Then
                CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
            Else
                DeActivationExternalGiftCard = IExternalGiftCardReturnCode.OK
                CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction,
                                                     p.Controller,
                                                     CSV,
                                                     InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation,
                                                     Me.Parameters)

            p.Status = ArgenteaExternalGiftCardStatus.Deactivated.ToString

            Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = p.Transaction.GetTALine(p.Transaction.taCollection.Count)
            If MyTaBaseRec.sid = TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE Then
                Dim myTaArgEmv As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(MyTaBaseRec, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                If myTaArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardActivation.ToString OrElse myTaArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation.ToString Then
                    If Not myTaArgEmv.ExistField("szITSerialCode") Then myTaArgEmv.AddField("szITSerialCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                    myTaArgEmv.setPropertybyName("szITSerialCode", p.Barcode.ToString)
                End If
            End If
            'DeActivationExternalGiftCard = IExternalGiftCardReturnCode.OK

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try

    End Function
#End Region

#Region "IExternalGiftCardConfirm"
    Public Function ConfirmExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardConfirm.ConfirmExternalGiftCard
        ConfirmExternalGiftCard = IExternalGiftCardReturnCode.KO
        Dim funcName As String = "IExternalGiftCardConfirm"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As ExternalGiftCardConfirmParameters = New ExternalGiftCardConfirmParameters
        Dim CSV As String = String.Empty
        Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

        Try
            p.LoadCommonFunctionParameter(Parameters)
            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            Dim argenteaTA As TPDotnet.Pos.TA = Nothing
            Me.Parameters.LoadParametersByReflection(p.Controller)
            Dim xDoc As Xml.Linq.XDocument = p.Transaction.TAtoXDocument(False, 0, False)

            Dim xElActiveList As List(Of XElement) = xDoc.XPathSelectElements("/TAS/NEW_TA/EXTERNAL_SERVICE[szFunctionID='ExternalGiftCardActivation']/szITSerialCode").ToList()
            Dim xElActiveSerialCode As New List(Of String)
            For Each xEl As XElement In xElActiveList
                xElActiveSerialCode.Add(xEl.Value)
            Next
            Dim xElDeActiveList As List(Of XElement) = xDoc.XPathSelectElements("/TAS/NEW_TA/EXTERNAL_SERVICE[szFunctionID='ExternalGiftCardDeActivation']/szITSerialCode").ToList()
            Dim xElDeActiveSerialCode As New List(Of String)
            For Each xEl As XElement In xElDeActiveList
                xElDeActiveSerialCode.Add(xEl.Value)
            Next
            Dim xel1 As List(Of String) = xElActiveSerialCode.Except(xElDeActiveSerialCode).ToList()
            Dim xel2 As List(Of String) = xElDeActiveSerialCode.Except(xElActiveSerialCode).ToList()

            Dim szITSerialCodeList As List(Of String) = xel1.Concat(xel2).ToList()


            For i As Integer = 1 To p.Transaction.taCollection.Count
                Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = p.Transaction.GetTALine(i)
                If MyTaBaseRec.sid = TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE Then
                    Dim myTaArgEmv As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(MyTaBaseRec, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                    If myTaArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardActivation.ToString OrElse myTaArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation.ToString Then
                        If szITSerialCodeList.Contains(myTaArgEmv.GetPropertybyName("szITSerialCode")) Then
                            argenteaTA = objTPTAHelperArgentea.CreateTA(p.Transaction, p.Controller, MyTaBaseRec, Me.Parameters.ExtGiftCardConfirmSave)
                            For j As Integer = 1 To Me.Parameters.ExtGiftCardConfirmCopies
                                If Not objTPTAHelperArgentea.PrintReceipt(argenteaTA, p.Controller) Then
                                    Throw New Exception("CANNOT_PRINT_ARGENTEA_TA")
                                End If
                            Next

                            If Me.Parameters.ExtGiftCardConfirmPrintWithinTa Then
                                Dim myTaCloneArgEmv As New TPDotnet.IT.Common.Pos.TaExternalServiceRec
                                myTaCloneArgEmv.Clone(myTaArgEmv, 0)
                                myTaCloneArgEmv.szFunctionID = InternalArgenteaFunctionTypes.ExternalGiftCardConfirm.ToString
                                myTaCloneArgEmv.theHdr.lTaCreateNmbr = 0
                                myTaCloneArgEmv.theHdr.lTaRefToCreateNmbr = 0
                                myTaCloneArgEmv.bPrintReceipt = True
                                myTaCloneArgEmv.szReceipt = myTaArgEmv.szOriginalReceipt
                                myTaCloneArgEmv.szOriginalReceipt = myTaArgEmv.szOriginalReceipt

                                p.Transaction.Add(myTaCloneArgEmv)

                            End If
                        End If
                    End If
                End If

            Next
            Dim lCreateNumbers As List(Of XElement) = xDoc.XPathSelectElements("/TAS/NEW_TA/EXTERNAL_SERVICE[szFunctionID='ExternalGiftCardActivation' or szFunctionID='ExternalGiftCardDeActivation']/Hdr/lTaCreateNmbr").ToList()
            For Each lCreateNmbr As XElement In lCreateNumbers
                p.Transaction.Remove(p.Transaction.GetPositionFromCreationNmbr(CInt(lCreateNmbr.Value)))
            Next

            p.Transaction.Reorg()

        Catch ex As Exception

        End Try


    End Function

#End Region
#End Region

#Region "Overridable"
    Protected Overridable Function GetBarcodeFromTemplate(taobj As TPDotnet.Pos.TA, TheModCntr As ModCntr, ByVal szInputString As String) As String
        Dim funcName As String = "GetBarcodeFromTemplate"

        GetBarcodeFromTemplate = ""
        Try
            TheModCntr.InputString_Renamed = szInputString
            TheModCntr.theBarcodeCls.SetNormalBuffer = TheModCntr.InputString_Renamed
            If ScanTemplateWithCheck("all", BARCODE_TEMPLATE, TheModCntr, taobj, False) = True Then

                If TheModCntr.theBarcodeCls.Hits > 0 Then
                    ' ok , we have hits for Barcode
                    For i As Integer = 1 To TheModCntr.theBarcodeCls.Hits
                        'we identified the barcode
                        'it is a voucher barcode, which is not serialized
                        'the MediaMemberNo is stored as ExternalID
                        Dim szString As String = TheModCntr.theBarcodeCls.GetMatchByName(i)
                        Select Case szString
                            Case TD_SERIAL_NUMBER
                                GetBarcodeFromTemplate = TheModCntr.theBarcodeCls.GetMatch(i)
                        End Select
                    Next i
                End If
            End If
        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            TheModCntr.theBarcodeCls.SetNormalBuffer = String.Empty
        End Try


        Return GetBarcodeFromTemplate
    End Function

    Protected Overridable Sub ShowError(ByRef TheModCntr As TPDotnet.Pos.ModCntr,
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

#End Region

End Class
