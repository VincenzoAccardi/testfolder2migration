Imports System
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports Microsoft.VisualBasic
Imports System.Xml.Linq
Imports System.Xml.XPath
Imports System.Linq
Imports PAGAMENTOLib

Public Class BPCeliacController
    Implements IBPCeliachia


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
    Public Function PaymentCeliachia(ByRef Parameters As Dictionary(Of String, Object)) As IBPReturnCode Implements IBPCeliachia.PaymentCeliachia
        PaymentCeliachia = IBPReturnCode.KO
        Dim funcName As String = "PaymentCeliachia"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As BPParameters = New BPParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea PaymentCeliachia function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            p.LoadCommonFunctionParameter(Parameters)
            Dim CSV As String = String.Empty
            Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO


            ' call in check mode
            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()

            FormHelper.ShowWaitScreen(p.Controller, False, frm)
            Dim TransID As String = p.Transaction.lRetailStoreID.ToString().PadLeft(7, "0") + p.Transaction.lWorkStationNmbr.ToString().PadLeft(3, "0") + p.Transaction.lactTaNmbr.ToString().PadLeft(6, "0")


            Dim szListEan As String = String.Empty
            If Not Common.ApplyFilterStyleSheet(p.Controller, p.Transaction, "BPCeliac.xslt", szListEan) Then

            End If
            Dim xDoc As XDocument = p.Transaction.TAtoXDocument(False, 0, False)
            Dim xel As List(Of XElement) = xDoc.XPathSelectElements("//ART_SALE[Hdr/bTaValid='1']/ARTICLE[szITSpecialItemType='CELIAC']/../dTaTotal").ToList()
            Dim lAmount As Integer = CInt((xel.Sum(Function(item) CDec(item.Value.ToString().Replace(".", ",")))) * 100)

            Dim xelMedia As List(Of XElement) = xDoc.XPathSelectElements("//MEDIA[Hdr/bTaValid='1']/PAYMENT[szExternalID='BPCeliac']/../dTaPaid").ToList()
            Dim lAmountMediaPayed As Integer = CInt((xelMedia.Sum(Function(item) CDec(item.Value.ToString().Replace(".", ",")))) * 100)
            Dim myLastMediaRec As TPDotnet.Pos.TaMediaRec = CType(p.Transaction.GetTALine(p.Transaction.getLastMediaRecNr), TPDotnet.Pos.TaMediaRec)
            lAmountMediaPayed = lAmountMediaPayed - (CInt(myLastMediaRec.dTaPaid) * 100)
            If lAmount = 0 Then
                ShowError(p.Controller, "ArtCeliacNotFound")
                Return IBPReturnCode.KO
            End If

            If lAmountMediaPayed >= lAmount Then
                ShowError(p.Controller, "PaymentCeliacExceeded")
                Return IBPReturnCode.KO
            End If

            lAmount = Math.Min(lAmount, (CInt(myLastMediaRec.dTaPaid) * 100))

            Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject

            retCode = PaymentCeliac(p.Controller, ArgenteaCOMObject, argenteaFunctionReturnObject, lAmount, TransID, "", szListEan, p.MessageOut)

            If retCode <> ArgenteaFunctionsReturnCode.OK Then
                Exit Function
            Else
                CSV = p.MessageOut
                PaymentCeliachia = IBPReturnCode.OK
            End If


            'argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction,
                                                                     p.Controller,
                                                                     CSV,
                                                                     InternalArgenteaFunctionTypes.BPCeliacPayment,
                                                                     Me.Parameters)
            If Not argenteaFunctionReturnObject(0).Successfull Then
                Return IBPReturnCode.KO
            End If

            Dim dAmount As Decimal = lAmount / 100
            p.MediaRecord.dPaidForeignCurr = dAmount
            p.MediaRecord.dPaidForeignCurrTotal = dAmount
            p.MediaRecord.dTaPaid = dAmount
            p.MediaRecord.dTaPaidTotal = dAmount

            RegistryHelper.SetLastPaymentBPCeliacTransactionIdentifier(argenteaFunctionReturnObject(0).TerminalID)


        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try

    End Function

    Private Function PaymentCeliac(ByRef TheModCntr As TPDotnet.Pos.ModCntr, argenteaCOMObject As argpay, ByRef argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject, ByRef lAmount As Integer, szReceiptNumber As String, szTransID As String, szListEan As String, ByRef szMessageOut As String) As ArgenteaFunctionsReturnCode
        PaymentCeliac = argenteaCOMObject.PaymentCeliachia(lAmount, szReceiptNumber, szTransID, szListEan, String.Empty, szMessageOut)

        If PaymentCeliac <> ArgenteaFunctionsReturnCode.OK Then
            PaymentCeliac = ArgenteaFunctionsReturnCode.KO
        Else

            If (Not CSVHelper.ParseReturnString(szMessageOut,
                                                InternalArgenteaFunctionTypes.BPCeliacPayment,
                                                argenteaFunctionReturnObject)) Then
                Throw New Exception("CSV_INVALID")
            End If
            Dim frm As System.Windows.Forms.Form = Nothing
            If Not argenteaFunctionReturnObject(0).Successfull Then
                If argenteaFunctionReturnObject(0).CodeResult = "0300" Then
                    FormHelper.ShowWaitScreen(TheModCntr, True, frm)
                    If MsgBoxResult.Ok = ShowQuestion(TheModCntr, "UnderFunded", CDec((CInt(argenteaFunctionReturnObject(0).Amount) / 100)).ToString) Then
                        lAmount = CInt(argenteaFunctionReturnObject(0).Amount)
                        FormHelper.ShowWaitScreen(TheModCntr, False, frm)

                        PaymentCeliac = PaymentCeliac(TheModCntr, argenteaCOMObject, argenteaFunctionReturnObject, lAmount, szReceiptNumber, szTransID, szListEan, szMessageOut)
                    Else
                        Return ArgenteaFunctionsReturnCode.KO
                    End If
                Else
                    Return ArgenteaFunctionsReturnCode.KO
                End If
            End If

        End If
        Return PaymentCeliac
    End Function

    Public Function StornoCeliachia(ByRef Parameters As Dictionary(Of String, Object)) As IBPReturnCode Implements IBPCeliachia.StornoCeliachia
        StornoCeliachia = IBPReturnCode.KO
        Dim funcName As String = "StornoCeliachia"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As BPParameters = New BPParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea PaymentCeliachia function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            p.LoadCommonFunctionParameter(Parameters)
            Dim CSV As String = String.Empty
            Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO


            ' call in check mode
            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()

            FormHelper.ShowWaitScreen(p.Controller, False, frm)

            retCode = ArgenteaCOMObject.StornoCeliachia(RegistryHelper.GetLastPaymentBPCeliacTransactionIdentifier, p.MessageOut)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". Card: " & p.Barcode & ". Error: " & p.ErrorMessage & ". Output: " & p.MessageOut)

            If retCode <> ArgenteaFunctionsReturnCode.OK Then
                StornoCeliachia = IBPReturnCode.KO
                ' Show an error for each gift card that cannot be definitely activated
                LOG_Error(getLocationString(funcName), "Payment for celiac  " & p.Barcode & " returns error: " & p.ErrorMessage)
                CSV = p.MessageOut
                Exit Function
            Else
                LOG_Debug(getLocationString(funcName), "Gift card number " & p.Barcode & " successfuly activated")
                CSV = p.MessageOut
                StornoCeliachia = IBPReturnCode.OK
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction,
                                                                     p.Controller,
                                                                     CSV,
                                                                     InternalArgenteaFunctionTypes.BPCeliacVoid,
                                                                     Me.Parameters)


        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            FormHelper.ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try
    End Function


#Region "Overridable"

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

    Protected Overridable Function ShowQuestion(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef err As String, ByVal szAdditionalMessage As String) As Integer
        Dim funcName As String = "ShowQuestion"

        Dim szTranslatedError As String = err

        Try
            If Not TheModCntr Is Nothing AndAlso Not String.IsNullOrEmpty(err) Then

                LOG_Debug(getLocationString(funcName), err)

                szTranslatedError = getPosTxtNew(TheModCntr.contxt, "LevelITCommonArgentea" & err, 0)
                szTranslatedError = String.Format(szTranslatedError, szAdditionalMessage)
                If String.Equals(szTranslatedError, "message  0 not found", StringComparison.InvariantCultureIgnoreCase) Then
                    LOG_Error(getLocationString(funcName), "Message does not exists:" & err & ". Use the original one.")
                    szTranslatedError = err
                End If

                ' not nice but we don't have a list of error codes
                Return TPMsgBoxRet(PosDef.TARMessageTypes.TPQUESTION,
                             szTranslatedError, 1,
                             TheModCntr, 1,
                             "LevelITCommonArgentea" & err)

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Function

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
