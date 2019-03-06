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

    Public Function Check(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "PaymentCeliachia"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        'Dim p As BPParameters = New BPParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea PaymentCeliachia function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            'p.LoadCommonFunctionParameter(Parameters)
            Dim CSV As String = String.Empty
            Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

            Dim szTransID As String = taobj.lRetailStoreID.ToString().PadLeft(7, "0") + taobj.lWorkStationNmbr.ToString().PadLeft(3, "0") + taobj.lactTaNmbr.ToString().PadLeft(6, "0")


            Dim szListEan As String = String.Empty
            If Not Common.ApplyFilterStyleSheet(TheModCntr, taobj, "BPCeliac.xslt", szListEan) Then

            End If
            Dim xDoc As XDocument = taobj.TAtoXDocument(False, 0, False)
            Dim xel As List(Of XElement) = xDoc.XPathSelectElements("//ART_SALE[Hdr/bTaValid='1']/ARTICLE[szITSpecialItemType='CELIAC']/../dTaTotal").ToList()
            Dim lAmount As Integer = CInt((xel.Sum(Function(item) CDec(item.Value.ToString().Replace(".", ",")))) * 100)

            Dim xelMedia As List(Of XElement) = xDoc.XPathSelectElements("//MEDIA[Hdr/bTaValid='1']/PAYMENT[szExternalID='BPCeliac']/../dTaPaid").ToList()
            Dim lAmountMediaPayed As Integer = CInt((xelMedia.Sum(Function(item) CDec(item.Value.ToString().Replace(".", ",")))) * 100)
            Dim myLastMediaRec As TPDotnet.Pos.TaMediaRec = CType(taobj.GetTALine(taobj.getLastMediaRecNr), TPDotnet.Pos.TaMediaRec)
            lAmountMediaPayed = lAmountMediaPayed - (CInt(myLastMediaRec.dTaPaid * 100))
            Dim cmd As TPDotnet.IT.Common.Pos.Common
            If lAmount = 0 Then
                cmd.ShowError(TheModCntr, "ArtCeliacNotFound", "LevelITCommonArgenteaArtCeliacNotFound")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            If lAmountMediaPayed >= lAmount Then
                cmd.ShowError(TheModCntr, "PaymentCeliacExceeded", "LevelITCommonArgenteaPaymentCeliacExceeded")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            lAmount = Math.Min(lAmount, (CInt(myLastMediaRec.dTaPaid * 100)))
            Dim szMessageOut As String = String.Empty
            response.ReturnCode = ArgenteaCOMObject.PaymentCeliachia(lAmount, szTransID, response.TransactionID, szListEan, String.Empty, szMessageOut)
            response.MessageOut = szMessageOut
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPCeliacPayment

        Catch ex As Exception
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response

    End Function

    Public Function Payment(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "PaymentCeliachia"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        'Dim p As BPParameters = New BPParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea PaymentCeliachia function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            'p.LoadCommonFunctionParameter(Parameters)

            Dim szTransID As String = taobj.lRetailStoreID.ToString().PadLeft(7, "0") + taobj.lWorkStationNmbr.ToString().PadLeft(3, "0") + taobj.lactTaNmbr.ToString().PadLeft(6, "0")
            Dim szListEan As String = String.Empty

            Common.ApplyFilterStyleSheet(TheModCntr, taobj, "BPCeliac.xslt", szListEan)

            Dim lAmount As Integer = CInt(CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).dTaPaid * 100)
            Dim szMessageOut As String = String.Empty
            response.ReturnCode = ArgenteaCOMObject.PaymentCeliachia(lAmount, szTransID, response.TransactionID, szListEan, String.Empty, szMessageOut)
            response.MessageOut = szMessageOut
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPCeliacPayment

        Catch ex As Exception
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response

    End Function


    Public Function Void(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                           ByRef taobj As TPDotnet.Pos.TA,
                           ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                           ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                           ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                           ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "StornoCeliachia"
        Dim response As New ArgenteaResponse
        Dim myTaExternalService As New TPDotnet.IT.Common.Pos.TaExternalServiceRec

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea PaymentCeliachia function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            Dim CSV As String = String.Empty
            Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO

            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
            For i As Integer = taobj.taCollection.Count To 1 Step -1
                Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = taobj.GetTALine(i)
                If MyTaBaseRec.sid = TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE AndAlso MyTaBaseRec.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr Then
                    myTaExternalService = CType(MyTaBaseRec, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                    Exit For
                End If
            Next
            Dim szMessageOut As String = String.Empty

            Dim transactionIdentifier As String = IIf(myTaExternalService.ExistField("szTransactionID"), myTaExternalService.GetPropertybyName("szTransactionID"), String.Empty)

            If transactionIdentifier = String.Empty Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            response.ReturnCode = ArgenteaCOMObject.StornoCeliachia(transactionIdentifier, szMessageOut)
            response.MessageOut = szMessageOut
            response.CharSeparator = CharSeparator.Minus
            response.FunctionType = InternalArgenteaFunctionTypes.BPCeliacVoid


        Catch ex As Exception
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
        End Try
        Return response
    End Function


#Region "Overridable"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
