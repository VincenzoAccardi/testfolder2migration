Imports System
Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
Imports ARGLIB = PAGAMENTOLib

Public Class BPEController
    Public Function Payment(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                        ByRef taobj As TPDotnet.Pos.TA,
                        ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                        ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                        ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                        ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "Payment"
        Dim transactionIdentifier As String = String.Empty
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea payment function")
            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO

            Dim szTotal As String = String.Empty
            If Not Common.ApplyFilterStyleSheet(TheModCntr, taobj, "BPEType.xslt", szTotal) Then
                szTotal = String.Empty
            End If

            Dim dTotal As Integer = CInt(MyTaMediaRec.dTaPaidTotal * 100)
            If Not String.IsNullOrEmpty(szTotal) Then
                Dim lxsltTotal As Integer = CInt(szTotal.Replace(".", ",") * 100)
                dTotal = Math.Min(dTotal, lxsltTotal)
            End If

            response.ReturnCode = ArgenteaCOMObject.PaymentBPE(dTotal, response.TransactionID, returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPEPayment
            response.TransactionID = transactionIdentifier

            paramArg.Copies = paramArg.BPECopies
            paramArg.PrintWithinTA = paramArg.BPEPrintWithinTa

        Catch ex As Exception
            Try
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & response.ReturnCode.ToString)
        End Try
        Return response

    End Function

    Public Function Balance(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "Balance"
        Dim transactionIdentifier As String = String.Empty
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea balance function")
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO

            response.ReturnCode = ArgenteaCOMObject.BalanceBPE(returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPEBalance
            response.TransactionID = transactionIdentifier

            paramArg.Copies = paramArg.BPECopies
            paramArg.PrintWithinTA = paramArg.BPEPrintWithinTa

        Catch ex As Exception
            Try
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & response.ReturnCode.ToString)
        End Try
        Return response

    End Function

    Public Function Void(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                           ByRef taobj As TPDotnet.Pos.TA,
                           ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                           ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                           ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                           ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "Void"
        Dim returnString As String = String.Empty
        Dim transactionIdentifier As String = String.Empty
        Dim transactionAmount As Double = 0
        Dim myTaExternalService As TaExternalServiceRec = Nothing
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea void function")
            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)

            For i As Integer = taobj.taCollection.Count To 1 Step -1
                Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = taobj.GetTALine(i)
                If MyTaBaseRec.sid = TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE AndAlso MyTaBaseRec.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr Then
                    myTaExternalService = CType(MyTaBaseRec, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                    Exit For
                End If
            Next

            myTaExternalService.lCopies = 0
            myTaExternalService.bPrintReceipt = False

            transactionIdentifier = String.Empty
            transactionAmount = IIf(myTaExternalService.ExistField("lAmount"), CInt(myTaExternalService.GetPropertybyName("lAmount")), String.Empty)
            transactionIdentifier = IIf(myTaExternalService.ExistField("szTransactionID"), myTaExternalService.GetPropertybyName("szTransactionID"), String.Empty)

            If transactionAmount < 0 Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            response.ReturnCode = ArgenteaCOMObject.VoidBPE(transactionAmount, transactionIdentifier, returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPEVoid

            paramArg.Copies = paramArg.BPECopies
            paramArg.PrintWithinTA = False

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & response.ReturnCode.ToString)
        End Try
        Return response
    End Function

    Public Function Closure(ByRef ArgenteaCOMObject As ARGLIB.argpay, ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "Close"
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea close function")
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO

            ' close
            response.ReturnCode = ArgenteaCOMObject.ChiusuraBPE(returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)


            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPEClosure

            paramArg.Copies = paramArg.BPEClosureCopies
            paramArg.PrintWithinTA = False

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & response.ReturnCode.ToString)
        End Try
        Return response
    End Function

    Public Function Totals(ByRef ArgenteaCOMObject As ARGLIB.argpay, ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "GetTotals"
        'Dim I, J As Integer
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea get totals function")
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            response.ReturnCode = ArgenteaCOMObject.TotaliBPE(returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.BPETotals

            paramArg.Copies = paramArg.BPETotalsCopies
            paramArg.PrintWithinTA = False

        Catch ex As Exception
            Try
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & response.ReturnCode.ToString)
        End Try
        Return response
    End Function


    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

End Class
