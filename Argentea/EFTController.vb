Imports System
Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
'If DEBUG Then
'
'Imports ARGLIB = TPDotnet.IT.Common.Pos.EFT.PAGAMENTOLib_TESTOFFLINE
'#Else
Imports ARGLIB = PAGAMENTOLib
'#End If

Public Class EFTController

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


#Region "Public functions"
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

            transactionIdentifier = GetTransactionIdentifier(taobj)

            response.ReturnCode = ArgenteaCOMObject.PagamentoPlus(0, CInt(MyTaMediaRec.dTaPaidTotal * 100), 30, transactionIdentifier, returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.EFTPayment
            response.TransactionID = transactionIdentifier

            paramArg.Copies = paramArg.EftReceiptCashierCopiesPayment
            paramArg.PrintWithinTA = paramArg.EftPaymentReceiptWithinTA

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

            transactionIdentifier = IIf(myTaExternalService.ExistField("szTransactionID"), myTaExternalService.GetPropertybyName("szTransactionID"), String.Empty)
            transactionAmount = IIf(myTaExternalService.ExistField("lAmount"), CInt(myTaExternalService.GetPropertybyName("lAmount")), String.Empty)
            If transactionIdentifier = String.Empty Or transactionAmount < 0 Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            response.ReturnCode = ArgenteaCOMObject.StornoPlus(0, CInt(transactionAmount), 30, transactionIdentifier, returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.EFTVoid

            paramArg.Copies = paramArg.EftReceiptCashierCopiesVoid

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & response.ReturnCode.ToString)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
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
            response.ReturnCode = ArgenteaCOMObject.Chiusura(0, returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)


            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.EFTClose

            paramArg.Copies = paramArg.EftReceiptCashierCopiesClose
            paramArg.PrintWithinTA = paramArg.EftPaymentReceiptWithinTA

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
            response.ReturnCode = ArgenteaCOMObject.RichiestaTotaliHost(0, 0, returnString)

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.EFTGetTotals

            paramArg.Copies = paramArg.EftReceiptCashierCopiesTotals
            paramArg.PrintWithinTA = paramArg.EftPaymentReceiptWithinTA

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


#End Region

#Region "Protected functions"

    Protected Function GetTransactionIdentifier(ByRef taobj As TPDotnet.Pos.TA) As String
        GetTransactionIdentifier = String.Empty
        Dim funcName As String = "GetTransactionIdentifier"

        Try

            GetTransactionIdentifier =
                    New String(" ", 5) _
                    + "|" _
                    + Microsoft.VisualBasic.Format(taobj.lRetailStoreID, "00000") + Microsoft.VisualBasic.Format(taobj.lWorkStationNmbr, "00000") _
                    + "|" _
                    + Microsoft.VisualBasic.Format(taobj.lActOperatorID, "00000000") _
                    + "|" _
                    + taobj.szStartTaTime.Substring(0, 8) _
                    + "|" _
                    + taobj.szStartTaTime.Substring(8, 6) _
                    + "|" _
                    + Microsoft.VisualBasic.Format(taobj.lactTaNmbr, "00000")

        Catch ex As Exception

        End Try

    End Function

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
