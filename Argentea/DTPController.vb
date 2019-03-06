Imports System
Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
Imports ARGLIB = PAGAMENTOLib
Public Class DTPController
    Public Function Payment(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "Payment"
        Dim transactionIdentifier As String = String.Empty
        Dim returnString As String = String.Empty
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim response As New ArgenteaResponse

        Try
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)

            LOG_Debug(getLocationString(funcName), "We are entered in Argentea payment function")

            Dim FieldName As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "FIELD_NAME_AUDIT").Trim
            Dim hdr As TPDotnet.Pos.TaHdrRec = CType(taobj.GetTALine(1), TPDotnet.Pos.TaHdrRec)
            Dim cust As TPDotnet.Pos.TaCustomerRec = CType(taobj.GetTALine(taobj.getCustRecNr), TPDotnet.Pos.TaCustomerRec)
            Dim szAdditionalInfo As String = IIf(hdr.ExistField(FieldName), hdr.GetPropertybyName(FieldName), "0")
            Dim szCustomerID As String = String.Empty
            If cust IsNot Nothing Then szCustomerID = cust.CUSTinCustomer.szCustomerID

            transactionIdentifier =
                Microsoft.VisualBasic.Format(taobj.lRetailStoreID, "0000000") +
                Microsoft.VisualBasic.Format(taobj.lWorkStationNmbr, "000") +
                Microsoft.VisualBasic.Format(taobj.lactTaNmbr, "000000")


            response.ReturnCode = ArgenteaCOMObject.PagamentoDTP(CInt(MyTaMediaRec.dTaPaidTotal * 100), transactionIdentifier, szCustomerID, szAdditionalInfo, returnString)
            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.EFTPayment
            response.TransactionID = transactionIdentifier

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

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
        Dim response As New ArgenteaResponse
        response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Return response
    End Function

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

End Class
