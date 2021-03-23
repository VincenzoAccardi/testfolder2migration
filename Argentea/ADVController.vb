Imports System
Imports TPDotnet.Pos
Imports TPDotnet.IT.Common.Pos
Imports System.Xml.Linq
Imports System.Xml.XPath
Imports Microsoft.VisualBasic
'If DEBUG Then
'
'Imports ARGLIB = TPDotnet.IT.Common.Pos.EFT.PAGAMENTOLib_TESTOFFLINE
'#Else
Imports ARGLIB = PAGAMENTOLib
Imports System.Collections.Generic
Imports System.Linq

Public Class ADVController


    Public Enum PaymentAdv
        JIFFY = 1
        SATISPAY = 2
        BITCOIN = 3
        SATISPAYQRC = 4
        CONADPAY = 5
        TINABA = 6
        ALIPAY = 7
        WECHAT = 8
        POSTEPAY = 9
        FIERAPAY = 10
    End Enum

    Public Function Payment(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "PaymentADV"
        Dim transactionIdentifier As String = String.Empty
        Dim returnString As String = String.Empty
        Dim response As New ArgenteaResponse

        Try
            'Init(Parameters)

            LOG_Debug(getLocationString(funcName), "We are entered in Argentea paymentADV function")

            Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
            ' pay
            Dim szCustomerID As String = String.Empty

            If taobj.getCustRecNr <> -1 Then
                Dim cust As New CUST
                taobj.getCustInfos(cust)
                szCustomerID = cust.szCustomerID
            End If
            Dim OpType As Integer = GetOperationType(MyTaMediaRec.PAYMENTinMedia.szExternalID)

            response.ReturnCode = ArgenteaCOMObject.PagamentoADV(OpType, CInt(MyTaMediaRec.dTaPaidTotal * 100), szCustomerID, "", "", returnString)
            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Minus
            response.FunctionType = InternalArgenteaFunctionTypes.ADVPayment

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            paramArg.Copies = paramArg.EftReceiptCashierCopiesPayment
            paramArg.PrintWithinTA = paramArg.EftPaymentReceiptWithinTA
            ' handle TP transaction

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

    Private Function GetOperationType(ByVal szExternalID As String) As Integer
        Dim GetService As String
        Dim funcName As String = "GetService"
        Dim availableServices As String() = {JiffyMedia, SatispayMedia, BitCoinMedia, SatispayQRCodeMedia, ConadPayMedia, TinabaMedia, AliPayMedia, WeChatMedia, PostePayMedia, FieraPayMedia}

        Try

            ' get & check the current payment service
            GetService = availableServices.OrderByDescending(Function(x) x).Where(Function(s) szExternalID.StartsWith(s.ToUpper)).FirstOrDefault
            If String.IsNullOrEmpty(GetService) Then Exit Function

            Return DirectCast([Enum].Parse(GetType(PaymentAdv), GetService), PaymentAdv)
        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "returns ")
        End Try
    End Function
    Public Function Void(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                         ByRef taobj As TPDotnet.Pos.TA,
                         ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                         ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                         ByRef MyCurrentDetailRecord As TPDotnet.Pos.TaBaseRec,
                         ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "ADVVoid"
        Dim returnString As String = String.Empty
        Dim transactionIdentifier As String = String.Empty
        Dim transactionAmount As Integer = 0
        Dim transactionType As Integer = 0
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim response As New ArgenteaResponse
        Dim myTaExternalService As New TPDotnet.IT.Common.Pos.TaExternalServiceRec


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

            transactionIdentifier = IIf(myTaExternalService.ExistField("szTransactionID"), myTaExternalService.GetPropertybyName("szTransactionID"), String.Empty)
            transactionAmount = IIf(myTaExternalService.ExistField("lAmount"), CInt(myTaExternalService.GetPropertybyName("lAmount")), String.Empty)
            transactionType = DirectCast([Enum].Parse(GetType(PaymentAdv), myTaExternalService.szServiceType), PaymentAdv)
            If transactionIdentifier = String.Empty Or transactionAmount < 0 Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                Return response
            End If

            ' void
            Dim szCustomerID As String = String.Empty

            If taobj.getCustRecNr <> -1 Then
                Dim cust As New CUST
                taobj.getCustInfos(cust)
                szCustomerID = cust.szCustomerID
            End If

            response.ReturnCode = ArgenteaCOMObject.StornoADV(transactionType, CInt(transactionAmount), szCustomerID, IIf(transactionType = CInt(PaymentAdv.JIFFY), String.Empty, transactionIdentifier), IIf(transactionType = CInt(PaymentAdv.JIFFY), transactionIdentifier, String.Empty), "", returnString)
            response.MessageOut = returnString
            response.CharSeparator = CharSeparator.Minus
            response.FunctionType = InternalArgenteaFunctionTypes.ADVVoid


            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            ' check CSV


            paramArg.Copies = paramArg.EftReceiptCashierCopiesVoid

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
    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function


End Class
