Imports System
Imports TPDotnet.Pos
Imports System.Xml.Linq
Imports System.Xml.XPath
Imports Microsoft.VisualBasic
'If DEBUG Then
'
'Imports ARGLIB = TPDotnet.IT.Common.Pos.EFT.PAGAMENTOLib_TESTOFFLINE
'#Else
Imports ARGLIB = PAGAMENTOLib
Public Class ADVController

#Region "Argentea specific"
    Protected ArgenteaCOMObject As ARGLIB.argpay
#End Region

#Region "Instance related functions"

    Protected Shared clsInstance As ADVController = Nothing ' internal member of instance

    Public Shared ReadOnly Property Instance() As ADVController
        Get
            If clsInstance Is Nothing Then clsInstance = New ADVController()
            Return clsInstance
        End Get
    End Property

    Protected Sub New()
        ArgenteaCOMObject = New ARGLIB.argpay()
        Parameters = New ArgenteaParameters()
    End Sub

#End Region

#Region "Parameters"

    Protected Parameters As ArgenteaParameters

#End Region

    Public Enum PaymentAdv
        JIFFY = 1
        SATISPAY = 2
        BITCOIN = 3
    End Enum

    Public Function Init(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        Init = False
        Dim funcName As String = "Init"

        Try
            ' read all parameters
            Parameters.LoadParameters(TheModCntr)

        Catch ex As Exception

        End Try

    End Function

    Public Function Payment(ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec,
                            ByRef MyTaMediaMemberDetailRec As TPDotnet.Pos.TaMediaMemberDetailRec
                            ) As Boolean
        Payment = False
        Dim funcName As String = "PaymentADV"
        Dim transactionIdentifier As String = String.Empty
        Dim returnString As String = String.Empty
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim taArgenteaEMVRec As TaArgenteaEMVRec = Nothing
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim frm As System.Windows.Forms.Form = Nothing

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea paymentADV function")

            Parameters.LoadParameters(TheModCntr)

            ' open form
            FormHelper.ShowWaitScreen(TheModCntr, False, frm)

            ' get the transaction identifier
            'transactionIdentifier = GetTransactionIdentifier(taobj)

            ' pay
            Dim szCustomerID As String = String.Empty

            If taobj.getCustRecNr <> -1 Then
                Dim cust As New CUST
                taobj.getCustInfos(cust)
                szCustomerID = cust.szCustomerID
            End If
            Dim OpType As Integer = DirectCast([Enum].Parse(GetType(PaymentAdv), MyTaMediaRec.PAYMENTinMedia.szExternalID), PaymentAdv)
            If ArgenteaCOMObject.PagamentoADV(OpType, CInt(MyTaMediaRec.dTaPaidTotal * 100), szCustomerID, "", "", returnString) <> ArgenteaFunctionsReturnCode.OK Then
                Exit Function
            End If

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            'store in registry the values of transaction and amount for future void of the EFT transaction 

            ' check CSV
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.ADVPayment, argenteaFunctionReturnObject, "-")) Then
                Exit Function
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea
            taArgenteaEMVRec = objTPTAHelperArgentea.ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(taobj, argenteaFunctionReturnObject(0))
            If taArgenteaEMVRec Is Nothing Then
                ' error
            End If
            RegistryHelper.SetLastPaymentADVTransactionIdentifier(argenteaFunctionReturnObject(0).TerminalID)
            RegistryHelper.SetLastPaymentADVTransactionAmount(CInt(MyTaMediaRec.dTaPaidTotal * 100))
            RegistryHelper.SetLastPaymentADVTransactionType(OpType)


            ' to do create a completely new transaction
            eftTA = objTPTAHelperArgentea.CreateTA(taobj, TheModCntr, taArgenteaEMVRec, False)
            If eftTA Is Nothing Then
                'error
            End If

            For I As Integer = 1 To Parameters.EftReceiptCashierCopiesPayment
                Payment = objTPTAHelperArgentea.PrintReceipt(eftTA, TheModCntr)
                If Not Payment Then
                    ' error
                    If argenteaFunctionReturnObject(0).Successfull Then
                        LOG_Debug(getLocationString(funcName), "Printer failure")
                        TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                                                    getPosTxtNew(TheModCntr.contxt,
                                                    "POSLevelITCommonPrinterFailed", 0),
                                                    0,
                                                    TheModCntr,
                                                    "POSLevelITCommonPrinterFailed")
                        ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                    End If
                End If

            Next

            ' handle TP transaction
            If Not argenteaFunctionReturnObject(0).Successfull Then
                Payment = False
                Exit Function
            End If
            'If Not MyTaMediaRec.PAYMENTinMedia.szExternalID = PAYFAST Then
            '    objTPTAHelperArgentea.SwapElectronicMedia(taobj, TheModCntr, MyTaMediaRec, argenteaFunctionReturnObject(0).Acquirer)
            'End If
            taArgenteaEMVRec.theHdr.lTaCreateNmbr = 0
            taArgenteaEMVRec.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr
            taArgenteaEMVRec.bPrintReceipt = Parameters.EftPaymentReceiptWithinTA
            taobj.Add(taArgenteaEMVRec)

            ' confirm 
            ' to do : understand the following description from specification
            '   This function is called to confirm that the receipt has been correctly printed by cash counter.
            '   Function “Conferma” is only for Ingenico Telium pos and call to this function is optional.
            If ArgenteaCOMObject.Conferma(0) <> ArgenteaFunctionsReturnCode.OK Then
                ' conferma has failed but the transaction has to be considered as successfully executed
            End If

            ' to do : understand if the transaction should be considered as valid before this step.
            Payment = True

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & Payment.ToString)
            FormHelper.ShowWaitScreen(TheModCntr, True, frm)
        End Try

    End Function

    Public Function Void(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        Void = False
        Dim funcName As String = "ADVVoid"
        Dim returnString As String = String.Empty
        Dim transactionIdentifier As String = String.Empty
        Dim transactionAmount As Double = 0
        Dim transactionType As Integer = 0
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim taArgenteaEMVRec As TaArgenteaEMVRec = Nothing
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim frm As System.Windows.Forms.Form = Nothing

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea void function")

            Parameters.LoadParameters(TheModCntr)

            transactionIdentifier = RegistryHelper.GetLastPaymentADVTransactionIdentifier
            transactionAmount = RegistryHelper.GetLastPaymentADVTransactionAmount / 100
            transactionType = RegistryHelper.GetLastPaymentADVTransactionType

            If transactionIdentifier = String.Empty Or transactionAmount < 0 Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                Return False
            End If

            ' open form
            FormHelper.ShowWaitScreen(TheModCntr, False, frm)

            ' void
            Dim szCustomerID As String = String.Empty

            If taobj.getCustRecNr <> -1 Then
                Dim cust As New CUST
                taobj.getCustInfos(cust)
                szCustomerID = cust.szCustomerID
            End If
            If ArgenteaCOMObject.StornoADV(transactionType, CInt(transactionAmount * 100), szCustomerID, IIf(transactionType = CInt(PaymentAdv.JIFFY), String.Empty, transactionIdentifier), IIf(transactionType = CInt(PaymentAdv.JIFFY), transactionIdentifier, String.Empty), "", returnString) <> ArgenteaFunctionsReturnCode.OK Then
                Exit Function
            End If

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            ' check CSV
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.ADVVoid, argenteaFunctionReturnObject, "-")) Then
                Exit Function
            End If

            Dim objTPTAHelperArgentea As New TPTAHelperArgentea
            taArgenteaEMVRec = objTPTAHelperArgentea.ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(taobj, argenteaFunctionReturnObject(0))
            If taArgenteaEMVRec Is Nothing Then
                ' error
            End If

            ' to do create a completely new transaction
            eftTA = objTPTAHelperArgentea.CreateTA(taobj, TheModCntr, taArgenteaEMVRec, False)
            If eftTA Is Nothing Then
                'error
            End If

            For I As Integer = 1 To Parameters.EftReceiptCashierCopiesVoid
                Void = objTPTAHelperArgentea.PrintReceipt(eftTA, TheModCntr)
                If Not Void Then
                    ' error
                    If argenteaFunctionReturnObject(0).Successfull Then
                        LOG_Debug(getLocationString(funcName), "Printer failure")
                        TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                                                    getPosTxtNew(TheModCntr.contxt,
                                                    "POSLevelITCommonPrinterFailed", 0),
                                                    0,
                                                    TheModCntr,
                                                    "POSLevelITCommonPrinterFailed")
                        ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                    End If
                End If
            Next
            Void = objTPTAHelperArgentea.WriteTA(eftTA, TheModCntr)
            If Not Void Then
                ' error
            End If
            Dim TaBase As TPDotnet.Pos.TaBaseRec = taobj.GetTALine(taobj.sSelReceiptLine)
            Dim xel As XElement = taobj.TAtoXDocument(False, 0, False).XPathSelectElement("/TAS/NEW_TA/ARGENTEA_EMV/Hdr[lTaRefToCreateNmbr=" + TaBase.theHdr.lTaCreateNmbr.ToString + "]/lTaSeqNmbr")

            If xel IsNot Nothing Then
                Dim lTaSeqNmbr As Integer = CInt(xel.Value)
                Dim taArg As TaArgenteaEMVRec = CType(taobj.GetTALine(lTaSeqNmbr), TaArgenteaEMVRec)
                If Not taArg.ExistField("bIsVoided") Then taArg.AddField("bIsVoided", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                taArg.setPropertybyName("bIsVoided", -1)
            End If
            Void = True

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & Void.ToString)
            FormHelper.ShowWaitScreen(TheModCntr, True, frm)
        End Try

    End Function

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
End Class
