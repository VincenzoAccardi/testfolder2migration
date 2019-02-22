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

#Region "Argentea specific"
    Protected ArgenteaCOMObject As ARGLIB.argpay
#End Region

#Region "Instance related functions"

    Protected Shared clsInstance As EFTController = Nothing ' internal member of instance

    Public Shared ReadOnly Property Instance() As EFTController
        Get
            If clsInstance Is Nothing Then clsInstance = New EFTController()
            Return clsInstance
        End Get
    End Property

    Protected Sub New()
        ArgenteaCOMObject = New ARGLIB.argpay()
        Parameters = New ArgenteaParameters()
    End Sub

#End Region
    Protected Const PAYFAST As String = "PAYFAST"
#Region "Protected functions"

    Protected Function GetTransactionIdentifier(ByRef taobj As TPDotnet.Pos.TA) As String
        GetTransactionIdentifier = String.Empty
        Dim funcName As String = "GetTransactionIdentifier"

        Try

            GetTransactionIdentifier = _
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

#End Region

#Region "Parameters"

    Protected Parameters As ArgenteaParameters

#End Region

#Region "Public functions"

    Public Function Init(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        Init = False
        Dim funcName As String = "Init"

        Try
            ' read all parameters
            Parameters.LoadParameters(TheModCntr)

        Catch ex As Exception

        End Try

    End Function

    
    'Public Function Payment(ByRef taobj As TPDotnet.Pos.TA,
    '                        ByRef TheModCntr As TPDotnet.Pos.ModCntr,
    '                        ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec,
    '                        ByRef MyTaMediaMemberDetailRec As TPDotnet.Pos.TaMediaMemberDetailRec
    '                        ) As Boolean
    '    Payment = False
    '    Dim funcName As String = "Payment"
    '    Dim transactionIdentifier As String = String.Empty
    '    Dim returnString As String = String.Empty
    '    Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
    '    Dim taArgenteaEMVRec As TaArgenteaEMVRec = Nothing
    '    Dim eftTA As TPDotnet.Pos.TA = Nothing
    '    Dim frm As System.Windows.Forms.Form = Nothing

    '    Try
    '        LOG_Debug(getLocationString(funcName), "We are entered in Argentea payment function")

    '        Parameters.LoadParameters(TheModCntr)

    '        ' open form
    '        FormHelper.ShowWaitScreen(TheModCntr, False, frm)

    '        ' get the transaction identifier
    '        transactionIdentifier = GetTransactionIdentifier(taobj)

    '        ' pay
    '        If MyTaMediaRec.PAYMENTinMedia.szExternalID = PAYFAST Then
    '            Dim FieldName As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "FIELD_NAME_AUDIT").Trim
    '            Dim hdr As TPDotnet.Pos.TaHdrRec = CType(taobj.GetTALine(1), TPDotnet.Pos.TaHdrRec)
    '            Dim cust As TPDotnet.Pos.TaCustomerRec = CType(taobj.GetTALine(taobj.getCustRecNr), TPDotnet.Pos.TaCustomerRec)
    '            Dim AuditID As String = IIf(hdr.ExistField(FieldName), hdr.GetPropertybyName(FieldName), "0")
    '            Dim szCustomerID As String = String.Empty
    '            If cust IsNot Nothing Then szCustomerID = cust.CUSTinCustomer.szCustomerID 
    '            If (AuditID <> "0") AndAlso Not String.IsNullOrEmpty(szCustomerID) Then
    '                Dim TransID As String = _
    '                    Microsoft.VisualBasic.Format(taobj.lRetailStoreID, "0000000") + _
    '                    Microsoft.VisualBasic.Format(taobj.lWorkStationNmbr, "000") + _
    '                Microsoft.VisualBasic.Format(taobj.lactTaNmbr, "000000")
    '                LOG_Error(getLocationString(funcName), "Import: " + (CInt(MyTaMediaRec.dTaPaidTotal * 100)).ToString)
    '                LOG_Error(getLocationString(funcName), "TransID: " + TransID)
    '                LOG_Error(getLocationString(funcName), "szCustomerID: " + szCustomerID)
    '                LOG_Error(getLocationString(funcName), "AuditID: " + AuditID)
    '                If ArgenteaCOMObject.PagamentoDTP(CInt(MyTaMediaRec.dTaPaidTotal * 100), TransID, szCustomerID, AuditID, returnString) <> ArgenteaFunctionsReturnCode.OK Then
    '                    Exit Function
    '                End If
    '            Else
    '                TPMsgBox(PosDef.TARMessageTypes.TPERROR,
    '                                               getPosTxtNew(TheModCntr.contxt,
    '                                               "LevelITCommonArgenteaFidelityCardError", 0),
    '                                               0,
    '                                               TheModCntr,
    '                                               "LevelITCommonArgenteaFidelityCardError")
    '                Exit Function
    '            End If
    '        Else
    '            If ArgenteaCOMObject.PagamentoPlus(0, CInt(MyTaMediaRec.dTaPaidTotal * 100), 30, transactionIdentifier, returnString) <> ArgenteaFunctionsReturnCode.OK Then
    '                Exit Function
    '            End If
    '        End If


    '        LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

    '        'store in registry the values of transaction and amount for future void of the EFT transaction 
    '        RegistryHelper.SetLastPaymentTransactionIdentifier(transactionIdentifier)
    '        RegistryHelper.SetLastPaymentTransactionAmount(CInt(MyTaMediaRec.dTaPaidTotal * 100))

    '        ' check CSV
    '        argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
    '        If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.EFTPayment, argenteaFunctionReturnObject)) Then
    '            Exit Function
    '        End If

    '        Dim objTPTAHelperArgentea As New TPTAHelperArgentea
    '        taArgenteaEMVRec = objTPTAHelperArgentea.ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(taobj, argenteaFunctionReturnObject(0))
    '        If taArgenteaEMVRec Is Nothing Then
    '            ' error
    '        End If

    '        ' to do create a completely new transaction
    '        eftTA = objTPTAHelperArgentea.CreateTA(taobj, TheModCntr, taArgenteaEMVRec, False)
    '        If eftTA Is Nothing Then
    '            'error
    '        End If

    '        For I As Integer = 1 To Parameters.EftReceiptCashierCopiesPayment
    '            Payment = objTPTAHelperArgentea.PrintReceipt(eftTA, TheModCntr)
    '            If Not Payment Then
    '                ' error
    '                If argenteaFunctionReturnObject(0).Successfull Then
    '                    LOG_Debug(getLocationString(funcName), "Printer failure")
    '                    TPMsgBox(PosDef.TARMessageTypes.TPERROR,
    '                                                getPosTxtNew(TheModCntr.contxt,
    '                                                "POSLevelITCommonPrinterFailed", 0),
    '                                                0,
    '                                                TheModCntr,
    '                                                "POSLevelITCommonPrinterFailed")
    '                    ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
    '                End If
    '            End If

    '        Next

    '        ' handle TP transaction
    '        If Not argenteaFunctionReturnObject(0).Successfull Then
    '            Payment = False
    '            Exit Function
    '        End If
    '        If Not MyTaMediaRec.PAYMENTinMedia.szExternalID = PAYFAST Then
    '            objTPTAHelperArgentea.SwapElectronicMedia(taobj, TheModCntr, MyTaMediaRec, argenteaFunctionReturnObject(0).Acquirer)
    '        End If
    '        taArgenteaEMVRec.theHdr.lTaCreateNmbr = 0
    '        taArgenteaEMVRec.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr
    '        taArgenteaEMVRec.bPrintReceipt = Parameters.EftPaymentReceiptWithinTA
    '        taobj.Add(taArgenteaEMVRec)

    '        ' confirm 
    '        ' to do : understand the following description from specification
    '        '   This function is called to confirm that the receipt has been correctly printed by cash counter.
    '        '   Function “Conferma” is only for Ingenico Telium pos and call to this function is optional.
    '        If ArgenteaCOMObject.Conferma(0) <> ArgenteaFunctionsReturnCode.OK Then
    '            ' conferma has failed but the transaction has to be considered as successfully executed
    '        End If

    '        ' to do : understand if the transaction should be considered as valid before this step.
    '        Payment = True

    '    Catch ex As Exception
    '        Try
    '            LOG_Error(getLocationString(funcName), ex)
    '        Catch InnerEx As Exception
    '            LOG_ErrorInTry(getLocationString(funcName), InnerEx)
    '        End Try
    '    Finally
    '        LOG_FuncExit(getLocationString(funcName), "Function returns " & Payment.ToString)
    '        FormHelper.ShowWaitScreen(TheModCntr, True, frm)
    '    End Try

    'End Function

    Public Function Payment(ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec,
                            ByRef MyTaMediaMemberDetailRec As TPDotnet.Pos.TaMediaMemberDetailRec
                            ) As Boolean
        Payment = False
        Dim funcName As String = "Payment"
        Dim transactionIdentifier As String = String.Empty
        Dim returnString As String = String.Empty
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim taArgenteaEMVRec As TaExternalServiceRec = Nothing
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim frm As System.Windows.Forms.Form = Nothing

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea payment function")

            Parameters.LoadParameters(TheModCntr)

            ' open form
            FormHelper.ShowWaitScreen(TheModCntr, False, frm)

            ' get the transaction identifier
            transactionIdentifier = GetTransactionIdentifier(taobj)

            ' pay
            If MyTaMediaRec.PAYMENTinMedia.szExternalID = PAYFAST Then
                Dim FieldName As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "FIELD_NAME_AUDIT").Trim
                Dim hdr As TPDotnet.Pos.TaHdrRec = CType(taobj.GetTALine(1), TPDotnet.Pos.TaHdrRec)
                Dim cust As TPDotnet.Pos.TaCustomerRec = CType(taobj.GetTALine(taobj.getCustRecNr), TPDotnet.Pos.TaCustomerRec)
                Dim AuditID As String = IIf(hdr.ExistField(FieldName), hdr.GetPropertybyName(FieldName), "0")
                Dim szCustomerID As String = String.Empty
                If cust IsNot Nothing Then szCustomerID = cust.CUSTinCustomer.szCustomerID

                LOG_Error(getLocationString(funcName), "szCustomerID: " + szCustomerID)
                LOG_Error(getLocationString(funcName), "AuditID: " + AuditID)

                If (AuditID <> "0") AndAlso Not String.IsNullOrEmpty(szCustomerID) Then
                    Dim TransID As String = _
                        Microsoft.VisualBasic.Format(taobj.lRetailStoreID, "0000000") + _
                        Microsoft.VisualBasic.Format(taobj.lWorkStationNmbr, "000") + _
                    Microsoft.VisualBasic.Format(taobj.lactTaNmbr, "000000")
                    LOG_Error(getLocationString(funcName), "Import: " + (CInt(MyTaMediaRec.dTaPaidTotal * 100)).ToString)
                    LOG_Error(getLocationString(funcName), "TransID: " + TransID)
                    LOG_Error(getLocationString(funcName), "szCustomerID: " + szCustomerID)
                    LOG_Error(getLocationString(funcName), "AuditID: " + AuditID)
                    If ArgenteaCOMObject.PagamentoDTP(CInt(MyTaMediaRec.dTaPaidTotal * 100), TransID, szCustomerID, AuditID, returnString) <> ArgenteaFunctionsReturnCode.OK Then
                        Exit Function
                    End If
                Else
                    TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                                                   getPosTxtNew(TheModCntr.contxt,
                                                   "LevelITCommonArgenteaFidelityCardError", 0),
                                                   0,
                                                   TheModCntr,
                                                   "LevelITCommonArgenteaFidelityCardError")
                    Exit Function
                End If
            Else
                If ArgenteaCOMObject.PagamentoPlus(0, CInt(MyTaMediaRec.dTaPaidTotal * 100), 30, transactionIdentifier, returnString) <> ArgenteaFunctionsReturnCode.OK Then
                    Exit Function
                End If
            End If


            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            'store in registry the values of transaction and amount for future void of the EFT transaction 
            RegistryHelper.SetLastPaymentTransactionIdentifier(transactionIdentifier)
            RegistryHelper.SetLastPaymentTransactionAmount(CInt(MyTaMediaRec.dTaPaidTotal * 100))

            ' check CSV
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.EFTPayment, argenteaFunctionReturnObject)) Then
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
            If Not MyTaMediaRec.PAYMENTinMedia.szExternalID = PAYFAST Then
                objTPTAHelperArgentea.SwapElectronicMedia(taobj, TheModCntr, MyTaMediaRec, argenteaFunctionReturnObject(0).Acquirer)
            End If
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

    'Public Function Void(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec, ByRef MyTaMediaMemberDetailRec As TPDotnet.Pos.TaMediaMemberDetailRec) As Boolean
    Public Function Void(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        Void = False
        Dim funcName As String = "Void"
        Dim returnString As String = String.Empty
        Dim transactionIdentifier As String = String.Empty
        Dim transactionAmount As Double = 0
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim taArgenteaEMVRec As TaExternalServiceRec = Nothing
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim frm As System.Windows.Forms.Form = Nothing

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea void function")

            Parameters.LoadParameters(TheModCntr)

            transactionIdentifier = RegistryHelper.GetLastPaymentTransactionIdentifier
            transactionAmount = RegistryHelper.GetLastPaymentTransactionAmount / 100

            If transactionIdentifier = String.Empty Or transactionAmount < 0 Then
                LOG_Debug(getLocationString(funcName), "No Argentea transaction to void")
                Return False
            End If

            ' open form
            FormHelper.ShowWaitScreen(TheModCntr, False, frm)

            ' void
            If ArgenteaCOMObject.StornoPlus(0, CInt(transactionAmount * 100), 30, transactionIdentifier, returnString) <> ArgenteaFunctionsReturnCode.OK Then
                Exit Function
            End If

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            ' check CSV
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.EFTVoid, argenteaFunctionReturnObject)) Then
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

    Public Function Close(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        Close = False
        Dim funcName As String = "Close"
        Dim returnString As String = String.Empty
        Dim transactionAmount As Double = 0
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim taArgenteaEMVRec As TaExternalServiceRec = Nothing
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim frm As System.Windows.Forms.Form = Nothing

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea close function")

            Parameters.LoadParameters(TheModCntr)

            ' open form
            FormHelper.ShowWaitScreen(TheModCntr, False, frm)

            ' close
            If ArgenteaCOMObject.Chiusura(0, returnString) <> ArgenteaFunctionsReturnCode.OK Then
                Exit Function
            End If

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            ' check CSV
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.EFTClose, argenteaFunctionReturnObject)) Then
                Exit Function
            End If

            For J As Integer = 0 To argenteaFunctionReturnObject.GetUpperBound(0)
                Dim o As New TPTAHelperArgentea
                taArgenteaEMVRec = o.ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(taobj, argenteaFunctionReturnObject(J))
                If taArgenteaEMVRec Is Nothing Then
                    ' error
                End If
                ' to do create a completely new transaction
                eftTA = o.CreateTA(taobj, TheModCntr, taArgenteaEMVRec, False)
                For I As Integer = 1 To Parameters.EftReceiptCashierCopiesClose
                    Close = o.PrintReceipt(eftTA, TheModCntr)
                    If Not Close Then
                        If argenteaFunctionReturnObject(J).Successfull Then
                            LOG_Debug(getLocationString(funcName), "Printer failure")
                            TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                                                        getPosTxtNew(TheModCntr.contxt,
                                                        "POSLevelITCommonPrinterFailed", 0),
                                                        0,
                                                        TheModCntr,
                                                        "POSLevelITCommonPrinterFailed")
                            ' message box: attenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                        End If
                    End If
                Next
                Close = o.WriteTA(eftTA, TheModCntr)
                Close = True
            Next

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & Close.ToString)
            FormHelper.ShowWaitScreen(TheModCntr, True, frm)
        End Try

    End Function

    'Public Function GetStatus(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
    '    GetStatus = False
    '    Dim funcName As String = "GetStatus"

    '    Try

    '    Catch ex As Exception

    '    End Try

    'End Function

    Public Function GetTotals(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        GetTotals = False
        Dim funcName As String = "GetTotals"
        'Dim I, J As Integer
        Dim returnString As String = String.Empty
        Dim transactionIdentifier As String = String.Empty
        Dim transactionAmount As Double = 0
        Dim eftTA As TPDotnet.Pos.TA = Nothing
        Dim taArgenteaEMVRec As TaExternalServiceRec = Nothing
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        'Dim RichiestaTotaliHostDelegate As New RichiestaTotaliHostDelegate(AddressOf ArgenteaCOMObject.RichiestaTotaliHost)
        'Dim asyncResult As IAsyncResult
        Dim frm As System.Windows.Forms.Form = Nothing

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea get totals function")

            Parameters.LoadParameters(TheModCntr)

            ' open form
            FormHelper.ShowWaitScreen(TheModCntr, False, frm)

            If ArgenteaCOMObject.RichiestaTotaliHost(0, 0, returnString) <> ArgenteaFunctionsReturnCode.OK Then
                Exit Function
            End If

            LOG_Debug(getLocationString(funcName), "Argentea returns string: " & returnString)

            ' check CSV
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject

            If (Not CSVHelper.ParseReturnString(returnString, InternalArgenteaFunctionTypes.EFTGetTotals, argenteaFunctionReturnObject)) Then
                Exit Function
            End If

            For J As Integer = 0 To argenteaFunctionReturnObject.GetUpperBound(0)
                Dim objTPTAHelperArgentea As New TPTAHelperArgentea
                taArgenteaEMVRec = objTPTAHelperArgentea.ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(taobj, argenteaFunctionReturnObject(J))
                If taArgenteaEMVRec Is Nothing Then
                    ' error
                End If
                ' to do create a completely new transaction
                eftTA = objTPTAHelperArgentea.CreateTA(taobj, TheModCntr, taArgenteaEMVRec, False)
                For I As Integer = 1 To Parameters.EftReceiptCashierCopiesTotals
                    GetTotals = objTPTAHelperArgentea.PrintReceipt(eftTA, TheModCntr)
                    If Not GetTotals Then
                        If argenteaFunctionReturnObject(J).Successfull Then
                            LOG_Debug(getLocationString(funcName), "Printer failure")
                            TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                                                        getPosTxtNew(TheModCntr.contxt,
                                                        "POSLevelITCommonPrinterFailed", 0),
                                                        0,
                                                        TheModCntr,
                                                        "POSLevelITCommonPrinterFailed")
                            ' message box: attenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                        End If
                    End If
                Next
                GetTotals = objTPTAHelperArgentea.WriteTA(eftTA, TheModCntr)
            Next

            If Not GetTotals Then
                'error
            End If

            GetTotals = True

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & GetTotals.ToString)
            FormHelper.ShowWaitScreen(TheModCntr, True, frm)
        End Try

    End Function

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
