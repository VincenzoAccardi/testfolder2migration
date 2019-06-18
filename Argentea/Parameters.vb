﻿Imports System

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

Public Class ArgenteaParameters
    Inherits TPDotnet.IT.Common.Pos.CommonParameters
#Region "EFT"

    Private _EftOperationTimeoutInSeconds As Integer
    <TPParametersAttribute("EFT_OPERATION_TIMEOUT_IN_SEC", 60)>
    Public Property EftOperationTimeoutInSeconds() As Integer
        Get
            Return _EftOperationTimeoutInSeconds
        End Get
        Set(ByVal value As Integer)
            _EftOperationTimeoutInSeconds = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesPayment As Integer
    <TPParametersAttribute("NUM_RECEIPT_TO_PRINT_PAYMENT", 1)>
    Public Property EftReceiptCashierCopiesPayment() As Integer
        Get
            Return _EftReceiptCashierCopiesPayment
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesPayment = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesVoid As Integer
    <TPParametersAttribute("NUM_RECEIPT_TO_PRINT_VOID", 1)>
    Public Property EftReceiptCashierCopiesVoid() As Integer
        Get
            Return _EftReceiptCashierCopiesVoid
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesVoid = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesClose As Integer
    <TPParametersAttribute("NUM_RECEIPT_TO_PRINT_CLOSE", 1)>
    Public Property EftReceiptCashierCopiesClose() As Integer
        Get
            Return _EftReceiptCashierCopiesClose
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesClose = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesTotals As Integer
    <TPParametersAttribute("NUM_RECEIPT_TO_PRINT_TOTALS", 1)>
    Public Property EftReceiptCashierCopiesTotals() As Integer
        Get
            Return _EftReceiptCashierCopiesTotals
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesTotals = value
        End Set
    End Property

    Private _EftPaymentReceiptWithinTA As Boolean
    <TPParametersAttribute("EFT_PAYMENT_RECEIPT_WITHIN_TA", True)>
    Public Property EftPaymentReceiptWithinTA() As Boolean
        Get
            Return _EftPaymentReceiptWithinTA
        End Get
        Set(ByVal value As Boolean)
            _EftPaymentReceiptWithinTA = value
        End Set
    End Property

    Private _EftChiusuraLegacy As Integer
    <TPParametersAttribute("EFT_CHIUSURA_LEGACY", 0)>
    Public Property EftChiusuraLegacy() As Integer
        Get
            Return _EftChiusuraLegacy
        End Get
        Set(ByVal value As Integer)
            _EftChiusuraLegacy = value
        End Set
    End Property

#End Region

#Region "GiftCard"

#Region "Balance"

    Private _GiftCardBalanceLineIdentifier As String
    <TPParametersAttribute("GC_BALANCE_LINEIDENTIFIER", "Saldo EUR ")>
    Public Property GiftCardBalanceLineIdentifier() As String
        Get
            Return _GiftCardBalanceLineIdentifier
        End Get
        Set(ByVal value As String)
            _GiftCardBalanceLineIdentifier = value
        End Set
    End Property

    Private _GiftCardBalanceCopies As String
    <TPParametersAttribute("GC_BALANCE_COPIES", 1)>
    Public Property GiftCardBalanceCopies() As String
        Get
            Return _GiftCardBalanceCopies
        End Get
        Set(ByVal value As String)
            _GiftCardBalanceCopies = value
        End Set
    End Property

    Private _GiftCardBalanceSave As Boolean
    <TPParametersAttribute("GC_BALANCE_SAVE", True)>
    Public Property GiftCardBalanceSave() As Boolean
        Get
            Return _GiftCardBalanceSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardBalanceSave = value
        End Set
    End Property

    Private _GiftCardBalancePrintWithinTa As Boolean
    <TPParametersAttribute("GC_BALANCE_PRINT_WITHIN_TA", True)>
    Public Property GiftCardBalancePrintWithinTa() As Boolean
        Get
            Return _GiftCardBalancePrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _GiftCardBalancePrintWithinTa = value
        End Set
    End Property

    Private _GiftCardBalanceInternalInquiry As Boolean
    Public Property GiftCardBalanceInternalInquiry() As Boolean
        Get
            Return _GiftCardBalanceInternalInquiry
        End Get
        Set(ByVal value As Boolean)
            _GiftCardBalanceInternalInquiry = value
        End Set
    End Property

#End Region

#Region "Activation"

    Private _GiftCardActivationCopies As Integer
    <TPParametersAttribute("GC_ACTIVATION_COPIES", 1)>
    Public Property GiftCardActivationCopies() As Integer
        Get
            Return _GiftCardActivationCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardActivationCopies = value
        End Set
    End Property

    Private _GiftCardActivationSave As Boolean
    <TPParametersAttribute("GC_ACTIVATION_SAVE", False)>
    Public Property GiftCardActivationSave() As Boolean
        Get
            Return _GiftCardActivationSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardActivationSave = value
        End Set
    End Property

    Private _GiftCardActivationPrintWithinTa As Boolean
    <TPParametersAttribute("GC_ACTIVATION_PRINT_WITHIN_TA", True)>
    Public Property GiftCardActivationPrintWtihinTa() As Boolean
        Get
            Return _GiftCardActivationPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _GiftCardActivationPrintWithinTa = value
        End Set
    End Property

#End Region

#Region "ActivationCheck"

    Private _GiftCardActivationCheckCopies As Integer
    <TPParametersAttribute("GC_ACTIVATION_CHECK_COPIES", 1)>
    Public Property GiftCardActivationCheckCopies() As Integer
        Get
            Return _GiftCardActivationCheckCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardActivationCheckCopies = value
        End Set
    End Property

    Private _GiftCardActivationCheckSave As Boolean
    <TPParametersAttribute("GC_ACTIVATION_CHECK_SAVE", False)>
    Public Property GiftCardActivationCheckSave() As Boolean
        Get
            Return _GiftCardActivationCheckSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardActivationCheckSave = value
        End Set
    End Property

#End Region

#Region "Redeem"

    Private _GiftCardRedeemCopies As Integer
    <TPParametersAttribute("GC_REDEEM_COPIES", 1)>
    Public Property GiftCardRedeemCopies() As Integer
        Get
            Return _GiftCardRedeemCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardRedeemCopies = value
        End Set
    End Property

    Private _GiftCardRedeemSave As Boolean
    <TPParametersAttribute("GC_REDEEM_SAVE", False)>
    Public Property GiftCardRedeemSave() As Boolean
        Get
            Return _GiftCardRedeemSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardRedeemSave = value
        End Set
    End Property

    Private _GiftCardRedeemCheckPrintWithinTa As Boolean
    <TPParametersAttribute("GC_REDEEM_CHECK_PRINT_WITHIN_TA", True)>
    Public Property GiftCardRedeemCheckPrintWithinTa() As Boolean
        Get
            Return _GiftCardRedeemCheckPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _GiftCardRedeemCheckPrintWithinTa = value
        End Set
    End Property


#End Region

#Region "RedeemCheck"

    Private _GiftCardRedeemCheckCopies As Integer
    <TPParametersAttribute("GC_REDEEM_CHECK_COPIES", 1)>
    Public Property GiftCardRedeemCheckCopies() As Integer
        Get
            Return _GiftCardRedeemCheckCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardRedeemCheckCopies = value
        End Set
    End Property

    Private _GiftCardRedeemCheckSave As Boolean
    <TPParametersAttribute("GC_REDEEM_CHECK_SAVE", False)>
    Public Property GiftCardRedeemCheckSave() As Boolean
        Get
            Return _GiftCardRedeemCheckSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardRedeemCheckSave = value
        End Set
    End Property

#End Region

#Region "RedeemCancel"

    Private _GiftCardRedeemCancelCopies As Integer
    <TPParametersAttribute("GC_REDEEM_CANCEL_COPIES", 1)>
    Public Property GiftCardRedeemCancelCopies() As Integer
        Get
            Return _GiftCardRedeemCancelCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardRedeemCancelCopies = value
        End Set
    End Property

    Private _GiftCardRedeemCancelSave As Boolean
    <TPParametersAttribute("GC_REDEEM_CANCEL_SAVE", False)>
    Public Property GiftCardRedeemCancelSave() As Boolean
        Get
            Return _GiftCardRedeemCancelSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardRedeemCancelSave = value
        End Set
    End Property

#End Region

#End Region

#Region "ExtGiftCard"
#Region "Activation"

    Private _ExtGiftCardActivationCopies As Integer
    <TPParametersAttribute("EXT_GC_ACTIVATION_COPIES", 1)>
    Public Property ExtGiftCardActivationCopies() As Integer
        Get
            Return _ExtGiftCardActivationCopies
        End Get
        Set(ByVal value As Integer)
            _ExtGiftCardActivationCopies = value
        End Set
    End Property

    Private _ExtGiftCardActivationSave As Boolean
    <TPParametersAttribute("EXT_GC_ACTIVATION_SAVE", False)>
    Public Property ExtGiftCardActivationSave() As Boolean
        Get
            Return _ExtGiftCardActivationSave
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardActivationSave = value
        End Set
    End Property

    Private _ExtGiftCardActivationPrintWithinTa As Boolean
    <TPParametersAttribute("EXT_GC_ACTIVATION_PRINT_WITHIN_TA", True)>
    Public Property ExtGiftCardActivationPrintWithinTa() As Boolean
        Get
            Return _ExtGiftCardActivationPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardActivationPrintWithinTa = value
        End Set
    End Property

#End Region
#Region "DeActivation"
    Private _ExtGiftCardDeActivationCancelCopies As Integer
    <TPParametersAttribute("EXT_GC_DEACTIVATION_CANCEL_COPIES", 1)>
    Public Property ExtGiftCardDeActivationCancelCopies() As Integer
        Get
            Return _ExtGiftCardDeActivationCancelCopies
        End Get
        Set(ByVal value As Integer)
            _ExtGiftCardDeActivationCancelCopies = value
        End Set
    End Property

    Private _ExtGiftCardDeActivationCancelSave As Boolean
    <TPParametersAttribute("EXT_GC_DEACTIVATION_CANCEL_SAVE", False)>
    Public Property ExtGiftCardDeActivationCancelSave() As Boolean
        Get
            Return _ExtGiftCardDeActivationCancelSave
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardDeActivationCancelSave = value
        End Set
    End Property

    Private _ExtGiftCardDeActivationPrintWithinTa As Boolean
    <TPParametersAttribute("EXT_GC_DEACTIVATION_PRINT_WITHIN_TA", True)>
    Public Property ExtGiftCardDeActivationPrintWithinTa() As Boolean
        Get
            Return _ExtGiftCardDeActivationPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardDeActivationPrintWithinTa = value
        End Set
    End Property

#End Region
#Region "Confirm"
    Private _ExtGiftCardConfirmCopies As Integer
    <TPParametersAttribute("EXT_GC_CONFIRM_COPIES", 1)>
    Public Property ExtGiftCardConfirmCopies() As Integer
        Get
            Return _ExtGiftCardConfirmCopies
        End Get
        Set(ByVal value As Integer)
            _ExtGiftCardConfirmCopies = value
        End Set
    End Property

    Private _ExtGiftCardConfirmSave As Boolean
    <TPParametersAttribute("EXT_GC_CONFIRM_SAVE", False)>
    Public Property ExtGiftCardConfirmSave() As Boolean
        Get
            Return _ExtGiftCardConfirmSave
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardConfirmSave = value
        End Set
    End Property

    Private _ExtGiftCardConfirmPrintWithinTa As Boolean
    <TPParametersAttribute("EXT_GC_CONFIRM_PRINT_WITHIN_TA", True)>
    Public Property ExtGiftCardConfirmPrintWithinTa() As Boolean
        Get
            Return _ExtGiftCardConfirmPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardConfirmPrintWithinTa = value
        End Set
    End Property
#End Region

#End Region
#Region "PhoneCard"

#Region "CheckPhoneRecharge"

    Private _PhoneRechargeCheckCopies As Integer
    <TPParametersAttribute("PC_RECHARGE_CHECK_COPIES", 0)>
    Public Property PhoneRechargeCheckCopies() As Integer
        Get
            Return _PhoneRechargeCheckCopies
        End Get
        Set(ByVal value As Integer)
            _PhoneRechargeCheckCopies = value
        End Set
    End Property

    Private _PhoneRechargeCheckSave As Boolean
    <TPParametersAttribute("PC_RECHARGE_CHECK_SAVE", False)>
    Public Property PhoneRechargeCheckSave() As Boolean
        Get
            Return _PhoneRechargeCheckSave
        End Get
        Set(ByVal value As Boolean)
            _PhoneRechargeCheckSave = value
        End Set
    End Property

    Private _PhoneRechargeCheckPrintWithinTa As Boolean
    <TPParametersAttribute("PC_CHECK_PRINT_WITHIN_TA", False)>
    Public Property PhoneRechargeCheckPrintWithinTa() As Boolean
        Get
            Return _PhoneRechargeCheckPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _PhoneRechargeCheckPrintWithinTa = value
        End Set
    End Property

#End Region

#Region "ActivatePhoneRecharge"

    Private _PhoneRechargeActivationCopies As Integer
    <TPParametersAttribute("PC_ACTIVATION_COPIES", 1)>
    Public Property PhoneRechargeActivationCopies() As Integer
        Get
            Return _PhoneRechargeActivationCopies
        End Get
        Set(ByVal value As Integer)
            _PhoneRechargeActivationCopies = value
        End Set
    End Property

    Private _PhoneRechargeActivationSave As Boolean
    <TPParametersAttribute("PC_ACTIVATION_CHECK_SAVE", False)>
    Public Property PhoneRechargeActivationSave() As Boolean
        Get
            Return _PhoneRechargeActivationSave
        End Get
        Set(ByVal value As Boolean)
            _PhoneRechargeActivationSave = value
        End Set
    End Property

    Private _PhoneRechargeActivationPrintWithinTa As Boolean
    <TPParametersAttribute("PC_ACTIVATION_PRINT_WITHIN_TA", True)>
    Public Property PhoneRechargeActivationPrintWithinTa() As Boolean
        Get
            Return _PhoneRechargeActivationPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _PhoneRechargeActivationPrintWithinTa = value
        End Set
    End Property

#End Region


#End Region

#Region "BuoniPasto"

    Private _BPRupp As String
    <TPParametersAttribute("BP_ParameterRuppArgentea", "")>
    Public Property BPRupp() As String
        Get
            Return _BPRupp
        End Get
        Set(ByVal value As String)
            _BPRupp = value
        End Set
    End Property


    Private _BP_AcceptExcedeedValues As Boolean
    <TPParametersAttribute("BP_AcceptExcedeedValues", False)>
    Public Property BP_AcceptExcedeedValues() As Boolean
        Get
            Return _BP_AcceptExcedeedValues
        End Get
        Set(ByVal value As Boolean)
            _BP_AcceptExcedeedValues = value
        End Set
    End Property

    Private _BP_AccorpateOnTA As Boolean
    <TPParametersAttribute("BP_AccorpateOnTA", False)>
    Public Property BP_AccorpateOnTA() As Boolean
        Get
            Return _BP_AccorpateOnTA
        End Get
        Set(ByVal value As Boolean)
            _BP_AccorpateOnTA = value
        End Set
    End Property

    Private _BP_NumMaxPayablesOnVoid As Integer
    <TPParametersAttribute("BP_NumMaxPayablesOnVoid", 0)>
    Public Property BP_NumMaxPayablesOnVoid() As Integer
        Get
            Return _BP_NumMaxPayablesOnVoid
        End Get
        Set(ByVal value As Integer)
            _BP_NumMaxPayablesOnVoid = value
        End Set
    End Property

    Private _BP_MaxBPPayableSomeSession As Integer
    <TPParametersAttribute("BP_MaxBPPayableSomeSession", 0)>
    Public Property BP_MaxBPPayableSomeSession() As Integer
        Get
            Return _BP_MaxBPPayableSomeSession
        End Get
        Set(ByVal value As Integer)
            _BP_MaxBPPayableSomeSession = value
        End Set
    End Property

#Region "BPE"
    Private _BPECopies As Integer
    <TPParametersAttribute("BPE_COPIES", 1)>
    Public Property BPECopies() As Integer
        Get
            Return _BPECopies
        End Get
        Set(ByVal value As Integer)
            _BPECopies = value
        End Set
    End Property


    Private _BPEPrintWithinTa As Boolean
    <TPParametersAttribute("BPE_PRINT_WITHIN_TA", True)>
    Public Property BPEPrintWithinTa() As Boolean
        Get
            Return _BPEPrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _BPEPrintWithinTa = value
        End Set
    End Property

    Private _BPEClosureCopies As Integer
    <TPParametersAttribute("BPE_CLOSURE_COPIES", 1)>
    Public Property BPEClosureCopies() As Integer
        Get
            Return _BPEClosureCopies
        End Get
        Set(ByVal value As Integer)
            _BPEClosureCopies = value
        End Set
    End Property

    Private _BPETotalsCopies As Integer
    <TPParametersAttribute("BPE_TOTALS_COPIES", 1)>
    Public Property BPETotalsCopies() As Integer
        Get
            Return _BPETotalsCopies
        End Get
        Set(ByVal value As Integer)
            _BPETotalsCopies = value
        End Set
    End Property

#End Region

#Region "BPCeliac"

#End Region

#End Region

#Region "Common"
    Private _PrintLogoOnExternalReceipts As Boolean
    <TPParametersAttribute("PRINT_LOGO_ON_EXTERNAL_RECEIPTS", True)>
    Public Property PrintLogoOnExternalReceipts() As Boolean
        Get
            Return _PrintLogoOnExternalReceipts
        End Get
        Set(ByVal value As Boolean)
            _PrintLogoOnExternalReceipts = value
        End Set
    End Property

#End Region

    Public Overloads Function LoadParametersByReflection(ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        LoadParametersByReflection = False
        Dim funcName As String = "LoadParametersByReflection"
        Dim szTmp As String = String.Empty
        Dim attr As TPParametersAttribute

        Try

            For Each pi As System.Reflection.PropertyInfo In Me.GetType.GetProperties()

                Try
                    attr = pi.GetCustomAttributes(GetType(TPParametersAttribute), False)(0)
                Catch ex As Exception
                    Continue For
                End Try

                szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + attr.Name)
                If String.IsNullOrEmpty(szTmp) OrElse String.IsNullOrWhiteSpace(szTmp) Then
                    ' use default
                    szTmp = attr.DefaultValue
                End If

                Try

                    If pi.PropertyType = GetType(Int32) Then

                        pi.SetValue(Me, Integer.Parse(szTmp), Nothing)

                    ElseIf pi.PropertyType = GetType(Boolean) Then

                        If szTmp.ToUpper = "Y" Then
                            pi.SetValue(Me, True, Nothing)
                        ElseIf szTmp.ToUpper = "N" Then
                            pi.SetValue(Me, False, Nothing)
                        Else
                            pi.SetValue(Me, Boolean.Parse(szTmp), Nothing)
                        End If

                    Else

                        pi.SetValue(Me, szTmp, Nothing)

                    End If

                Catch ex As Exception

                End Try

            Next pi

        Catch ex As Exception

        End Try

    End Function


    Public Function LoadParameters(ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        LoadParameters = False
        Dim funcName As String = "LoadParameters"
        Dim szTmp As String = String.Empty

        Try
            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "EFT_OPERATION_TIMEOUT")
            If String.IsNullOrEmpty(szTmp) OrElse Not Integer.TryParse(szTmp, EftOperationTimeoutInSeconds) Then
                EftOperationTimeoutInSeconds = 60
            End If

            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "NUM_RECEIPT_TO_PRINT_PAYMENT")
            If String.IsNullOrEmpty(szTmp) OrElse Not Integer.TryParse(szTmp, EftReceiptCashierCopiesPayment) Then
                EftReceiptCashierCopiesPayment = 1
            End If

            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "NUM_RECEIPT_TO_PRINT_VOID")
            If String.IsNullOrEmpty(szTmp) OrElse Not Integer.TryParse(szTmp, EftReceiptCashierCopiesVoid) Then
                EftReceiptCashierCopiesVoid = 1
            End If

            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "NUM_RECEIPT_TO_PRINT_CLOSE")
            If String.IsNullOrEmpty(szTmp) OrElse Not Integer.TryParse(szTmp, EftReceiptCashierCopiesClose) Then
                EftReceiptCashierCopiesClose = 1
            End If

            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "NUM_RECEIPT_TO_PRINT_TOTALS")
            If String.IsNullOrEmpty(szTmp) OrElse Not Integer.TryParse(szTmp, EftReceiptCashierCopiesTotals) Then
                EftReceiptCashierCopiesTotals = 1
            End If

            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "EFT_PAYMENT_RECEIPT_WITHIN_TA")
            If String.IsNullOrEmpty(szTmp) OrElse String.IsNullOrWhiteSpace(szTmp) OrElse szTmp.ToUpper = "Y" Then
                EftPaymentReceiptWithinTA = 1
            ElseIf szTmp.ToUpper = "N" Then
                EftPaymentReceiptWithinTA = 0
            End If

            szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + "GIFTCARD_BALANCE_LINE_ID")
            If Not String.IsNullOrEmpty(szTmp) AndAlso Not String.IsNullOrWhiteSpace(szTmp) Then
                GiftCardBalanceLineIdentifier = szTmp
            Else
                GiftCardBalanceLineIdentifier = "Saldo EUR " ' default value based on the development environment
            End If

            ' ...

        Catch ex As Exception

        Finally

        End Try

    End Function

End Class
