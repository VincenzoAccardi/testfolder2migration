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

Public Enum InternalArgenteaFunctionTypes
    EFTPayment
    EFTVoid
    EFTGetTotals
    EFtGetStatus
    EFTClose
    EFTConfirm
    GiftCardActivationPreCheck
    GiftCardActivation
    GiftCardBalance
    GiftCardRedeemPreCkeck
    GiftCardRedeem
    GiftCardRedeemCancel
    PhoneRechargeActivation
    PhoneRechargeCheck
    ExternalGiftCardActivation
    ExternalGiftCardDeActivation
    ExternalGiftCardConfirm
    ADVPayment
    ADVVoid
    ''' <summary>
    '''     Protocollo in riposta a Pagamento su Servizio remoto Argentea
    ''' </summary>
    BPCPayment
    ''' <summary>
    '''     Protocollo in risposta a Pagamento su POS locale fornito da Argentea
    ''' </summary>
    BPEPayment
    ''' <summary>
    '''     Protocollo in riposta a Storno su Servizio remoto Argentea
    ''' </summary>
    BPEVoid
    ''' <summary>
    '''     Protocollo in risposta a Storno su POS locale fornito da Argentea
    ''' </summary>
    BPCVoid
End Enum

Public Enum ArgenteaFunctionsReturnCode
    KO = 0
    OK = 1
End Enum

Public Enum ArgenteaFormStates
    OperationSelection
    OperationInProgress
End Enum

Public Enum ArgenteaFunctionPagamentoplus
    TerminalID
    Amount
    Result
    Description
    Acquirer
    Receipt
End Enum

Public Enum ArgenteaGiftCardStatus
    Deactivated = 0
    ActivatedWithCheckMode
    ActivatedDefinitively
    RedeemWithCheckMode
    RedeemDefinitively
    RedeemCanceled
End Enum

Public Enum ArgenteaPhoneRechargeStatus
    Deactivated = 0
    ActivatedWithCheckMode
    ActivatedDefinitively
End Enum

Public Enum ArgenteaExternalGiftCardStatus
    Deactivated = 0
    ActivatedWithCheckMode
    ActivatedDefinitively
End Enum

