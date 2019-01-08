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
    '''     Chiamata alle API di Argentea per l'Inizializazzione di un chiamata di dematerializzazione o storno
    ''' </summary>
    Initialization_AG

    ''' <summary>
    '''     Chiamata alle API di Argentea per il Reset del conteggio remoto delle chiamate
    ''' </summary>
    ResetCounter_AG

    ''' <summary>
    '''     Chiamata alle API di Argentea per alcuni tipi di dematerializzazione che richiedono un ultriore conferma alla chiamata
    ''' </summary>
    Confirmation_AG

    ''' <summary>
    '''     Chiamata alle API di Argentea per interrogare il servizio su un codice (coupon/buono pasto) per controllare se valido
    ''' </summary>
    Check_BP

    ''' <summary>
    '''     Protocollo in riposta a Pagamento su Servizio remoto Argentea
    ''' </summary>
    SinglePaid_BP

    ''' <summary>
    '''     Protocollo in risposta a Dematerializzazione come Pagato su POS locale fornito da Argentea
    '''     Multi elementi di risposta
    ''' </summary>
    MultiPaid_BP

    ''' <summary>
    '''     Protocollo in riposta a Storno su Servizio remoto Argentea
    '''     Singolo elemento di risposta.
    ''' </summary>
    SingleVoid_BP

    ''' <summary>
    '''     Protocollo in risposta a Storno su POS locale fornito da Argentea
    '''     Multi elementi di risposta
    ''' </summary>
    MultiVoid_BP

    ''' <summary>
    '''     Protocollo in risposta a Chiamata verso terminale POS Hardware POS locale fornito da Argentea
    '''     Multi elementi di risposta
    ''' </summary>
    MultiItemsIC_BP

    ''' <summary>
    '''     Chiamata alle API di Argentea per il Close finale su un inizio di comunicazione per il demat remoto 
    ''' </summary>
    Close_AG


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

