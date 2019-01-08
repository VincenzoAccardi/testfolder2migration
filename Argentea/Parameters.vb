Imports System

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

Public Class ArgenteaParametersAttribute
    Inherits Attribute

    Public Name As String
    Public DefaultValue As Object

    Public Sub New(ByVal name As String, ByVal defaultValue As Object)
        Me.Name = name
        Me.DefaultValue = defaultValue
    End Sub

End Class

Public Class ArgenteaParameters
#Region "EFT"

    Private _EftOperationTimeoutInSeconds As Integer
    <ArgenteaParameters("EFT_OPERATION_TIMEOUT_IN_SEC", 60)>
    Public Property EftOperationTimeoutInSeconds() As Integer
        Get
            Return _EftOperationTimeoutInSeconds
        End Get
        Set(ByVal value As Integer)
            _EftOperationTimeoutInSeconds = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesPayment As Integer
    <ArgenteaParameters("NUM_RECEIPT_TO_PRINT_PAYMENT", 1)>
    Public Property EftReceiptCashierCopiesPayment() As Integer
        Get
            Return _EftReceiptCashierCopiesPayment
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesPayment = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesVoid As Integer
    <ArgenteaParameters("NUM_RECEIPT_TO_PRINT_VOID", 1)>
    Public Property EftReceiptCashierCopiesVoid() As Integer
        Get
            Return _EftReceiptCashierCopiesVoid
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesVoid = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesClose As Integer
    <ArgenteaParameters("NUM_RECEIPT_TO_PRINT_CLOSE", 1)>
    Public Property EftReceiptCashierCopiesClose() As Integer
        Get
            Return _EftReceiptCashierCopiesClose
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesClose = value
        End Set
    End Property

    Private _EftReceiptCashierCopiesTotals As Integer
    <ArgenteaParameters("NUM_RECEIPT_TO_PRINT_TOTALS", 1)>
    Public Property EftReceiptCashierCopiesTotals() As Integer
        Get
            Return _EftReceiptCashierCopiesTotals
        End Get
        Set(ByVal value As Integer)
            _EftReceiptCashierCopiesTotals = value
        End Set
    End Property

    Private _EftPaymentReceiptWithinTA As Boolean
    <ArgenteaParameters("EFT_PAYMENT_RECEIPT_WITHIN_TA", True)>
    Public Property EftPaymentReceiptWithinTA() As Boolean
        Get
            Return _EftPaymentReceiptWithinTA
        End Get
        Set(ByVal value As Boolean)
            _EftPaymentReceiptWithinTA = value
        End Set
    End Property

#End Region

#Region "GiftCard"

#Region "Balance"

    Private _GiftCardBalanceLineIdentifier As String
    <ArgenteaParameters("GC_BALANCE_LINEIDENTIFIER", "Saldo EUR ")>
    Public Property GiftCardBalanceLineIdentifier() As String
        Get
            Return _GiftCardBalanceLineIdentifier
        End Get
        Set(ByVal value As String)
            _GiftCardBalanceLineIdentifier = value
        End Set
    End Property

    Private _GiftCardBalanceCopies As String
    <ArgenteaParameters("GC_BALANCE_COPIES", 1)>
    Public Property GiftCardBalanceCopies() As String
        Get
            Return _GiftCardBalanceCopies
        End Get
        Set(ByVal value As String)
            _GiftCardBalanceCopies = value
        End Set
    End Property

    Private _GiftCardBalanceSave As Boolean
    <ArgenteaParameters("GC_BALANCE_SAVE", True)>
    Public Property GiftCardBalanceSave() As Boolean
        Get
            Return _GiftCardBalanceSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardBalanceSave = value
        End Set
    End Property

    Private _GiftCardBalancePrintWithinTa As Boolean
    <ArgenteaParameters("GC_BALANCE_PRINT_WITHIN_TA", True)>
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
    <ArgenteaParameters("GC_ACTIVATION_COPIES", 1)>
    Public Property GiftCardActivationCopies() As Integer
        Get
            Return _GiftCardActivationCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardActivationCopies = value
        End Set
    End Property

    Private _GiftCardActivationSave As Boolean
    <ArgenteaParameters("GC_ACTIVATION_SAVE", False)>
    Public Property GiftCardActivationSave() As Boolean
        Get
            Return _GiftCardActivationSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardActivationSave = value
        End Set
    End Property

    Private _GiftCardActivationPrintWithinTa As Boolean
    <ArgenteaParameters("GC_ACTIVATION_PRINT_WITHIN_TA", True)>
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
    <ArgenteaParameters("GC_ACTIVATION_CHECK_COPIES", 1)>
    Public Property GiftCardActivationCheckCopies() As Integer
        Get
            Return _GiftCardActivationCheckCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardActivationCheckCopies = value
        End Set
    End Property

    Private _GiftCardActivationCheckSave As Boolean
    <ArgenteaParameters("GC_ACTIVATION_CHECK_SAVE", False)>
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
    <ArgenteaParameters("GC_REDEEM_COPIES", 1)>
    Public Property GiftCardRedeemCopies() As Integer
        Get
            Return _GiftCardRedeemCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardRedeemCopies = value
        End Set
    End Property

    Private _GiftCardRedeemSave As Boolean
    <ArgenteaParameters("GC_REDEEM_SAVE", False)>
    Public Property GiftCardRedeemSave() As Boolean
        Get
            Return _GiftCardRedeemSave
        End Get
        Set(ByVal value As Boolean)
            _GiftCardRedeemSave = value
        End Set
    End Property

    Private _GiftCardRedeemCheckPrintWithinTa As Boolean
    <ArgenteaParameters("GC_REDEEM_CHECK_PRINT_WITHIN_TA", True)>
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
    <ArgenteaParameters("GC_REDEEM_CHECK_COPIES", 1)>
    Public Property GiftCardRedeemCheckCopies() As Integer
        Get
            Return _GiftCardRedeemCheckCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardRedeemCheckCopies = value
        End Set
    End Property

    Private _GiftCardRedeemCheckSave As Boolean
    <ArgenteaParameters("GC_REDEEM_CHECK_SAVE", False)>
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
    <ArgenteaParameters("GC_REDEEM_CANCEL_COPIES", 1)>
    Public Property GiftCardRedeemCancelCopies() As Integer
        Get
            Return _GiftCardRedeemCancelCopies
        End Get
        Set(ByVal value As Integer)
            _GiftCardRedeemCancelCopies = value
        End Set
    End Property

    Private _GiftCardRedeemCancelSave As Boolean
    <ArgenteaParameters("GC_REDEEM_CANCEL_SAVE", False)>
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
    <ArgenteaParameters("EXT_GC_ACTIVATION_COPIES", 1)>
    Public Property ExtGiftCardActivationCopies() As Integer
        Get
            Return _ExtGiftCardActivationCopies
        End Get
        Set(ByVal value As Integer)
            _ExtGiftCardActivationCopies = value
        End Set
    End Property

    Private _ExtGiftCardActivationSave As Boolean
    <ArgenteaParameters("EXT_GC_ACTIVATION_SAVE", False)>
    Public Property ExtGiftCardActivationSave() As Boolean
        Get
            Return _ExtGiftCardActivationSave
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardActivationSave = value
        End Set
    End Property

    Private _ExtGiftCardActivationPrintWithinTa As Boolean
    <ArgenteaParameters("EXT_GC_ACTIVATION_PRINT_WITHIN_TA", True)>
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
    <ArgenteaParameters("EXT_GC_DEACTIVATION_CANCEL_COPIES", 1)>
    Public Property ExtGiftCardDeActivationCancelCopies() As Integer
        Get
            Return _ExtGiftCardDeActivationCancelCopies
        End Get
        Set(ByVal value As Integer)
            _ExtGiftCardDeActivationCancelCopies = value
        End Set
    End Property

    Private _ExtGiftCardDeActivationCancelSave As Boolean
    <ArgenteaParameters("EXT_GC_DEACTIVATION_CANCEL_SAVE", False)>
    Public Property ExtGiftCardDeActivationCancelSave() As Boolean
        Get
            Return _ExtGiftCardDeActivationCancelSave
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardDeActivationCancelSave = value
        End Set
    End Property

    Private _ExtGiftCardDeActivationPrintWithinTa As Boolean
    <ArgenteaParameters("EXT_GC_DEACTIVATION_PRINT_WITHIN_TA", True)>
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
    <ArgenteaParameters("EXT_GC_CONFIRM_COPIES", 1)>
    Public Property ExtGiftCardConfirmCopies() As Integer
        Get
            Return _ExtGiftCardConfirmCopies
        End Get
        Set(ByVal value As Integer)
            _ExtGiftCardConfirmCopies = value
        End Set
    End Property

    Private _ExtGiftCardConfirmSave As Boolean
    <ArgenteaParameters("EXT_GC_CONFIRM_SAVE", False)>
    Public Property ExtGiftCardConfirmSave() As Boolean
        Get
            Return _ExtGiftCardConfirmSave
        End Get
        Set(ByVal value As Boolean)
            _ExtGiftCardConfirmSave = value
        End Set
    End Property

    Private _ExtGiftCardConfirmPrintWithinTa As Boolean
    <ArgenteaParameters("EXT_GC_CONFIRM_PRINT_WITHIN_TA", True)>
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
    <ArgenteaParameters("PC_RECHARGE_CHECK_COPIES", 0)>
    Public Property PhoneRechargeCheckCopies() As Integer
        Get
            Return _PhoneRechargeCheckCopies
        End Get
        Set(ByVal value As Integer)
            _PhoneRechargeCheckCopies = value
        End Set
    End Property

    Private _PhoneRechargeCheckSave As Boolean
    <ArgenteaParameters("PC_RECHARGE_CHECK_SAVE", False)>
    Public Property PhoneRechargeCheckSave() As Boolean
        Get
            Return _PhoneRechargeCheckSave
        End Get
        Set(ByVal value As Boolean)
            _PhoneRechargeCheckSave = value
        End Set
    End Property

    Private _PhoneRechargeCheckPrintWithinTa As Boolean
    <ArgenteaParameters("PC_CHECK_PRINT_WITHIN_TA", False)>
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
    <ArgenteaParameters("PC_ACTIVATION_COPIES", 1)>
    Public Property PhoneRechargeActivationCopies() As Integer
        Get
            Return _PhoneRechargeActivationCopies
        End Get
        Set(ByVal value As Integer)
            _PhoneRechargeActivationCopies = value
        End Set
    End Property

    Private _PhoneRechargeActivationSave As Boolean
    <ArgenteaParameters("PC_ACTIVATION_CHECK_SAVE", False)>
    Public Property PhoneRechargeActivationSave() As Boolean
        Get
            Return _PhoneRechargeActivationSave
        End Get
        Set(ByVal value As Boolean)
            _PhoneRechargeActivationSave = value
        End Set
    End Property

    Private _PhoneRechargeActivationPrintWithinTa As Boolean
    <ArgenteaParameters("PC_ACTIVATION_PRINT_WITHIN_TA", True)>
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
    ''' <summary>
    '''     Codice concesso da Argntea al
    '''     cliente che fa uso del Modulo.
    ''' </summary>
    ''' <returns>String</returns>
    <ArgenteaParameters("BP_RUPP_ARGENTEA", "")>
    Public Property BPRupp() As String
        Get
            Return _BPRupp
        End Get
        Set(ByVal value As String)
            _BPRupp = value
        End Set
    End Property


    Private _BP_AcceptExcedeedValues As Boolean
    ''' <summary>
    '''	Se il Totale in ingresso è minore rispetto 
    '''	al valore di facciata del Buono Pasto una volta
    '''	ottenuto dalla materializzazione, si tronca il totale	
    '''	riportando la differenza nel metadata del record, oppure	
    '''	si nega l'operzione alla cassa.
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BP_ACCEPT_EXCEDEED_VALUES", False)>
    Public Property BP_AcceptExcedeedValues() As Boolean
        Get
            Return _BP_AcceptExcedeedValues
        End Get
        Set(ByVal value As Boolean)
            _BP_AcceptExcedeedValues = value
        End Set
    End Property

    Private _BP_AccorpateOnTA As Boolean
    ''' <summary>
    '''     Se si vuole accorpare i BP per ogni
    '''     sessione in un record dello scontrino
    '''     oppure per ogni demat effettuato si ha 
    '''     una riga sullo scontrino
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BP_ACCORPATE_ON_TA", False)>
    Public Property BP_AccorpateOnTA() As Boolean
        Get
            Return _BP_AccorpateOnTA
        End Get
        Set(ByVal value As Boolean)
            _BP_AccorpateOnTA = value
        End Set
    End Property

    Private _BP_MaxBPPayableSomeSession As Integer
    ''' <summary>
    '''     stabilisce il numero massimo di BP
    '''     utilizzabili per essere usati come
    '''     pagamento per ogni scontrino.
    '''     Con 0 il numero massimo non ha un limite.
    ''' </summary>
    ''' <returns>Un numero massimo di elementi per scontrino utilizzabili o 0 per nessun limite</returns>
    <ArgenteaParameters("BP_MAX_BP_PAYABLES_SOME_SESSION", 0)>
    Public Property BP_MaxBPPayableSomeSession() As Integer
        Get
            Return _BP_MaxBPPayableSomeSession
        End Get
        Set(ByVal value As Integer)
            _BP_MaxBPPayableSomeSession = value
        End Set
    End Property

    Private _ViewModeElementsDeleteted As Byte
    ''' <summary>
    '''     Con questa proprietà definiamo alla visualizzazione
    '''     dell'emulatore software se viene visualizzato il 
    '''     checkbox per la scelta utente di vedere o meno gli
    '''     elementi elimanti con 2 definito come impostazione di
    '''     default a vederli con il check impostato a true o 3
    '''     con il check impostato a false di non vederli, mentre con
    '''     0 di Vederli sempre e senza dare scelta all'utente e con 
    '''     1 a non vederli mai e senza scelta ultente.
    ''' </summary>
    ''' <returns>0 = Sempre Visibili - 1 = Sempre non visibili - 2 = Visibili o meno a scelta dell'utente con default a Visibili 3 = Visibili o meno a scelta dell'utente con default a Non Visibili</returns>
    <ArgenteaParameters("BP_VIEW_MODE_ELEMENTS_DELETETED", 0)>
    Public Property BP_ViewModeElementsDeleteted() As Byte
        Get
            Return _ViewModeElementsDeleteted
        End Get
        Set(ByVal value As Byte)
            _ViewModeElementsDeleteted = value
        End Set
    End Property

    Private _ViewModeTotalsAndPartials As Byte
    ''' <summary>
    '''     Con questa proprietà definiamo alla visualizzazione
    '''     dell'emulatore software se viene visualizzato per
    '''     le caselle di totali e parziali scalabili anche
    '''     i valori rispetto alla sessione o al solo raggruppamento.
    '''     per default a vederli 0 riporterà solo il totale di sessione
    ''' </summary>
    ''' <remarks>
    '''     0 -> Visualizziamo 2 caselle con l'importo totale comprensivo del totale delle sessioni precedenti
    '''     1 -> Visualizziamo 2 caselle con l'importo totale comprensivo del totale delle sessioni precedenti ( e nella lista gli importi paganti delle sessioni precedenti )
    '''     2 -> Visualizziamo 2 caselle con l'importo totale comprensivo del totale delle sessioni precedenti ( e nella lista gli importi paganti e non delle sessioni precedenti )
    '''     3 -> Visualizziamo 2 caselle escluso l'importo totale delle sessioni precedenti riportante solo il parziale della sessione corrente ( l'elenco non mostra niente solo quelli nuovi  della sesione corrente )
    '''     4 -> Visualizziamo 4 caselle con l'importo totale delle sessioni precedenti e quello parziale della sessione corrente ( l'elenco non mostra niente solo quelli nuovi  della sesione corrente )
    '''     5 -> Visualizziamo 4 caselle con l'importo totale delle sessioni precedenti e quello parziale della sessione corrente ( e nella lista gli importi paganti delle sessioni precedenti )
    '''     6 -> Visualizziamo 4 caselle con l'importo totale delle sessioni precedenti e quello parziale della sessione corrente ( e nella lista gli importi paganti e non delle sessioni precedenti )
    ''' </remarks>
    ''' <returns>Valore da 0 a 6 a secondo dell'impostazione definita</returns>
    <ArgenteaParameters("BP_VIEW_MODE_TOTALS_AND_PARTIALS", 0)>
    Public Property BP_ViewModeTotalsAndPartials() As Byte
        Get
            Return _ViewModeTotalsAndPartials
        End Get
        Set(ByVal value As Byte)
            _ViewModeTotalsAndPartials = value
        End Set
    End Property

    Private _BP_ShowWaitScreenLevel As Byte
    ''' <summary>
    '''     In base al livello mostra o meno i Waitscreen
    '''     nella chiamata al servizio di emulazione POS.
    '''     Livello 0 Non mostra in nessun caso i WaitScreen
    '''     Livello 1 Mostra in tutti i casi i WaitScreen
    '''     Livello 2 Mostra i WaitScreen solo necessari a apecifiche azioni utente
    ''' </summary>
    ''' <returns>Byte da 0 a 2 1 a default</returns>
    <ArgenteaParameters("BP_SHOW_WAITSCREEN_LEVEL", 1)>
    Public Property BP_ShowWaitScreenLevel() As Byte
        Get
            Return _BP_ShowWaitScreenLevel
        End Get
        Set(ByVal value As Byte)
            _BP_ShowWaitScreenLevel = value
        End Set
    End Property

    Private _BP_ShowMessageBoxLevel As Byte
    ''' <summary>
    '''     Mostra i messaggi di stato a condizione
    '''     che il livello di mostrare i messaggi sia
    '''     definito secondo questa condizione.:
    '''     Livello 0 Non mostra in nessun caso i Messagi di Errore Tranne quelli con Eccezioni Interne o di Parsing
    '''     Livello 1 Mostra in tutti i casi i Messaggi di Errore di azioni utente e il resto
    '''     Livello 2 Mostra i MessageBox solo necessari a apecifiche azioni utente e le eccezioni Interne o di Parsing
    ''' </summary>
    ''' <returns></returns>
    <ArgenteaParameters("BP_SHOW_MESSAGEBOX_LEVEL", 1)>
    Public Property BP_ShowMessageBoxLevel() As Byte
        Get
            Return _BP_ShowMessageBoxLevel
        End Get
        Set(ByVal value As Byte)
            _BP_ShowMessageBoxLevel = value
        End Set
    End Property

    Private _BP_ViewFormForMessageStatusPos As Boolean
    ''' <summary>
    '''     Se si vogliono far vedere i messaggi di stato
    '''     in basso al form durante le operazioni di 
    '''     dematerializzazione che arrivano dal POS Hardware
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BP_VIEW_FORM_FOR_MESSAGES_STATUS_POS", False)>
    Public Property BP_ViewFormForMessagesStatusPos() As Boolean
        Get
            Return _BP_ViewFormForMessageStatusPos
        End Get
        Set(ByVal value As Boolean)
            _BP_ViewFormForMessageStatusPos = value
        End Set
    End Property

#Region "BPE"

    Private _BPE_AutoCloseOnCompleteOperation As Boolean
    ''' <summary>
    '''     Nelle operazioni di  interrogazione
    '''     e stato con il POS Hardware su True
    '''     fa in modo che il form di comunicazione
    '''     venga chiuso immediatamente senza 
    '''     intervento operatore una volta completata.
    '''     Con false appare il tasto OK sul form di
    '''     comunicazione per far chiudere all'operatore.
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BPE_AUTOCLOSE_ON_COMPLETE_OPERATION", True)>
    Public Property BPE_AutoCloseOnCompleteOperation() As Boolean
        Get
            Return _BPE_AutoCloseOnCompleteOperation
        End Get
        Set(ByVal value As Boolean)
            _BPE_AutoCloseOnCompleteOperation = value
        End Set
    End Property

    Private _BPE_ReattemptToCompleteOperation As Boolean
    ''' <summary>
    '''     Nelle operazioni di interrogazione
    '''     e stato con il POS Hardware su True
    '''     fa in modo che in caso di errore di
    '''     comunicazione con il dispositivo o
    '''     per annulli utente sul dispositivo 
    '''     si voglia che appaia il tasto di ritentare
    '''     piuttosto che uscire dal form sempre
    '''     che l'autoclose non sia impostato a True
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BPE_REATTEMPT_TO_COMPLETE_OPERATION", False)>
    Public Property BPE_ReattemptToCompleteOperation() As Boolean
        Get
            Return _BPE_ReattemptToCompleteOperation
        End Get
        Set(ByVal value As Boolean)
            _BPE_ReattemptToCompleteOperation = value
        End Set
    End Property


    Private _BPE_Copies As Integer
    ''' <summary>
    '''     Definisce il numero di copie della
    '''     transazione Argentea da stampare una
    '''     volta chiuso lo scontrino sempre che
    '''     la stampa sia abilitata in chiusura.
    ''' </summary>
    ''' <returns>Integer</returns>
    <ArgenteaParameters("BPE_COPIES", 1)>
    Public Property BPECopies() As Integer
        Get
            Return _BPE_Copies
        End Get
        Set(ByVal value As Integer)
            _BPE_Copies = value
        End Set
    End Property

    Private _BPE_PrintWithinTa As Boolean
    ''' <summary>
    '''     Abilita la stampa a chiusura dello
    '''     scontrino una volta conclusa l'operazione
    '''     con la transazione del POS Hardware
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BPE_PRINT_WITHIN_TA", True)>
    Public Property BPEPrintWithinTa() As Boolean
        Get
            Return _BPE_PrintWithinTa
        End Get
        Set(ByVal value As Boolean)
            _BPE_PrintWithinTa = value
        End Set
    End Property

    Private _BPE_ViewMessagesOnReturnFromPos As Boolean
    ''' <summary>
    '''     Nel Form Hardware se si vuole far vedere
    '''     le messagebox di errore pos che arrivano
    '''     all'operatore.
    ''' </summary>
    ''' <returns>Boolean</returns>
    <ArgenteaParameters("BPE_VIEW_MESSAGES_ON_RETURN_FROM_POS", False)>
    Public Property BP_ViewMessagesOnReturnFromPos() As Boolean
        Get
            Return _BPE_ViewMessagesOnReturnFromPos
        End Get
        Set(ByVal value As Boolean)
            _BPE_ViewMessagesOnReturnFromPos = value
        End Set
    End Property

#End Region

#End Region

#Region "Common"
    Private _PrintLogoOnExternalReceipts As Boolean
    <ArgenteaParameters("PRINT_LOGO_ON_EXTERNAL_RECEIPTS", True)>
    Public Property PrintLogoOnExternalReceipts() As Boolean
        Get
            Return _PrintLogoOnExternalReceipts
        End Get
        Set(ByVal value As Boolean)
            _PrintLogoOnExternalReceipts = value
        End Set
    End Property

#End Region

    Public Function LoadParametersByReflection(ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        LoadParametersByReflection = False
        Dim funcName As String = "LoadParametersByReflection"
        Dim szTmp As String = String.Empty
        Dim attr As ArgenteaParametersAttribute

        Try

            For Each pi As System.Reflection.PropertyInfo In Me.GetType.GetProperties()

                Try
                    attr = pi.GetCustomAttributes(GetType(ArgenteaParametersAttribute), False)(0)
                Catch ex As Exception
                    Continue For
                End Try

                szTmp = TheModCntr.getParam(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + attr.Name)
                If String.IsNullOrEmpty(szTmp) OrElse String.IsNullOrWhiteSpace(szTmp) Then
                    ' use default
                    szTmp = attr.DefaultValue
                End If

                If True Then
                    System.Diagnostics.Debug.WriteLine(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + attr.Name + " = " + szTmp)
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

                    ElseIf pi.PropertyType = GetType(Byte) Then

                        pi.SetValue(Me, Byte.Parse(szTmp), Nothing)

                    Else

                        pi.SetValue(Me, szTmp, Nothing)

                    End If

                Catch ex As Exception

                    System.Diagnostics.Debug.WriteLine(TPDotnet.Pos.PARAMETER_DLL_NAME + "." + "Argentea" + "." + attr.Name + " EXCEPTED ")

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
