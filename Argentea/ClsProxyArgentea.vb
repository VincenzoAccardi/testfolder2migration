Imports System
Imports ARGLIB = PAGAMENTOLib
Imports System.Collections
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos.EFT
Imports TPDotnet.Pos
Imports System.Windows.Forms

#Const DEBUG_SERVICE = 1

''' <summary>
'''     Oggetto Proxy di Comunicazione e  gestione
'''     dei dati che arrivano dal POS Esterno dato
'''     da Argentea secondo la dll esterna COM Monetica.
''' </summary>
Public Class ClsProxyArgentea

    ''' <summary>
    '''     Argomenti specifici per la comunicazione con Argentea
    ''' </summary>
    Protected ArgenteaCOMObject As ARGLIB.argpay

    ' COSTANTI PARAMETRI UTILIZZATE in Operator su IT.Parameter
    Private Const OPT_BPNumMaxPayablesOnVoid As String = "BP_NumMaxPayablesOnVoid"          ' <-- Numero massimo di Buoni Pasto utilizzati per la vendita in corso 0 o ^n

    ''' <summary>
    '''     Collection usata per  le Transazioni
    '''     Argentea andati a buon  fine dopo la
    '''     chiamata alla funzione Argentea o da 
    '''     POS hardware
    ''' </summary>
    'Private _listBpCompletated As New Collections.Generic.Dictionary(Of String, BPType)(System.StringComparer.InvariantCultureIgnoreCase)
    Private _DataResponse As DataResponse

#Region "CONST di INFO e ERRORE Private"

    ' Messaggeria per codifica segnalazioni ID di errore remoti 
    Private msgUtil As New TPDotnet.IT.Common.Pos.Common

    ' Su Errore di double Connect()
    Private Const GLB_PROXY_ALREDAY_RUNNING As String = "ERROR-PROXY-ALREADY-RUN"

    ' Su Errore Parsing utilizzata per segnalazione
    ' di errore su protocollo non previsto.
    '   Errore classificato su risposta Argentea .: RetVal.CodeResult 
    Private Const GLB_ERROR_PARSING As String = "ERROR-PARSING"

    ' Su Errore Pagabile rispetto all'ammontare del pagamento
    ' assegno errore di non valido per uscire dal pagamento BPC
    Private Const GLB_ERROR_PAYABLE As String = "ERROR_PAYABLE"

    ' Su errore quando l'opzione di girare il resto in eccesso
    ' assegno la costante di operazione non valida ai fine di pagamento.
    Private Const GLB_OPT_ERROR_VALUE_EXCEDEED As String = "ERROR-OPTION-PAYABLE-WITH-REST"

    ' BarCode già utilizzato in precedenza evitiamo
    ' di richiamare argentea per il controllo
    Private Const GLB_OPT_ERROR_NUMEBP_EXCEDEED As String = "ERROR-OPTION-PAYABLE-NUMBP-EXCEDEED"

    ' Nell'istanziare e castare il Form di appoggio
    ' a scansione dei Barcode solleva questa eccezione.
    Private Const GLB_ERROR_FORM_INSTANCE As String = "Error-INSTANCE-FORM"

    ' Nell'eseguire l'evento al chiamante per restituire i dati
    ' c'è stato un errore o un eccezione non gestita nell'evento.
    ' intercettiamo e proponiamo di controllare l'aggiornamento sul consumer.
    Private Const GLB_ERROR_ON_EVENT_DATA As String = "Error-Eventa-Data"

    ' Nell'aggiungere elementi alla griglia o per 
    ' motivi legati alla gestione del form non
    ' previsti solleva questa eccezione interna.
    Private Const GLB_ERROR_FORM_DATA As String = "Error-FORM-FLOWDATA"

    ' BarCode già utilizzato in precedenza evitiamo
    ' di richiamare argentea per il controllo
    Private Const GLB_INFO_CODE_ALREADYINUSE As String = "Error-BARCODE-ALREADYINUSE"

    ' Quando l'importo delle righe è già stato raggiunto (con o 
    ' senza eccesso per eventuale ipotesi di resto) non procediamo.
    Private Const GLB_INFO_IMPORT_ALREADYCOMPLETED As String = "Error-IMPORT-ALREADY_COMPLETED"

    ' BarCode da rimuovere da quelli già scanditi in precedenza 
    ' non presente in elenco
    Private Const GLB_INFO_CODE_NOTPRESENT As String = "Error-BARCODE-NOTPRESENT"

    ' Nel Flow della funzione Entry il Throw non
    ' previsto.
    Private Const GLB_ERROR_NOT_UNEXPECTED As String = "Error-EXCEPTION-UNEXPECTED"

    ' Se un operazione interna ha sollevato eccezione e non ha
    ' potuto completare l'evento per ripassarlo ai media.
    Private Const GLB_ERROR_NOT_UNEXPECTED2 As String = "Error-EXCEPTION-UNEXPECTED2"

    ' Eccezione nell'istanza di questa classe se  non  dovesse
    ' caricare per qualche motivo i parametri globali dedicati
    ' alla gestione di esecuzione del  servizio  corrente  con
    ' argentea.
    Private Const GLB_ERROR_INSTANCE_PARAMETERS As String = "Error-ONLOAD-PARAMETERS-ARGENTEA"

    ' Su errori non bloccanti ma da segnalare all'operatore
    ' come stmapa scontrino non effettuata o altro usiamo questo.
    Private Const GLB_SIGNAL_OPERATOR As String = "Error-SIGNALS_GENERIC"

    ' Su chiamate a Nomi di Funzioni API remote
    ' che non sono gestite da questo proxy
    Private Const GLB_ERROR_API_NOT_VALID As String = "Error-API_NOT_VALID"

    ' Se un tentativo di inizializzazione e uno di riallinemento
    ' e ritentativo di inzializzazione ha fallito abbiamo FALLITO punto e basta.
    Private Const GLB_FAILED_INITIALIZATION As String = "Error-FAILED_INIATIALIZATION"

    ''' <summary>
    '''     Stati interni per la Risposta 
    '''     tra funzioni nel triplo stato.
    ''' </summary>
    Private Enum StatusCode

        ''' <summary>
        '''     OK
        ''' </summary>
        OK

        ''' <summary>
        '''     KO
        ''' </summary>
        KO

        ''' <summary>
        '''     In risposta di OK
        '''     ma richiede conferma.
        ''' </summary>
        CONFIRMREQUEST

    End Enum

    '
    ' Elementi necessari per il parsing
    ' sul protocollo previsto e di frazione
    ' per il risultato.
    '
    Private m_ParseSplitMode As String = "-"
    Private m_ParseFractMode As Integer = 100
    Private m_ProtoFractMode As Integer = 1

#End Region

#Region "Enum o Costanti Pubbliche"

    ''' <summary>
    '''     Il tipo di istanza corrente.:
    '''     
    '''     -> Service:
    '''         Si occupa di comunicare con il servizio remoto di
    '''         Argentea e gestirne il protoccolo per le azioni 
    '''         passando sempre dalla dll comune di Argentea ed 
    '''         operando su un form di istanza legato alla cassa
    '''         con la gestione degli Handler per ogni Barcode di
    '''         BP passato.
    '''         
    '''     -> Pos:
    '''         Si occupa di comunicare con il dispostivo locale
    '''         Pos che risponde su un file txt e trasformato in
    '''         protocollo dalla dll com di argentea. Il comportamento
    '''         rispetto al servizio è che il totalizztore arriva
    '''         dopo la prima chiamata e accumula l'insieme dei PB
    '''         che sono stati scansionati dal dispositivo.
    ''' 
    ''' </summary>
    Public Enum enTypeProxy

        ''' <summary>
        '''     Servizio remoto Argentea con Form di gestione locale per i Barcode relativi ai BP Cartacei
        ''' </summary>
        Service

        ''' <summary>
        '''     Servizio locale Argentea ttramite dispositivo Pos collegato per le Tessere BP usate per pagare
        ''' </summary>
        Pos

    End Enum

    ''' <summary>
    '''     Per la risposta in uscita specifichiamo
    '''     nel dataResponse che tipo di BP sono 
    '''     stati trattati durante lo svolgimento 
    '''     del proxy
    ''' </summary>
    Public Enum enTypeBP

        ''' <summary>
        '''     Tipo di BP Cartacei con ognuno un suo Barcode
        '''     espletati nella modalità software in Service mode
        ''' </summary>
        TicketsRestaurant

        ''' <summary>
        '''     Tipo di BP Elettronici in sequinza tipo Barcode
        '''     espletati nella modalità hardware in POS mode
        ''' </summary>
        TicketsCard

    End Enum

    ''' <summary>
    '''     Stato del service corrente
    '''     in modalità POS software verso Argentea
    '''     o in modalità POS locale verso POS hardware collegato
    ''' </summary>
    Public Enum enProxyStatus

        ''' <summary>
        '''     Non inizializzato
        ''' </summary>
        Uninitializated

        ''' <summary>
        '''     Iniziailizzato
        ''' </summary>
        Initializated

        ''' <summary>
        '''     Su operazioni in corso è avviato
        ''' </summary>
        InRunning

        ''' <summary>
        '''     Su errore di qualche procedura (Controllare LastError)
        ''' </summary>
        InError

        ''' <summary>
        '''     Dovuto a qualche KO inteso nella gestione della transazione (Quindi da intendere non concluso nel servizio)
        ''' </summary>
        KO

        ''' <summary>
        '''     Status operativo OK
        ''' </summary>
        OK

    End Enum


    Public Enum enCommandToCall

        ''' <summary>
        '''     Avvia il servizio remoto
        '''     o il pos locale per un Pagamento
        ''' </summary>
        Payment

        ''' <summary>
        '''     Avvia il servizio remoto
        '''     o il pos locale per uno Storno
        ''' </summary>
        Void


    End Enum

    ''' <summary>
    '''     Usato nelle chiamate Hundler per
    '''     formattare secondo le specifiche
    '''     del protocollo Argentea il CSV in
    '''     risposta sul MsgOut sul servizo POS
    '''     remoto di Argentea
    ''' </summary>
    Enum TypeCodifiqueProtocol

        ''' <summary>
        '''     IL CSV su risposta dietro Iniziazlizzazione
        ''' </summary>
        Inizialization
        ''' <summary>
        '''     IL CSV su risposta dietro Dematerializzazione
        ''' </summary>
        Dematerialization
        ''' <summary>
        '''     IL CSV su risposta dietro Undo Dematerializzazione
        ''' </summary>
        Reverse
        ''' <summary>
        '''     IL CSV su risposta dietro Confirm Dematerializzazione
        ''' </summary>
        Confirm

    End Enum

#End Region

#Region "Membri Privati"

    '
    ' Comportamento a secondo del proxy di utilizzo
    ' e status logico interno rispetto all'inizio e fine di vita.
    '
    Private m_TypeProxy As enTypeProxy                                          ' <-- Il Comportamento che deve assumere il Procy corrente
    Private m_CommandToCall As enCommandToCall = enCommandToCall.Payment        ' <-- Il comando che deve eseguire nella modalità in esecuzione (per default in pagamento)
    Private m_CurrentTransactionID As String = String.Empty                     ' <-- La Transazione in corso GUID 
    Private m_CurrentPaymentsTotal As Decimal = 0                               ' <-- L'importo pagato da altri media in ingresso sul servizio
    Private m_ServiceStatus As enProxyStatus = enProxyStatus.Uninitializated    ' <-- Lo stato iniziale ed in corso del flow del Proxy corrente
    Private m_SilentMode As Boolean = False                                     ' <-- Se mostrare all'utente i messaggi di errore e di avviso

    '
    ' Per il servizio che usa un Form interno
    ' per l'inserimento manuale dei Buoni Pasto
    ' mi appoggio su un Form della cassa già presente.
    '
    Private frmEmulation As Form = Nothing                      ' <-- Il Form di appoggio per servire il POS software sulla cassa corrente

    '
    ' Passati dal Chiamante per essere
    ' letti agiornati in uscita.
    '
    Private m_PaidAmount As Decimal                             ' <-- Il Pagato fino ad adesso all'entrata
    Private m_PayableAmount As Decimal                          ' <-- Il pagabile con le azioni del servizio
    Private m_VoidAmount As Decimal                             ' <-- Lo Storno attuale fino ad adesso all'entrata
    Private m_VoidableAmount As Decimal                         ' <-- Lo Stornabile o lo stornato con le azioni del servizio
    Private m_PrefillVoidable As Dictionary(Of String, PaidEntry) ' <-- Gestisce un possibile elenco di BP prefillato sul FORM di appoggio per gestire storni tramite operatore

    '
    '   Per le chiamate API dirette
    '
    Private m_CurrentApiNameToCall As String                    ' <-- Riporta l'ultima chiamata API tramite dll COM di argentea chiamata

    '
    ' Aggiornati per il Risultato
    '
    Shared m_TypeBPElaborated_CS As enTypeBP                ' <-- Il Tipo di BP o gruppo di BP elaborati nella sessione
    Shared m_TotalBPUsed_CS As Integer                      ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento o strorno
    '
    Shared m_TotalPayed_CS As Decimal                       ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente per il pagamento
    Shared m_TotalValueExcedeed_CS As Decimal               ' <-- Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata
    '
    Shared m_TotalVoided_CS As Decimal                      ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente per lo storno
    Shared m_TotalVoidedExcedeed_CS As Decimal              ' <-- Il Totale in eccesso/difetto se l'opzione per accettare valori maggiori è abilitata in storno

    '
    ' Variabili private
    '
    Private m_LastStatus As String                                      ' <-- Ultimo Status di Costante per errore in STDOUT
    Private m_LastErrorMessage As String                                ' <-- Ultimo Messaggio di errore STDOUT
    Private m_LastResponseRawArgentea As ArgenteaFunctionReturnObject   ' <-- Ultima risposta di Argentea per STDOUT (di utilità al reprint dello scontrino)

    '
    ' Status interni e ultime letture
    '
    Private m_FirstCall As Boolean = False                  ' <-- Inizializzazione alla prima chiamata dal Form di scansione per i Barcode vs Argentea
    Private m_CurrentBarcodeScan As String = String.Empty   ' <-- Ultimo Barcode scansionato
    Private m_CurrentValueOfBP As Decimal                   ' <-- Valore facciale dell'n Barcode di BP scansionato
    Private m_CurrentTerminalID As String = Nothing         ' <-- In Pos Hardware identifica il POS usato in Software l'ID del WebService

    '
    ' Status operativi proxy mode
    '
    Private m_bWaitActive As Boolean = False                ' <-- Stato iniziale in attesa
    Private m_bPosActive As Boolean = False                 ' <-- Hardware POS in attesa
    Private m_FlagUndoBPCForExcedeed As Boolean = False     ' <-- Flag interno per richiamare funzione di storno in event handler da altra funzione

    ' 
    ' Interni per gestione
    '
    Protected m_TheModcntr As ModCntr                       ' <-- Il controller principale applicativo che non deve mancare mai
    Protected m_taobj As TA                                 ' <-- la TA in corso da cui ricavare informazioni

    '
    ' Parametri Globali di applicazione predefiniti
    ' nel contesto backStore per Argentea.
    '
    Private st_Parameters_Argentea As ArgenteaParameters    ' <-- Riprende dal modello statico tutti i parametri globali nel contesto corrente dedicati ad Argentea

#End Region

#Region "Properties di Classe"

    ''' <summary>
    '''     L'insieme condivisio con l'applicativo corrente per
    '''     le Opzioni e Parametri di comportamento con il
    '''     servizio corrente di proxy verso Argentea.
    '''     Questi parametri e queste opzioni sono definibili
    '''     dentro il BackStore applicativo, determinate secondo
    '''     l'uso e consumo del plugin corrente.
    ''' </summary>
    ''' <returns>I Parametri applicativi per il plugin corrente <see cref="ArgenteaParameters"/></returns>
    Friend ReadOnly Property ArgenteaParameters() As ArgenteaParameters
        Get
            Return st_Parameters_Argentea
        End Get
    End Property

    ''' <summary>
    '''     Restituisce o imposta l'azione da eseguire nella modalità prevista
    '''     sul proxy corrente verso il servizio remoto o il pos terminal locale.
    ''' </summary>
    ''' <returns>String</returns>
    Friend Property Command() As enCommandToCall
        Get
            Return m_CommandToCall
        End Get
        Set(value As enCommandToCall)
            m_CommandToCall = value
        End Set
    End Property

    ''' <summary>
    '''     Stato del Pos al  momento
    '''     della prima comunicazione
    '''     e in quelle succesive.
    ''' </summary>
    ''' <returns>Uno dei possibili stati del proxy corrente in corso o come è uscito<see cref="enProxyStatus"/></returns>
    ''' <remarks>Se si effettua il close ricomincia da non inizializzato</remarks>
    Public ReadOnly Property ProxyStatus() As enProxyStatus
        Get
            Return m_ServiceStatus
        End Get
    End Property

    ''' <summary>
    '''     E' avviato ed in attesa
    '''     di operare in service o terminale.
    ''' </summary>
    ''' <returns>True se è già stato avviato</returns>
    Public ReadOnly Property IsLive() As Boolean
        Get
            Return m_bWaitActive
        End Get
    End Property

    ''' <summary>
    '''     Se visualizare o meno messagi di avviso
    '''     o di errore all'operatore tramite msgbox.
    ''' </summary>
    ''' <returns>True/False</returns>
    Public Property SilentMode() As Boolean
        Get
            Return m_SilentMode
        End Get
        Set(Value As Boolean)
            m_SilentMode = Value
        End Set
    End Property

#End Region

#Region "Properties Controller e Funzioni per Service Mode"

    '
    '   Variabili private per la gestione interna
    '   in emulazione software del servizio Argentea
    '
    Private m_ProgressiveCall As Integer = 1            '   <--     Il progressive call Privato (Relativo a tutte le chiamate in sequenza richieste dal protocollo)"
    Private m_RUPP As String = Nothing                  '   <--     Il RUPP necessario per comunicare con il servizio Remoto

    ''' <summary>
    '''     Restituisce per il numero delle chiamate verso il servizio remoto Argentea
    '''     ogni qualvolta si interagisce con lo stesso per eseguire le sue API.
    ''' </summary>
    ''' <returns>Il numero delle chiamate durante l'interrrogazioni API al servizio service remoto Argentea</returns>
    Private ReadOnly Property ProgressiveCalled() As Integer
        Get
            Return m_ProgressiveCall
        End Get
    End Property

    ''' <summary>
    '''     Internamente incrementa di uno il numero delle chiamate al servizio remoto per <see cref="ProgressiveCalled"/>
    ''' </summary>
    Private Function _IncrementProgressiveCall(Optional GetLast As Boolean = False) As Integer
        Static LastProgressive As Integer = 0

        If GetLast Then
            Return LastProgressive
        Else
            m_ProgressiveCall += 1
            LastProgressive = m_ProgressiveCall
            Return LastProgressive
        End If

    End Function

    ''' <summary>
    '''     Restituisce anche nelle nuove istanze quale
    '''     è l'ultimo numero che è stato  incrementato
    '''     sul sistema remoto Argentea.
    ''' </summary>
    ''' <returns>Il numeratore delle chiamate verso Argentea</returns>
    Private Function _GetLastProgressive() As Integer
        ' tech. Staticamente l'ultimo numero di contatore utilizzato
        Return _IncrementProgressiveCall(True)
    End Function

    ''' <summary>
    '''     Restituisce per la codifica il riferimento al Codice RUPP ripreso dalla Configurazione Globale presa
    '''     nel Backsotre di Cassa dal Client che identitica il POS Hardware e Service ID di Account verso Argentea.
    ''' </summary>
    ''' <returns>String</returns>
    Private ReadOnly Property GetPar_RUPP() As String
        Get
            ' tech. primo accesso ( * per i parametri uso questo trick )
            If m_RUPP <> Nothing And m_RUPP <> "" Then m_RUPP = st_Parameters_Argentea.BPRupp
            Return m_RUPP
        End Get
    End Property

    ''' <summary>
    '''     Restituisce per la codifica il riferimento al codice della cassa che sta effettuanto l'operazione
    ''' </summary>
    ''' <returns>String</returns>
    Private ReadOnly Property Get_CodeCashDevice() As String
        Get
            Return m_taobj.lRetailStoreID.ToString().PadLeft(5, "0") + m_taobj.lWorkStationNmbr.ToString().PadLeft(5, "0")
        End Get
    End Property

    ''' <summary>
    '''     Restituisce per la codifica il riferimento al codice dell'operatore che sta facendo l'operazione
    ''' </summary>
    ''' <returns>String</returns>
    Private ReadOnly Property Get_CodeOperatorID() As String
        Get
            Return m_taobj.lActOperatorID.ToString()
        End Get
    End Property

    ''' <summary>
    '''     Restituisce per la codifica il riferimento al numero di scontrino in corso sulla sessione
    '''     del controller princiaple di cassa.
    ''' </summary>
    ''' <returns>String</returns>
    Private ReadOnly Property Get_ReceiptNumber() As String
        Get
            Return m_taobj.lactTaNmbr
        End Get
    End Property

#End Region

#Region "Events Started and Finisched comunication POS"

    ''' <summary>
    '''     Evento al momento della chiusura di una intera Operazione
    '''     di transazione utente sul POS remoto con il collect dei
    '''     dati relativi ai buoni pasto utilizzati per pagare in 
    '''     formato elettronico.
    ''' </summary>
    ''' <param name="sender">Il POS Software o il POS hardware</param>
    ''' <param name="resultData">Il Collect dei dati arrivati dal POS per ricavare il totale del pagamento e le altre informazioni</param>
    Public Event Event_ProxyCollectDataTotalsAtEnd(ByRef sender As Object, ByRef resultData As DataResponse)

    ''' <summary>
    '''     Evento al momento della chiusura di una intera Operazione
    '''     di transazione utente sul POS remoto con il collect dei
    '''     dati relativi ai buoni pasto stornati da un pagamento 
    '''     effetuato in precedenzaformato elettronico.
    ''' </summary>
    ''' <param name="sender">Il POS Software o il POS hardware</param>
    ''' <param name="resultData">Il Collect dei dati arrivati dal POS per ricavare il totale dello storno e le altre informazioni</param>
    Public Event Event_ProxyCollectDataVoidedAtEnd(ByRef sender As Object, ByRef resultData As DataResponse)

#End Region

#Region ".ctor"

    ''' <summary>
    '''     .ctor
    ''' </summary>
    ''' <param name="theModCntr">controller -> Il Controller per riferimento dal chiamante</param>
    ''' <param name="taobj">transaction -> La TA per riferimento dal chiamante</param>
    ''' <param name="TypeBehavior">Definisce il comportamente di questa istanza proxy su servizio <see cref="enTypeProxy"/></param>
    ''' <param name="CurrentTransactionID">L'ID della transazione in corso.</param>
    ''' <param name="CurrentPaymentsTotal">Il Totale pagato fino adesso prima di effettuare l'eggiornamento dai dati del proxy corrente.</param>
    Protected Friend Sub New(
                             ByRef theModCntr As ModCntr,
                             ByRef taobj As TA,
                             TypeBehavior As enTypeProxy,
                             ByVal CurrentTransactionID As String,
                             ByVal CurrentPaymentsTotal As Decimal
                             )

        ' Tipo BEHAVIOR
        m_TypeProxy = TypeBehavior

        ' Dati fondamentali
        m_CurrentTransactionID = CurrentTransactionID
        m_CurrentPaymentsTotal = CurrentPaymentsTotal

        '
        ' Caricamento dei Parametri Argentea
        ' globali ripresi dal Backstore come
        ' utilizzati nel contesto corrente.
        '
        Try

            '
            ' Legge per impostazione statica 
            ' tutti i parametri  applicativi
            ' che influenzano il comportamento.
            '
            st_Parameters_Argentea = New ArgenteaParameters()
            'st_Parameters_Argentea.LoadParameters(theModCntr)
            st_Parameters_Argentea.LoadParametersByReflection(theModCntr)

        Catch ex As Exception


            ' Signal come errore di stampa ma non bloccante
            m_LastStatus = GLB_ERROR_INSTANCE_PARAMETERS
            m_LastErrorMessage = "Non è stato possibile caricare i parametri applicativi per eseguire il servizio Argentea"

            ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPSTOP)

            ' Bloccante
            Throw New Exception(m_LastStatus, ex)

        End Try

        '
        ' Gli oggetti di base
        '
        m_taobj = taobj
        m_TheModcntr = theModCntr

    End Sub

    ''' <summary>
    '''     In presenza di situazioni di Storno
    '''     può essere presente un  Prefill  di
    '''     tutti i BP facenti parte della TA o
    '''     di un elemento raggruppato come Media.
    ''' </summary>
    Private Sub PrefillVoidableOnPosSoftware(ByRef formTD As FormBuonoChiaro)

        If Not m_PrefillVoidable Is Nothing Then
            Dim ItemNew As PaidEntry

            For Each itm As KeyValuePair(Of String, PaidEntry) In m_PrefillVoidable

                '
                ' Questo è un Elemento che va visualizzato
                ' sul form e immesso nel dataResult per 
                ' una corretta gestione in uscita.
                '
                ItemNew = itm.Value
                m_TotalPayed_CS += ItemNew.Value
                m_TotalBPUsed_CS += 1

                'Aggiungo al dataResult per il calcolo in
                ' uscita da usare per aggiornare la TA
                WriterResultDataList.Add(ItemNew)

                ' Se è un elemento solo di riporto da stornato
                ' in sessioni precedenti
                If ItemNew.Emitter <> "[[VOIDED]]" Then

                    ' Aggiungo l'elemento al controllo Griglia
                    formTD.PaidEntryBindingSource.Add(ItemNew)

                End If

            Next

        End If

    End Sub




#End Region

#Region "Properties Pubbliche"

    ''' <summary>
    '''     All'entrata definisce il Pagato al momento
    '''     All'usscita è aggiornato con il pagato dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property AmountPaid() As Decimal
        Get
            Return m_PaidAmount
        End Get
        Set(ByVal value As Decimal)
            m_PaidAmount = value
            _updatePosForm()
        End Set
    End Property

    ''' <summary>
    '''     All'entrata definisce il Pagabile massimo
    '''     All'uscita è aggiornato con il pagato dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property AmountPayable() As Decimal
        Get
            Return m_PayableAmount
        End Get
        Set(ByVal value As Decimal)
            m_PayableAmount = value
            _updatePosForm()
        End Set
    End Property

    ''' <summary>
    '''     All'entrata definisce lo Storno al momento
    '''     All'uscita è aggiornato con lo Storno dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property AmountVoid() As Decimal
        Get
            Return m_VoidAmount
        End Get
        Set(ByVal value As Decimal)
            m_VoidAmount = value
            _updatePosForm()
        End Set
    End Property

    ''' <summary>
    '''     All'entrata definisce lo Stornabile massimo
    '''     All'uscita è aggiornato con lo stornato dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property AmountVoidable() As Decimal
        Get
            Return m_VoidableAmount
        End Get
        Set(ByVal value As Decimal)
            m_VoidableAmount = value
            _updatePosForm()
        End Set
    End Property

    ''' <summary>
    '''     All'entrata definisce un elenco di BP già  utilizzati
    '''     in fase di un azione di comando di storno per gestire
    '''     dinamicamente visa Pos software l'lenco di quelli  da
    '''     stornare interagendo con l'utente o per il POS hardware
    '''     presentando anche qui un elenco su un form gestire prima
    '''     dell'invio all'hardware l'elenco di quelli da stornare.
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Friend Property PrefillVoidable() As Dictionary(Of String, PaidEntry)
        Get
            Return m_PrefillVoidable
        End Get
        Set(ByVal value As Dictionary(Of String, PaidEntry))
            m_PrefillVoidable = value
        End Set
    End Property

    ''' <summary>
    '''     Restituisce i numeri decimali per la frazione
    '''     esposta dal protocollo in centesimi per avere
    '''     Euri secondo Argentea
    ''' </summary>
    ''' <returns>La Frazione per i centesimi che usa dal protocollo Argentea</returns>
    Friend ReadOnly Property FractParsing() As Integer
        Get
            Return m_ParseFractMode
        End Get
    End Property

#End Region

#Region "Actions functions Public"

    ''' <summary>
    '''     Esegue il Connect  al  terminale POS
    '''     per farlo attendere sulle operazioni
    '''     oppure emulando con un FORM  che  fa 
    '''     delle chiamate singole  al  SERVIZIO 
    '''     remoto visualizzando i risultati  in 
    '''     una griglia dedicata a video.
    ''' </summary>
    Friend Sub Connect()

        Dim funcName As String = "ProxyArgentea.Connect"


        ' Salvo Sempre che non sia stato già avviato
        If Not m_bWaitActive Then

            ' Istanza della Lib Argentea MONETICA
            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()

            ' Preparo la classe per il Set
            ' di risultati da restituire.
            _DataResponse = New DataResponse(0, 0, 0)

            ' Flag locale che stato attivo
            m_bWaitActive = True

            ' BEHAVIOR
            If m_TypeProxy = enTypeProxy.Service Then

                ' Preparo la risposta (Elaborazione per BP Cartacei)
                m_TypeBPElaborated_CS = enTypeBP.TicketsRestaurant

                '
                ' Creo l'istanza adatta al tipo di azione di  un 
                ' form in emulazione POS Sofwware per il Service.
                '
                frmEmulation = CreateInstanceFormForEmulationOrWaitAction()

                '
                ' Avvio il servizio con la gestione
                ' del POS software tramite service.
                ' Rimane in idle sul form attivo di emulazione.
                '
                StartPosSoftware(frmEmulation)

            Else

                ' Preparo la risposta (Elaborazione per Card con BP a tagli)
                m_TypeBPElaborated_CS = enTypeBP.TicketsCard

                '
                ' Creo l'istanza adatta al tipo di azione di  un 
                ' form in emulazione POS Sofwware per il Service.
                '
                frmEmulation = CreateInstanceFormForEmulationOrWaitAction()


                ' Avvio il pos locale con la gestione
                ' del POS hardware tramite terminale.
                StartPosHardware(frmEmulation)

            End If

            ' Flag locale che stato attivo
            m_bWaitActive = False

        Else
            ' Sollevo l'eccezione
            Throw New Exception(GLB_PROXY_ALREDAY_RUNNING)
        End If

    End Sub

    ''' <summary>
    '''     Chiamate Call ad API di servizio su
    '''     service Argentea remoto.
    ''' </summary>
    ''' <param name="APItoCALL">Il nome dell'api da richiamare presso l'endpoint di argentea</param>
    ''' <returns>Restituisce immeditamente lo stato OK KO o Error di Proxy Argentea <see cref="enProxyStatus"/></returns>
    Friend Function CallAPI(APItoCALL As String) As enProxyStatus
        Dim funcName As String = "ProxyArgentea.CallAPI"

        m_CurrentApiNameToCall = APItoCALL

        ' BEHAVIOR
        If m_TypeProxy = enTypeProxy.Service Then

            ' Le chiamate API sono solo in modalità Service e non Hardware

            ' Preparo la risposta (Elaborazione per BP Cartacei)
            m_TypeBPElaborated_CS = enTypeBP.TicketsRestaurant

            Select Case APItoCALL

                Case "CLOSE"

                    ' Chiamata al metodo remoto  CLOSE
                    ' che chiude la sessione con tutte
                    ' le transazioni che sono state fatte
                    ' presso Argentea, confermandole.
                    Return _API_Close()

                Case Else

                    m_LastStatus = GLB_ERROR_API_NOT_VALID
                    m_LastErrorMessage = "API not present on service argentea"

                    If Not m_SilentMode Then
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)
                    End If

                    Return enProxyStatus.InError


            End Select

        Else

            ' Le chiamate API sono solo in modalità Service e non Hardware

            ' Preparo la risposta (Elaborazione per Card con BP a tagli)
            m_TypeBPElaborated_CS = enTypeBP.TicketsCard

            Return enProxyStatus.InError

        End If


    End Function


    ''' <summary>
    '''     Utilità per parcheggiare in caso  di
    '''     service il form (cioè nasconderlo ma
    '''     non chiuso nella connessione) per le
    '''     attività sotostanti.
    '''     In caso di terminale hardware nasconde
    '''     la finestra di wait avviata in corso di
    '''     comunicazione con la dll COM.
    ''' </summary>
    Friend Sub Park()

        ' BEHAVIOR 
        If m_TypeProxy = enTypeProxy.Service Then

            ' Nasconde la finestra di emulazione pos software avviata al connect.
            frmEmulation.Hide()
            System.Windows.Forms.Application.DoEvents()

        Else

            ' Nasconde la finestra di Wati avviata al connect
            FormHelper.ShowWaitScreen(m_TheModcntr, True, Nothing)

        End If

    End Sub

    ''' <summary>
    '''     Utilità per togliere il proxy corrente 
    '''     dallo stato parcheggiato in caso di
    '''     service il form (cioè rimostrarlo se
    '''     non chiuso nella connessione) per le
    '''     attività sotostanti.
    '''     In caso di terminale hardware rifa vedere
    '''     la finestra di wait avviata in corso di
    '''     comunicazione con la dll COM.
    ''' </summary>
    Friend Sub Unpark()

        ' BEHAVIOR 
        If m_bWaitActive Then

            If m_TypeProxy = enTypeProxy.Service Then

                ' Mostra la finestra di emulazione pos software avviata al connect.
                frmEmulation.Show()
                System.Windows.Forms.Application.DoEvents()

            Else

                ' Mostra la finestra di Wati avviata al connect
                FormHelper.ShowWaitScreen(m_TheModcntr, False, Nothing)

            End If

        End If

    End Sub

    ''' <summary>
    '''     Azzera completamente lo stato del
    '''     Proxy per istanziare una nuova chiamata.
    ''' </summary>
    Friend Sub Close()

        '
        ' Chiudo eventuali finestre per il wait
        ' operatore. 
        '
        FormHelper.ShowWaitScreen(m_TheModcntr, True, Nothing)

        ' Reset conteggio
        Me.Reset()

        ' Deferenziazione (free mem)
        ArgenteaCOMObject = Nothing
        frmEmulation = Nothing
        _DataResponse = Nothing

    End Sub

    ''' <summary>
    '''     Funzione di Utility su chiamate
    '''     esterne combiante.
    ''' </summary>
    Friend Sub Reset()

        ' Reset dello status dei contatori
        _DataResponse = Nothing
        m_ProgressiveCall = 1

        ' Totalizzatori di sessione
        m_TotalPayed_CS = 0
        m_TotalValueExcedeed_CS = 0
        m_TotalBPUsed_CS = 0

        m_TotalVoided_CS = 0
        m_TotalVoidedExcedeed_CS = 0

        '
        m_CurrentTransactionID = String.Empty
        m_CurrentBarcodeScan = String.Empty
        m_bWaitActive = False
        m_ServiceStatus = enProxyStatus.Uninitializated
        m_PayableAmount = 0
        m_PaidAmount = 0
        m_VoidableAmount = 0
        m_VoidAmount = 0

    End Sub

#End Region

#Region "API remote su Servizio Argentea"

    ''' <summary>
    '''     API CLose di Argentea
    ''' </summary>
    ''' <returns><see cref="ArgenteaFunctionsReturnCode"/></returns>
    Private Function _API_Close() As ArgenteaFunctionsReturnCode
        Dim funcName As String = "_API_Close"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' ID Univoco di chiusura operazione
        Dim GUID_Operation As String

        ' Partiamo che non sia OK l'esito su chiamata remota Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' Active to first Argentea COM communication
#If DEBUG_SERVICE = 0 Then

        retCode = ArgenteaCOMObject.CloseTicketBC(
            _GetLastProgressive(), 
            Get_ReceiptNumber, 
            Get_CodeCashDevice, 
            RefTo_MessageOut
        )

#Else
        ''' Per Test
        RefTo_MessageOut = "OK--OPERAZIONI COMPLETATE-----0---" ' <-- x test 
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
#End If

        ' Riprendiamo la Risposta cos' come è stata
        ' data per il log di debug grezza
        LOG_Debug(funcName, "ReturnCode: " & retCode.ToString & ". API: " & m_CurrentApiNameToCall & ". Output: " & RefTo_MessageOut)

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Activation check for BPE with returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)
            Return False

        End If

        ' Riprendiamo la Risposta da protocollo Argentea
        m_LastResponseRawArgentea = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' ** API CLOSE corretamente chiamata ad Argentea
            LOG_Debug(getLocationString(funcName), "API " & m_CurrentApiNameToCall & " successfuly on response with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            ' Risposta da API non OK quindi non validata
            ' sul sistema remoto.
            LOG_Debug(getLocationString(funcName), "API " & m_CurrentApiNameToCall & " remote failed on response from service argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

#End Region

#Region "Function per Gestione del Form di appoggio"

    ''' <summary>
    '''     Nella modalità Service di emulazione
    '''     di un POS software crea un  form  ad
    '''     uso e consumo dell'operatore per scannerizzare 
    '''     i Barcode relativi ai BP da confermare
    '''     tramite i servizi Argentea
    ''' </summary>
    ''' <returns></returns>
    Private Function CreateInstanceFormForEmulationOrWaitAction() As Form

        Dim frmForEmulation As FormBuonoChiaro

        Try

            ' Istanza del form di appggio ad uso 
            ' operatore per l'inserimento per ogni
            ' BP che deve partecipare al pagamento.
            frmForEmulation = m_TheModcntr.GetCustomizedForm(GetType(FormBuonoChiaro), STRETCH_TO_SMALL_WINDOW)

            ' 
            ' Riporto come property al form  da 
            ' visualizzare per una sua gestione
            ' interna il Controller e la Transazione
            '
            frmForEmulation.theModCntr = m_TheModcntr
            frmForEmulation.taobj = m_taobj

            '
            ' BEHAVIOR
            '   In base tipo di azione e per come vogliamo
            '   gestire le varie inputazioni del  caso  lo
            '   gestiamo da qui.
            '
            If m_CommandToCall = enCommandToCall.Void Then

                ' Se gestiamo la modalità Storno delle API di Argentea

                If m_TypeProxy = enTypeProxy.Service Then

                    ' Se per usare la Modalità Service in emulazione Software
                    _SetFormForUseServiceVoid(frmForEmulation)

                ElseIf m_TypeProxy = enTypeProxy.Pos Then

                    ' Se per usare la Modalità Pos in uso con il terminale 
                    _SetFormForUsePosVoid(frmForEmulation)

                End If

            ElseIf m_CommandToCall = enCommandToCall.Payment Then

                ' Se gestiamo la modalità Pagamento con BP delle API di Argentea

                If m_TypeProxy = enTypeProxy.Service Then

                    ' Se per usare la Modalità Service in emulazione Software
                    _SetFormForUseServicePayment(frmForEmulation)

                ElseIf m_TypeProxy = enTypeProxy.Pos Then

                    ' Se per usare la Modalità Pos in uso con il terminale Hardware
                    _SetFormForUsePosPayment(frmForEmulation)

                End If

            End If

            '
            ' Se c'è un Refill di Buoni già
            ' in coda allora le reinserisco
            ' sul form per essere gestiti, 
            ' nel caso hardware con un form
            ' ausiliario di appoggio per l'iterazione
            ' con l'utente a rimuovere i tagli in elenco.
            '
            If m_CommandToCall = enCommandToCall.Void Then
                PrefillVoidableOnPosSoftware(frmForEmulation)
            End If


            ' Ed aggiorno anche il campo sul form per  il totale che rimane.

            If m_CommandToCall = enCommandToCall.Payment Then

                ' Riporto i Valori necessari a questa gestione
                frmForEmulation.Paid = m_TotalPayed_CS.ToString("###,##0.00")
                frmForEmulation.Payable = m_PayableAmount.ToString("###,##0.00")


            ElseIf m_CommandToCall = enCommandToCall.Void Then

                frmForEmulation.Paid = m_TotalVoided_CS.ToString("###,##0.00")
                frmForEmulation.Payable = m_TotalPayed_CS.ToString("###,##0.00")

            End If


        Catch ex As Exception

            ' Sollevo l'eccezione
            Throw New Exception(GLB_ERROR_FORM_INSTANCE, ex)

        End Try

        Return frmForEmulation

    End Function

    ''' <summary>
    '''     Mostra il form di attesa con l'azione da  parte
    '''     di un operatore, gestisce eventi e contesti con
    '''     l'operatore.
    ''' </summary>
    ''' <param name="frmTo">Il Form che è compatibile con questa iterazione</param>
    Private Sub ShowAndIdleOnFormForAction(ByRef frmTo As FormBuonoChiaro, ByVal NotIdle As Boolean)

        '
        ' Mostro il form per la gestione e comunicazione
        ' con il Servizio remoto di convalida su azioni di
        ' Dematerializzazione e Storno.
        '
        frmTo.Show() ' non modal VB Dialog

        ' Dispongo le proprietà del Form  Cassa
        ' ripreso nel Controller globale per la
        ' preparazione a non prendere lo  status
        ' attivo durante la scansione dove si sta 
        ' operando con il controllo locale del form 
        ' che ha la textbox per prendere i codici EAN
        m_TheModcntr.DialogActiv = True
        m_TheModcntr.DialogFormName = frmTo.Text
        m_TheModcntr.SetFuncKeys((False))

        '
        ' (Idle) sul Form Locale
        ' Finestra di dialogo avviata e rimango in idle 
        ' finchè l'operatore non finisce le azioni necessarie.
        '
        frmTo.bDialogActive = True


        ' (Idle)
        If Not NotIdle Then

            Do While frmTo.bDialogActive = True
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()
            Loop

        End If

    End Sub

    ''' <summary>
    '''     Trasforma il form per rimuovere tutti i buoni di tipo
    '''     cartace gestiti dall'emulatore software di POS
    ''' </summary>
    Private Sub _SetFormForUseServiceVoid(frmTo As FormBuonoChiaro)

        frmTo.cmdOK.Text = "Close"
        frmTo.lblBarcode.Text = "Barcode to Void"
        frmTo.lblPayable.Text = "Voidable"
        frmTo.lblPaid.Text = "Voided"

        frmTo.lblRest.Visible = False
        frmTo.txtRest.Visible = False

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere un taglio in
        ' base a dove clicca.
        '
        RemoveHandler frmTo.BarcodeRead, AddressOf BarcodeReadHandler
        RemoveHandler frmTo.BarcodeRemove, AddressOf BarcodeRemoveHandler
        '
        AddHandler frmTo.BarcodeRead, AddressOf BarcodeReadVoidHandler
        AddHandler frmTo.BarcodeRemove, AddressOf BarcodeRemoveVoidHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        AddHandler frmTo.FormClosed, AddressOf CloseOperationHandler


    End Sub

    ''' <summary>
    '''     Trasforma il form per utilizzarlo solo uso  copione
    '''     intanto che c'è la comunicazione con l'hardware con
    '''     un unica label riepilogativa.
    ''' </summary>
    Private Sub _SetFormForUsePosVoid(frmTo As FormBuonoChiaro)

        ' Nascondiamo tutto
        For Each qControl As Control In frmTo.Controls
            qControl.Visible = False
        Next

        ' Per visualizzare la sola label di riepilogo

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere un taglio in
        ' base a dove clicca.
        '
        RemoveHandler frmTo.BarcodeRead, AddressOf BarcodeReadHandler
        RemoveHandler frmTo.BarcodeRemove, AddressOf BarcodeRemoveHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        RemoveHandler frmTo.FormClosed, AddressOf CloseOperationHandler


    End Sub

    ''' <summary>
    '''     Trasforma il form per aggiungere buoni di tipo
    '''     cartace gestiti dall'emulatore software di POS.
    ''' </summary>
    Private Sub _SetFormForUseServicePayment(frmTo As FormBuonoChiaro)

        frmTo.cmdOK.Text = "Close"
        frmTo.lblBarcode.Text = "Barcode to Pay"
        frmTo.lblPayable.Text = "Payable"
        frmTo.lblPaid.Text = "Paid"

        frmTo.lblRest.Visible = True
        frmTo.txtRest.Visible = True


        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere un taglio in
        ' base a dove clicca.
        '
        '
        RemoveHandler frmTo.BarcodeRead, AddressOf BarcodeReadHandler
        RemoveHandler frmTo.BarcodeRemove, AddressOf BarcodeRemoveHandler
        '
        AddHandler frmTo.BarcodeRead, AddressOf BarcodeReadHandler
        AddHandler frmTo.BarcodeRemove, AddressOf BarcodeRemoveHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        AddHandler frmTo.FormClosed, AddressOf CloseOperationHandler

    End Sub

    ''' <summary>
    '''     Trasforma il form per utilizzarlo solo uso  copione
    '''     intanto che c'è la comunicazione con l'hardware con
    '''     un unica label riepilogativa.
    ''' </summary>
    Private Sub _SetFormForUsePosPayment(frmTo As FormBuonoChiaro)

        ' Nascondiamo tutto
        For Each qControl As Control In frmTo.Controls
            qControl.Visible = False
        Next

        ' Per visualizzare la sola label di riepilogo

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere un taglio in
        ' base a dove clicca.
        '
        RemoveHandler frmTo.BarcodeRead, AddressOf BarcodeReadHandler
        RemoveHandler frmTo.BarcodeRemove, AddressOf BarcodeRemoveHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        RemoveHandler frmTo.FormClosed, AddressOf CloseOperationHandler

    End Sub

#End Region

#Region "Parser Functions Privates"

    ''' <summary>
    '''     Riporta per il Parser  corrente
    '''     sul protocollo delle specifiche
    '''     Argentea quale utilizzare e quali
    '''     sono i separatori e le frazioni.
    ''' </summary>
    ''' <returns>Il tipo di parser da usare <see cref="InternalArgenteaFunctionTypes"/> usato dalla funzione ParseReturnString dell'helper CSV </returns>
    Private Function GetSplitAndFormatModeForParsing() As InternalArgenteaFunctionTypes

        ' Behavior
        Dim ParsingMode As InternalArgenteaFunctionTypes

        If m_TypeProxy = enTypeProxy.Pos Then

            If m_CommandToCall = enCommandToCall.Payment Then

                ParsingMode = InternalArgenteaFunctionTypes.MultiPaid_BP
                m_ParseSplitMode = ";"
                m_ParseFractMode = 100

            ElseIf m_CommandToCall = enCommandToCall.Void Then

                ParsingMode = InternalArgenteaFunctionTypes.MultiVoid_BP
                m_ParseSplitMode = ";"
                m_ParseFractMode = 100

            End If

        Else ' service

            If m_CommandToCall = enCommandToCall.Payment Then

                ParsingMode = InternalArgenteaFunctionTypes.SinglePaid_BP
                m_ParseSplitMode = "-"
                m_ParseFractMode = 100

            ElseIf m_CommandToCall = enCommandToCall.Void Then

                ParsingMode = InternalArgenteaFunctionTypes.SingleVoid_BP
                m_ParseSplitMode = "-"
                m_ParseFractMode = 100

            End If

        End If

#If DEBUG_SERVICE = 1 Then

#Else

#End If
        Return ParsingMode

    End Function

    ''' <summary>
    '''     Sui OK di Argentea eseguo il Parsing della Risposta
    '''     per formulare la risposta in esito (Success/Not Success) in codifica da protocollo.
    '''     Mappa gli Attributi su pParams del modulo per lo sbroglio del Flow
    ''' </summary>
    ''' <param name="funcName">Il Nome della Funzione che ha catturato la risposta</param>
    ''' <param name="CSV">Il MessageOut da codificare</param>
    Private Function ParseResponseProtocolArgentea(funcName As String, CSV As String) As ArgenteaFunctionReturnObject

        ' Tipo di codifica generalizzata Argentea wrappatra su un ReturnObject
        Dim objTPTAHelperArgentea(0), ResponseArgentea As ArgenteaFunctionReturnObject
        objTPTAHelperArgentea(0) = New ArgenteaFunctionReturnObject()

        ' Parsiamo la risposta argentea
        If CSV = "ERRORE SOCKET" Then

            ' Codificato Errore Socket 9001
            m_LastErrorMessage = CSV
            m_LastStatus = 9001
            Return New ArgenteaFunctionReturnObject(9001)

        End If

        ' Parser Type (Valorizza anche m_ParseSplitMode)
        Dim ParsingMode As InternalArgenteaFunctionTypes = GetSplitAndFormatModeForParsing()

        ' Parsiamo la risposta argentea per l'azione
        If (Not CSVHelper.ParseReturnString(CSV, ParsingMode, objTPTAHelperArgentea, m_ParseSplitMode, m_ParseFractMode)) Then

            LOG_Debug(getLocationString(funcName), "Parsing Protcol Argentea Fail to Parse 'Message Response' for this " & funcName & " response in MessageOut")

            ' Su Errore di Parsing solleviamo immediatamente l'eccezione per uscire dalla
            ' gestione della comunicazione Argentea.
            Throw New Exception(GLB_ERROR_PARSING)

        Else

            ' RIPORTO SUL FLOW quelli concerni allo Stato di OK Success o KO Error

            ' Risposta Codicficata da Risposta Raw Argentea
            ResponseArgentea = objTPTAHelperArgentea(0)

            ' Log in risposta e decodifica di rpotocoolo effettutato di argentea.
            LOG_Debug(getLocationString(funcName), "Parsed Protcol Argentea 'Message Response' for this " & funcName & " Status: " + ResponseArgentea.Successfull.ToString() + " Response: " + ResponseArgentea.SuccessMessage + ResponseArgentea.ErrorMessage)

            Return ResponseArgentea

        End If

    End Function

    ''' <summary>
    '''     Sui KO di Argentea eseguo il Parsing della Risposta
    '''     per formulare l'errore in codifica da protocollo.
    '''     Mappa gli Attributi su pParams del modulo per lo sbroglio del Flow
    ''' </summary>
    ''' <param name="funcName">Il Nome della Funzione che ha sollevato l'errore</param>
    ''' <param name="CSV">Il MessageOut da codificare</param>
    Private Function ParseErrorAndMapToParams(funcName As String, retCode As ArgenteaFunctionsReturnCode, CSV As String) As ArgenteaFunctionReturnObject

        ' Tipo di codifica generalizzata Argentea wrappatra su un ReturnObject
        Dim objTPTAHelperArgentea(0), ResponseArgentea As ArgenteaFunctionReturnObject
        objTPTAHelperArgentea(0) = New ArgenteaFunctionReturnObject()

        ' Parsiamo la risposta argentea
        If CSV = "ERRORE SOCKET" Then

            m_LastErrorMessage = CSV
            m_LastStatus = 9001
            ' Codificato Errore Socket 9001
            Return New ArgenteaFunctionReturnObject(9001)

        End If

        ' Parser Type (Valorizza anche m_ParseSplitMode)
        Dim ParsingMode As InternalArgenteaFunctionTypes = GetSplitAndFormatModeForParsing()

        If (Not CSVHelper.ParseReturnString(CSV, ParsingMode, objTPTAHelperArgentea, m_ParseSplitMode)) Then

            LOG_Debug(getLocationString(funcName), "BP Parsing Protcol Argentea Fail to Parse 'Error' for this " & funcName & " response in MessageOut")

            ' Su Errore di Parsing solleviamo immediatamente l'eccezione per uscire dalla
            ' gestione della comunicazione Argentea.
            Throw New Exception(GLB_ERROR_PARSING)

        Else

            ' RIPORTO SUL FLOW quelli concerni all'errore

            ' Risposta Codicficata da Risposta Raw Argentea
            ResponseArgentea = objTPTAHelperArgentea(0)

            ' Log in risposta e decodifica di rpotocoolo effettutato di argentea.
            LOG_Debug(getLocationString(funcName), "Parsed Protcol Argentea 'Message Response' for this " & funcName & " Status: " + ResponseArgentea.Successfull.ToString() + " Response: " + ResponseArgentea.SuccessMessage + ResponseArgentea.ErrorMessage)

            ' Risposta Codicficata da Risposta Raw Argentea
            ResponseArgentea = objTPTAHelperArgentea(0)

            ' Quindi mi riporto lo stato dell'operazione
            ' data dalla risposta remota di argentea.
            'pParams.Successfull = False

            ' E riporto nell'ordine corretto il messaggio di stato.
            'pParams.SuccessMessage = ""
            'pParams.ErrorMessage = (ResponseArgentea.Description & " " & ResponseArgentea.Result).Trim()

            ' E per questa specifica Azione fortunatamente
            ' abbiamo il codice di Stato
            'pParams.Status = retCode.ToString() & "-" & ResponseArgentea.CodeResult

            Return ResponseArgentea

        End If

    End Function

#End Region

#Region "Functions Utility"

    ''' <summary>
    '''     Stampa uno scontrino separato sulla cassa
    '''     riportando la transazione di argentea relativa
    '''     ai buoni pasto che sono stati dematerializzati.
    ''' </summary>
    ''' <param name="returnArgenteaObject"></param>
    Private Sub PrintReceipt(returnArgenteaObject As ArgenteaFunctionReturnObject)
        Dim funcName As String = "PrintReceipt"

        ' Riprendo l'helper per Argentea
        Dim objTPTAHelperArgentea As New TPTAHelperArgentea

        ' Creo un Record compatibile per Argenta
        Dim TaArgenteaEMVRec As TPDotnet.IT.Common.Pos.TaArgenteaEMVRec = objTPTAHelperArgentea.ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(m_taobj, returnArgenteaObject)
        If TaArgenteaEMVRec Is Nothing Then
            ' error non vincolante
        End If

        'RegistryHelper.SetLastPaymentADVTransactionIdentifier(ArgenteaFunctionReturnObject(0).TerminalID)
        'RegistryHelper.SetLastPaymentADVTransactionAmount(CInt(MyTaMediaRec.dTaPaidTotal * m_ParseFractMode))
        'RegistryHelper.SetLastPaymentADVTransactionType(OpType)


        ' Creo da 0 una nuova transaction che serve alla stampante
        Dim eftTA As TPDotnet.Pos.TA = objTPTAHelperArgentea.CreateTA(m_taobj, m_TheModcntr, TaArgenteaEMVRec, False)

        If eftTA Is Nothing Then
            'error non vincolante
        End If

        '
        ' ::Opzione:: Stampa su cassa o pos hardware.: 
        '       Definisce il comportamento dello  scontrino
        '       di buoni pagati in pos hardware se stampare 
        '       su stampante legata alla cassa o meno.
        Dim OptPrintLocalReceiptBPPayed As Integer = st_Parameters_Argentea.BPECopies  '   <-- Parametro globale se printare lo scontrino della transazione bp su stampante di cassa locale

        ' In base al parametro globale si decide se stamparlo o meno.
        For I As Integer = 1 To OptPrintLocalReceiptBPPayed

            ' Qui comunichiamo con la stampantina
            ' per stampare tutta la tranazione.
            Dim PAYMENT As Boolean = objTPTAHelperArgentea.PrintReceipt(eftTA, m_TheModcntr)

            If Not PAYMENT Then

                ' Errore non bloccante
                If returnArgenteaObject.Successfull Then

                    ' Signal come errore di stampa ma non bloccante
                    m_LastStatus = GLB_SIGNAL_OPERATOR
                    m_LastErrorMessage = getPosTxtNew(m_TheModcntr.contxt, "POSLevelITCommonPrinterFailed", 0)

                    ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPWARNING)

                End If

            End If

        Next

        Dim lTaCreateNmbr As Short = m_taobj.GetTALine(m_taobj.getLastMediaRecNr).theHdr.lTaCreateNmbr
        TaArgenteaEMVRec.theHdr.lTaCreateNmbr = 0
        TaArgenteaEMVRec.theHdr.lTaRefToCreateNmbr = lTaCreateNmbr
        TaArgenteaEMVRec.bPrintReceipt = st_Parameters_Argentea.BPEPrintWithinTa
        m_taobj.Add(TaArgenteaEMVRec)

    End Sub

#End Region

#Region "** TERMINALE LOCALE POS HARDWARE -> Con i sui metodi nella dll COM in gestione al dispositivo"

    ''' <summary>
    '''     Avvia e mette in attesa il termianale
    '''     hardware collegato alla cassa corente.
    ''' </summary>
    Private Sub StartPosHardware(ByRef frmTo As FormBuonoChiaro)
        Dim funcName As String = "StartPosHardware"

        ' Entrap sull'idle
        Try

            ' Status
            m_ServiceStatus = enProxyStatus.InRunning

            ' (NOT Idle)
            ShowAndIdleOnFormForAction(frmTo, True)

            ' In questa modalità avvio il waitwscreen
            ' modale a pieno schermo per  attendere 
            ' le operazioni dal Pos Locale collegato.
            FormHelper.ShowWaitScreen(m_TheModcntr, False, Nothing, "Attesa su Terminale Locale", "BP Wait")

            ' (Idle)
            Dim _CallHardware As Boolean = CallHardwareWaitMode("StartPosHardware")

            'Do While frmScanCodes.bDialogActive = True
            System.Threading.Thread.Sleep(100)
            System.Windows.Forms.Application.DoEvents()
            'Loop

            ' Emulo l'event Handler come in modalità service
            CloseOperationHandler(Nothing, Nothing)

            ' Dichiaro come concluso correttamente tutto
            If m_ServiceStatus = enProxyStatus.InRunning Then

                ' Se era rimasto in Running e non InError
                ' e tutto è filato liscio e torno con True
                If _CallHardware Then
                    m_ServiceStatus = enProxyStatus.OK
                Else
                    m_ServiceStatus = enProxyStatus.KO
                End If

            Else

                ' Scrive una riga di Log per monitorare....
                LOG_Info(getLocationString(funcName), m_ServiceStatus)

            End If

        Catch ex As Exception

            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED2
            m_LastErrorMessage = "Errore interno non previsto --exception on recalc bp remotes in hardware closed-- (Chiamare assistenza)"

            ' Scrive una riga di Log per l'errore in corso
            ' e lo gestisce in seguito sotto nel finally...
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
            LOG_ErrorInTry(getLocationString(funcName), ex)

        Finally

            Try

                '
                '  Provo a chiudere il Form del POS
                ' hardware se siamo in questa modalità.
                '
                If Not frmTo Is Nothing Then
                    m_TheModcntr.DialogActiv = False
                    m_TheModcntr.DialogFormName = ""
                    m_TheModcntr.SetFuncKeys((True))
                    m_TheModcntr.EndForm()
                    frmTo.Close()
                    frmTo = Nothing
                End If
                '

            Catch ex As Exception

                '
                ' Alla chiusura se uno degli eventi ha solevato 
                ' un eccezione nel consumer lo catturiamo espondendo
                ' il problema di mancato aggiornamento al chiamante.
                '
                If ex.Message = GLB_ERROR_ON_EVENT_DATA Then

                    ' Previsto negli evneti
                    m_LastStatus = GLB_ERROR_FORM_DATA
                    m_LastErrorMessage = "Errore interno alla procedura --exception on event update data-- (Chiamare assistenza)"

                Else

                    ' Non previsto (anomalo)
                    m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
                    m_LastErrorMessage = "Errore interno non previsto --exception on event close service-- (Chiamare assistenza)"

                End If

                ' Status
                m_ServiceStatus = enProxyStatus.InError

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "--" + ex.InnerException.ToString())

                ' Log e segnale non aggiornato in uscita e chiusura 
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            Finally

                ' Chiudo per memory leak con argentea
                ArgenteaCOMObject = Nothing

                ' Effettuo un Dispose forzato per 
                ' la chiusura del form su eccezioni.
                If Not frmTo Is Nothing Then
                    frmTo.Dispose()
                    frmTo = Nothing
                End If

            End Try

        End Try

    End Sub

#End Region

#Region "** SERVICE LOCALE POS FORM -> Con i suoi Handler per il Collect dalle azioni del FORM Locale **"

    ''' <summary>
    '''     Avvia mostrando il Form per il servizio
    '''     POS locale che attende le scansioni dei
    '''     barcode provenienti dall'operatore.
    ''' </summary>
    Private Sub StartPosSoftware(ByRef frmTo As FormBuonoChiaro)
        Dim funcName As String = "StartPosSoftware"

        ' Entrap sull'idle
        Try

            ' Status
            m_ServiceStatus = enProxyStatus.InRunning

            ' (Idle)
            ShowAndIdleOnFormForAction(frmTo, False)

            ' Dichiaro come concluso correttamente tutto
            If m_ServiceStatus = enProxyStatus.InRunning Then

                ' Se era rimasto in Running e non InError
                ' tutto è filato liscio e torno con OK
                m_ServiceStatus = enProxyStatus.OK

            Else

                ' Scrive una riga di Log per monitorare....
                LOG_Info(getLocationString(funcName), m_ServiceStatus)

            End If

        Catch ex As Exception

            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED2
            m_LastErrorMessage = "Errore interno non previsto --exception on recalc bp local emulation software closed-- (Chiamare assistenza)"

            ' Scrive una riga di Log per l'errore in corso
            ' e lo gestisce in seguito sotto nel finally...
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
            LOG_ErrorInTry(getLocationString(funcName), ex)

        Finally

            Try

                '
                '  Provo a chiudere il Form del POS
                ' software se siamo in questa modalità.
                '
                If Not frmTo Is Nothing Then
                    m_TheModcntr.DialogActiv = False
                    m_TheModcntr.DialogFormName = ""
                    m_TheModcntr.SetFuncKeys((True))
                    m_TheModcntr.EndForm()
                    frmTo.Close()
                    frmTo = Nothing
                End If
                '
            Catch ex As Exception

                '
                ' Alla chiusura se uno degli eventi ha solevato 
                ' un eccezione nel consumer lo catturiamo espondendo
                ' il problema di mancato aggiornamento al chiamante.
                '
                If ex.Message = GLB_ERROR_ON_EVENT_DATA Then

                    ' Previsto negli evneti
                    m_LastStatus = GLB_ERROR_FORM_DATA
                    m_LastErrorMessage = "Errore interno alla procedura --exception on event update data-- (Chiamare assistenza)"

                Else

                    ' Non previsto (anomalo)
                    m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
                    m_LastErrorMessage = "Errore interno non previsto --exception on event close service-- (Chiamare assistenza)"

                End If

                ' Status
                m_ServiceStatus = enProxyStatus.InError

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "--" + ex.InnerException.ToString())

                ' Segnalo l'utente
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            Finally

                ' Chiudo per memory leak con argentea
                ArgenteaCOMObject = Nothing

                ' Effettuo un Dispose forzato per 
                ' la chiusura del form su eccezioni.
                If Not frmTo Is Nothing Then
                    frmTo.Dispose()
                    frmTo = Nothing
                End If

            End Try

        End Try
    End Sub

    ''' <summary>
    '''     Handler sulla textbox che riceve in input i barcode in ingresso
    '''     per le azioni di materializzazione dei buoni pasto cartacei.
    ''' </summary>
    ''' <param name="sender">Il controllo textbox</param>
    ''' <param name="barcode">Il barcode stringato nell'evento come parametro di handler</param>
    Protected Overridable Sub BarcodeReadHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeReadHandler"
        Dim Inizializated As Boolean = True
        Dim faceValue As Decimal = 0
        Dim paidValue As Decimal = 0
        Dim retCode As Integer = 0
        Dim formBC As FormBuonoChiaro = Nothing

        '_
        '
        '    Tipi di Repsonse su Protoccolo Argentea
        '    
        '    ->  OPEN TICKET (CallInitialization)       :::         OK--TICKET APERTO-----0--- 
        '    ->  OK su ADD BPC (CallDematerialize)      :::         OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202-- 
        '    ->  KO su ADD BPC (CallDematerialize)      :::         KO-903 - PROGRESSIVO FUORI SEQUENZA-------- 
        '_
        '

        Try

            If TypeOf sender Is FormBuonoChiaro Then

                ' Catturiamo subito il Barcode
                m_CurrentBarcodeScan = barcode

                '
                ' ::Opzione:: Max BP pagabili per vendita.: 
                '       Se il Numero di Buoni pagabili per una vendita
                '       è superiore al numero di buoni passato procediamo
                '       con la sgnazlazione.
                Dim OptPayablesBP As Integer = st_Parameters_Argentea.BP_MaxBPPayableSomeSession

                ' Controllo anche il numero di BP in totale globali alla Vendita corrente
                If OptPayablesBP <> 0 And (((WriterResultDataList.Count + 1) > OptPayablesBP) Or ((m_TotalBPUsed_CS + 1) > OptPayablesBP)) Then

                    m_LastStatus = GLB_OPT_ERROR_NUMEBP_EXCEDEED
                    m_LastErrorMessage = "Il numero di buoni pasto per questa vendita è stato superato!!"

                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)
                    Return

                End If

                ' Controllo se nell'elenco è già stato considerato il BarCode
                If _DataResponse.ContainsBarcode(m_CurrentBarcodeScan) Then

                    ' Status di Errore interno da segnalare
                    m_LastStatus = GLB_INFO_CODE_ALREADYINUSE
                    m_LastErrorMessage = "Il barcode è già stato usato per questa vendita!!"

                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)
                    Return
                End If

                '''     Controllo per il Buono in Corso se il valore
                '''     supera l'importo pagabile, e se lo è  allora
                '''     riporto lo status del pagamento generale come
                '''     in eccesso (sempre se l'opzione lo permette)
                '''     e se lo permette riporto per l'ultimo buono in 
                '''     corso il totale di facciata diverso dal valore effettivo
                '''     di pagato (per scrivefe un media di resto all'uscita)
                If m_TotalValueExcedeed_CS <> 0 Then

                    ' Status di Errore interno da segnalare
                    m_LastStatus = GLB_INFO_IMPORT_ALREADYCOMPLETED
                    m_LastErrorMessage = "L'importo da pagare è già stato completato per questa vendita completare con inoltro!!"

                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)
                    Return

                End If

                '
                ' La prima chiamata Apre la sessione su host remoto Argentea
                '
                If Not m_FirstCall Then

                    ' Chiama per  la  Dematirializzazione
                    ' e incrementa di uno il numero delle
                    ' chiamate interne.
                    FormHelper.ShowWaitScreen(m_TheModcntr, False, sender)
                    Inizializated = Me.CallInitialization(funcName)
                    m_FirstCall = True

                End If

                If Inizializated Then

                    ' Mostriamo il Wait
                    FormHelper.ShowWaitScreen(m_TheModcntr, False, sender)

                    ' Chiama per  la  Dematirializzazione
                    ' e incrementa di uno il numero delle
                    ' chiamate interne.
                    Dim _CallDematerialize As StatusCode = Me.CallDematerialize(funcName)
                    Dim _CallConfirmation As StatusCode = StatusCode.OK

                    '
                    ' ::Opzione:: Operatività.: 
                    '       Se il Totale in ingresso è minore rispetto 
                    '       al valore di facciata del Buono Pasto una volta
                    '       ottenuto dalla materializzazione, opto per troncare su totale.
                    Dim OptAcceptExceeded As Boolean = st_Parameters_Argentea.BP_AcceptExcedeedValues

                    ' Mi conteggio eventuali eccessi su pagato
                    m_TotalValueExcedeed_CS = 0

                    '
                    If _CallDematerialize <> StatusCode.KO Then

                        'SU OK
                        m_TotalValueExcedeed_CS = Math.Min((m_CurrentPaymentsTotal - m_TotalPayed_CS) - m_CurrentValueOfBP, 0)

                        'Su Opzione accetta Valore in eccesso per resto
                        If OptAcceptExceeded Then

                            '
                            '       --> Accetta eccesso su Totale da Pagare
                            '               Alla fine scrive il media don i due riporti
                            '               concludendo il pagamento a totale.
                            '

                        Else

                            '
                            '       --> Non Accetta eccesso su Totale da Pagare
                            '               Richiama Argentea per fare l'annullo
                            '               alla demateriliazzazione fatta in precedenza
                            '

                            If m_TotalValueExcedeed_CS < 0 Then

                                ' Status di Errore interno da segnalare
                                m_LastStatus = GLB_OPT_ERROR_VALUE_EXCEDEED
                                m_LastErrorMessage = "Il Valore del Buono Pasto eccede il valore rispetto al totale (non è possibile dare resto)"

                                ' Log locale
                                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "-- Transaction Dematerialize Argentea ::KO:: Excedeed")

                                ' Segnalo Operatore di Cassa
                                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                                ' Immediatamente annullo verso il sistema argnetea l'operazione
                                ' Per rimuoverlo tramite il metodo stesso per l'annullo
                                m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
                                Me.BarcodeRemoveHandler(sender, m_CurrentBarcodeScan)
                                m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

                                ' Ripristiniamo l'importo d'eccedenza
                                m_TotalValueExcedeed_CS = 0

                                ' Torno all'inseirmento eventualemnete per optare su altri 
                                ' buoni pasto cliente con importi appropriati al completamento.
                                Return

                            End If

                        End If

                    End If

                    If _CallDematerialize = StatusCode.CONFIRMREQUEST Then

                        ' Chiama per  la  Dematirializzazione
                        ' e incrementa di uno il numero delle
                        ' chiamate interne.
                        _CallConfirmation = Me.CallConfirmOperation(funcName, "dematerialization")

                        If _CallConfirmation = StatusCode.KO Then

                            ' Log locale (Non confermato in demat)
                            LOG_Info(funcName, "Transaction Dematerialize on Argentea ::KO:: ON CONFIRM")

                        End If

                    End If

                    ' Una Volta richiamata la demateriliazzione 
                    ' ed eventuale conferma ed hanno dato esito positivo.
                    If _CallDematerialize <> StatusCode.KO And _CallConfirmation = StatusCode.OK Then

                        ' Riprendo per il BP i valori
                        ' e il pagato reale
                        faceValue = m_CurrentValueOfBP
                        paidValue = m_CurrentValueOfBP + m_TotalValueExcedeed_CS

                        ' Aggiungo in una collection specifica in uso
                        ' interno l'elemento Buono appena accodato in
                        ' modo univoco rispetto al suo BarCode.
                        Dim ItemNew As PaidEntry = WriterResultDataList.NewPaid(
                            m_CurrentBarcodeScan, m_CurrentTerminalID
                        )
                        ItemNew.Value = paidValue.ToString("###,##0.00")
                        ItemNew.FaceValue = faceValue.ToString("###,##0.00")
                        ItemNew.Emitter = ""

                        ' Per l'azione sull'elenco corrente mi
                        ' riprendo il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        m_TotalPayed_CS += m_CurrentValueOfBP
                        m_TotalBPUsed_CS += 1                         ' <-- Conteggio numero di bpc usati in local per ogni ingresso sulla vendita

                        Try
                            ' Riprendo il sender che p il Form
                            ' dove voglio aggiungere alla lista
                            ' l'n elemento appena validato.
                            formBC = TryCast(sender, FormBuonoChiaro)
                            If formBC Is Nothing Then Throw New Exception(GLB_ERROR_FORM_INSTANCE)

                            ' Aggiungo l'elemento al controllo Griglia
                            formBC.PaidEntryBindingSource.Add(ItemNew) ' New PaidEntry(m_CurrentBarcodeScan, paidValue.ToString("###,##0.00"), faceValue.ToString("###,##0.00"), ""))

                            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
                            formBC.Paid = m_TotalPayed_CS.ToString("###,##0.00")

                        Catch ex As Exception
                            Throw New Exception(GLB_ERROR_FORM_INSTANCE, ex)
                        End Try

                    Else

                        ' Errata Dematerializzione o Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                        ' Log locale
                        LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Remote")

                        ' Messaggio all'utente
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Remote")

                    ' Messaggio all'utente
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto
                Throw New Exception(GLB_ERROR_FORM_INSTANCE)
            End If

        Catch ex As Exception

            FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)
            SetExceptionsStatus(funcName, ex)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

            ' Riporto la firstcall a false
            ' per le istanze successive.
            m_FirstCall = False

            ' Svuoto il controllo del barcode
            If Not formBC Is Nothing Then formBC.txtBarcode.Text = String.Empty

        End Try

    End Sub

    ''' <summary>
    '''     Hanlder sulla gridbox che riceve in input il barcode selezionato
    '''     per le azioni di materializzazione del buono pasto cartaceo.
    ''' </summary>
    ''' <param name="sender">Istanza del form sugll'handler degli eventi</param>
    ''' <param name="barcode">Il Barcode all'evneto sulla scansione</param>
    ''' <param name="UndoBPCForExcedeed">Per irutilizzare la funzione su opzioni di Eccesso su totale previo Resto valido non valido</param>
    Protected Overridable Sub BarcodeRemoveHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeRemoveHandler"
        Dim Inizializated As Boolean = True
        Dim faceValue As Decimal = 0
        Dim paidValue As Decimal = 0
        Dim retCode As Integer = 0
        Dim formBC As FormBuonoChiaro = Nothing

        '_
        '
        '    Tipi di Repsonse su Protoccolo Argentea
        '    
        '    ->  OPEN TICKET (CallInitialization)       :::         OK--TICKET APERTO-----0--- 
        '    ->  OK su ADD BPC (CallDematerialize)      :::         OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202-- 
        '    ->  KO su ADD BPC (CallDematerialize)      :::         KO-903 - PROGRESSIVO FUORI SEQUENZA-------- 
        '_
        '

        Try

            If TypeOf sender Is FormBuonoChiaro Then

                ' Cattueriamo subito il Barcode
                m_CurrentBarcodeScan = barcode

                'Controllo se nell'elenco è già stato considerato il BarCode
                If Not m_FlagUndoBPCForExcedeed And Not _DataResponse.ContainsBarcode(m_CurrentBarcodeScan) Then

                    m_LastStatus = GLB_INFO_CODE_NOTPRESENT
                    m_LastErrorMessage = "Il BP non è presente tra le scelte possibili!!"

                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)
                    Return
                End If

                If Inizializated Then

                    FormHelper.ShowWaitScreen(m_TheModcntr, False, sender)

                    ' Chiama per l'anullamento di uno già 
                    ' Dematirializzato  e  incrementa  di 
                    ' uno il numero delle il numero delle 
                    ' chiamate interne.
                    Dim _CallUndoDematerialize As StatusCode = Me.CallReverseMaterializated(funcName)
                    Dim _CallConfirmation As StatusCode = StatusCode.OK

                    If _CallUndoDematerialize = StatusCode.CONFIRMREQUEST Then

                        ' Chiama  per  conferma  l'Annnullamento   
                        ' di uno già  esistente Dematerializzato  
                        ' in precedenza  e  incrementa di uno il 
                        ' numero delle chiamate interne.
                        _CallConfirmation = Me.CallConfirmOperation(funcName, "reverse")

                        If _CallConfirmation = StatusCode.KO Then

                            ' Log locale (Non confermato in demat per storno)
                            LOG_Info(funcName, "Transaction Reverse Dematerialize on Argentea ::KO:: ON CONFIRM")

                        End If

                    End If

                    ' Una Volta richiamata la demateriliazzione 
                    ' ed eventuale conferma ed hanno dato esito positivo.
                    If _CallUndoDematerialize <> StatusCode.KO And _CallConfirmation = StatusCode.OK Then

                        ' Argormento Opzione per Opzione 
                        ' su Flow operatore se non accetta
                        ' Sulla griglia e il form non deve
                        ' fare altro dato che non è stato aggiunto.
                        If m_FlagUndoBPCForExcedeed Then
                            Return
                        End If

                        ' Rimuovo dalla collection specifica in uso
                        ' interno l'elemento Buono da annullare individuandolo
                        ' in modo univoco rispetto al suo BarCode con cui era 
                        ' stato registrato all'aggiunta dell'handler di ADD.
                        For Each itm As PaidEntry In WriterResultDataList
                            If itm.Barcode = m_CurrentBarcodeScan Then
                                WriterResultDataList.Remove(itm)
                                Exit For
                            End If
                        Next

                        ' 
                        m_TotalBPUsed_CS -= 1                         ' <-- Conteggio numero di bpc usati in local per ogni ingresso sulla vendita

                        ' Per il Form in azione corrente mi
                        ' aggiorno il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        faceValue = faceValue
                        paidValue = m_CurrentValueOfBP
                        m_TotalPayed_CS -= paidValue

                        Try
                            ' Riprendo il sender che p il Form
                            ' dove voglio aggiungere alla lista
                            ' l'n elemento appena validato.
                            formBC = TryCast(sender, FormBuonoChiaro)
                            If formBC Is Nothing Then Throw New Exception(GLB_ERROR_FORM_INSTANCE)

                            ' Sul Form rimuovo dalla griglia l'elemento
                            formBC.PaidEntryBindingSource.RemoveCurrent()

                            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
                            formBC.Paid = m_TotalPayed_CS.ToString("###,##0.00")

                        Catch ex As Exception
                            Throw New Exception(GLB_ERROR_FORM_INSTANCE, ex)
                        End Try

                    Else

                        ' Errata Reverse per Dematerializzione o Reverse Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                        ' Log locale
                        LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Reverse Demat Argentea ::KO:: Local")

                        ' Messaggio all'utente
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Reverse Demat Argentea ::KO:: Not Intializated")

                    ' Messaggio all'utente
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto
                Throw New Exception(GLB_ERROR_FORM_INSTANCE)
            End If

        Catch ex As Exception

            FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)
            SetExceptionsStatus(funcName, ex)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

            ' Svuoto il controllo del barcode
            If Not m_FlagUndoBPCForExcedeed And Not formBC Is Nothing Then formBC.txtBarcode.Text = String.Empty

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla gridbox che riceve in input il barcode selezionato
    '''     per le azioni di storno del buono pasto cartaceo.
    ''' </summary>
    ''' <param name="sender">Istanza del form sugll'handler degli eventi</param>
    ''' <param name="barcode">Il Barcode all'evneto sulla scansione</param>
    ''' <param name="UndoBPCForExcedeed">Per irutilizzare la funzione su opzioni di Eccesso su totale previo Resto valido non valido</param>
    Protected Overridable Sub BarcodeReadVoidHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeReadVoidHandler"
        Dim formBC As FormBuonoChiaro = Nothing

        ' Immediatamente annullo verso il sistema argnetea l'operazione
        ' Per rimuoverlo tramite il metodo stesso per l'annullo
        m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
        Me.BarcodeRemoveHandler(sender, barcode)
        m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

        _updateVoidedForm(sender)

    End Sub

    ''' <summary>
    '''     Handler sulla gridbox che riceve in input la riga con barcode selezionato
    '''     per le azioni di storno del buono pasto cartaceo.
    ''' </summary>
    ''' <param name="sender">Istanza del form sugll'handler degli eventi</param>
    ''' <param name="barcode">Il Barcode all'evneto sulla scansione</param>
    ''' <param name="UndoBPCForExcedeed">Per irutilizzare la funzione su opzioni di Eccesso su totale previo Resto valido non valido</param>
    Protected Overridable Sub BarcodeRemoveVoidHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeRemoveVoidHandler"

        ' Immediatamente annullo verso il sistema argnetea l'operazione
        ' Per rimuoverlo tramite il metodo stesso per l'annullo
        m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
        Me.BarcodeRemoveHandler(sender, barcode)
        m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

        _updateVoidedForm(sender)

    End Sub

    Private Sub _updateVoidedForm(ByRef sender As Object)
        Dim formBC As FormBuonoChiaro = Nothing
        Dim _revoke As Boolean = False

        ' Aggiorno i dati di Storno
        formBC = CType(sender, FormBuonoChiaro)


        ' Rimuovo dalla collection specifica in uso
        ' interno l'elemento Buono da annullare individuandolo
        ' in modo univoco rispetto al suo BarCode con cui era 
        ' stato registrato all'aggiunta dell'handler di ADD.
        For Each itm As PaidEntry In WriterResultDataList
            If itm.Barcode = m_CurrentBarcodeScan Then
                itm.Emitter = "[[VOIDED]]"
                _revoke = True
                Exit For
            End If
        Next

        If Not _revoke Then

            m_LastStatus = GLB_INFO_CODE_NOTPRESENT
            m_LastErrorMessage = "Il BP non è presente come pagato!!"

            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)
            Return

        End If

        ' Se è stato rimosso correttametne procediamo

        m_TotalBPUsed_CS -= 1                         ' <-- Conteggio numero di bpc usati in local per ogni rimosso

        ' Per il Form in azione corrente mi
        ' aggiorno il Totale da Pagare rispetto a
        ' quelli già in elenco
        m_TotalPayed_CS -= m_CurrentValueOfBP
        m_TotalVoided_CS += m_CurrentValueOfBP

        ' Riportiamo aggiornato il form

        Try
            ' Riprendo il sender che p il Form
            ' dove voglio aggiungere alla lista
            ' l'n elemento appena validato.
            formBC = TryCast(sender, FormBuonoChiaro)
            If formBC Is Nothing Then Throw New Exception(GLB_ERROR_FORM_INSTANCE)

            ' Sul Form rimuovo dalla griglia l'elemento
            For Each itm As PaidEntry In formBC.PaidEntryBindingSource
                If itm.Barcode = m_CurrentBarcodeScan Then
                    formBC.PaidEntryBindingSource.Remove(itm)
                    Exit For
                End If
            Next

            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
            formBC.Paid = m_TotalVoided_CS.ToString("###,##0.00")
            formBC.Payable = m_TotalPayed_CS.ToString("###,##0.00")

        Catch ex As Exception
            Throw New Exception(GLB_ERROR_FORM_INSTANCE, ex)
        End Try


    End Sub

#End Region

#Region "Functions Private per Hardware Mode"

    ''' <summary>
    '''     Inizializza la Sessione verso Argentea
    '''     e parte la numerazione interna delle chiamate
    '''     da 1
    ''' </summary>
    Private Function CallHardwareWaitMode(funcName As String) As Boolean

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = String.Empty

        CallHardwareWaitMode = False

        ' Partiamo che non sia OK l'esito su chiamata remota Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        '
        '   OUT
        '   L'Id di transazione  recuperata
        '   dopo le chiamate verso Argentea.
        '   (Passato alla dll COM di Argentea e fillato dalla stessa)
        '
        Dim RefTo_Transaction_Identifier As String = String.Empty

        '
        ' (Idle) sulla chiamata diretta al POS 
        ' chiamato dalla funzione API di 
        ' argentea per avviare una  sessione
        ' sul POS locale di pagamento.
        '
        '   amount =                 ''' L'importo per avviare il POS a farsi pagare in BP l'importo dettato
        '

#If DEBUG_SERVICE = 0 Then

        ' (Idle)
        If m_CommandToCall = enCommandToCall.Payment Then

            retCode = ArgenteaCOMObject.PaymentBPE(
                    CInt(m_PayableAmount * m_ParseFractMode),
                     RefTo_Transaction_Identifier,
                     RefTo_MessageOut
                 )

        ElseIf m_CommandToCall = enCommandToCall.Void Then

            retCode = ArgenteaCOMObject.VoidBPE(
                    CInt(m_VoidableAmount * m_ParseFractMode),
                     RefTo_Transaction_Identifier,
                     RefTo_MessageOut
                 )

        End If
#Else

        ''' Per Test
        If m_CommandToCall = enCommandToCall.Payment Then
            ''' 
            RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;2|5|1020|1|414;104;PELLEGRINI;  PAGAMENTO BUONO PASTO " ' <-- x test 
        ElseIf m_CommandToCall = enCommandToCall.Void Then
            ''' 
            RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;2|4|1020|1|414;104;PELLEGRINI;  PAGAMENTO BUONO PASTO " ' <-- x test 
        End If
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
#End If

        ' Riprendiamo la Risposta così comè stata
        ' data per il log di debug grezza
        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". BP: Hardware Output: " & RefTo_MessageOut)

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Activation check for BPE with returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)
            Return False

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            m_LastResponseRawArgentea = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If m_LastResponseRawArgentea.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                'IncrementProgressiveCall()

                '
                ' A differenza del Software  Creo  voci
                ' di TA quanti sono stati inoltrati nel
                ' dispositivo.
                '
                m_TotalBPUsed_CS = m_LastResponseRawArgentea.NumBPEvalutated        ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento
                m_TotalPayed_CS = m_LastResponseRawArgentea.Amount                  ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente
                m_TotalValueExcedeed_CS = 0                                         ' <-- ?? TODO:: Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata

                ' Riprendo l'elenco riportato dall'hardware
                ' per ogni taglio e colloco ricopiandolo il 
                ' pezzo interessato
                For Each itm As Object In m_LastResponseRawArgentea.ListBPsEvaluated

                    ' Questo dall'hardware non c'è l'abbiamo
                    ' e portiamo un code contatore
                    Dim paidValue As Decimal = itm.Value
                    Dim faceValue As Decimal = itm.Value

                    m_CurrentBarcodeScan = itm.Key
                    m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                    ' Aggiungo in una collection specifica in uso
                    ' interno l'elemento Buono appena accodato in
                    ' modo univoco rispetto al suo BarCode.
                    Dim ItemNew As PaidEntry = WriterResultDataList.NewPaid(
                            m_CurrentBarcodeScan, m_CurrentTerminalID
                        )
                    ItemNew.Value = paidValue.ToString("###,##0.00")
                    ItemNew.FaceValue = faceValue.ToString("###,##0.00")
                    ItemNew.Emitter = RefTo_Transaction_Identifier

                Next


                ' ** ATTESA COMPLETATA e corretamente chiamata vs Hardware Terminal POS
                LOG_Debug(getLocationString(funcName), "BP comunication with terminal pos successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
                Return True

            Else

                ' Non inizializzata da parte di Argentea per
                ' errore remoto in risposta a questo codice.
                LOG_Debug(getLocationString(funcName), "BPE comunication remote failed on first call to terminal argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
                Return False

            End If

        End If

    End Function

#End Region

#Region "EventOut CollectData Close"

    ''' <summary>
    '''     Handdler dell'evento chiave sia quando è in modalità service che pos
    '''     per convogliare i dati in ingresso e restituirli al chiamante,.
    ''' </summary>
    ''' <param name="sender">Il form del Pos software o il COM del service pos locale hardware</param>
    ''' <param name="e"></param>
    Private Sub CloseOperationHandler(sender As Object, e As FormClosedEventArgs)
        Dim funcName As String = "CloseOperationHandler"

        Try

            '
            '   Print Last Receipt (Solo in pagamento e solo per quelli POS Hardware)
            '
            If m_CommandToCall = enCommandToCall.Payment And m_TypeProxy = enTypeProxy.Pos Then

                If Not m_LastResponseRawArgentea Is Nothing Then
                    PrintReceipt(m_LastResponseRawArgentea)
                End If

            End If

        Catch ex As Exception

            ' Log locale (Errore di reprint dello scontrino non bloccante)
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Printer for print recipient in proxy Argentea: Hardware Output")
            LOG_ErrorInTry(getLocationString("ProxyArgentea"), ex)

        End Try

        Try

            '
            '   Evento chiave di chiusura
            '
            If m_CommandToCall = enCommandToCall.Payment Then

                RaiseEvent Event_ProxyCollectDataTotalsAtEnd(Me, _DataResponse)

            ElseIf m_CommandToCall = enCommandToCall.Void Then

                RaiseEvent Event_ProxyCollectDataVoidedAtEnd(Me, _DataResponse)

            End If

        Catch ex As Exception

            ' Intercettiamo l'errore per il contesto probabilmente
            ' erchè il consumer non l'ha fatto per suo conto, quindi
            ' rimane che per noi il consumer con i dati non è aggioranto.
            Throw New Exception(GLB_ERROR_ON_EVENT_DATA, ex)

        End Try

    End Sub

#End Region

#Region "Functions Private per Emulation Pos Software Service mode"

    Private _flagCallOnetimeResetIncrement As Boolean = False

    ''' <summary>
    '''     Inizializza la Sessione verso Argentea
    '''     e parte la numerazione interna delle chiamate
    '''     da 1
    ''' </summary>
    Private Function CallInitialization(funcName As String) As Boolean

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        CallInitialization = False

        ' Partiamo che non sia OK l'esito su chiamata remota Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' Active to first Argentea COM communication
#If DEBUG_SERVICE = 0 Then
        retCode = ArgenteaCOMObject.OpenTicketBC(
            m_ProgressiveCall,
            Get_ReceiptNumber,
            Get_CodeCashDevice,
            RefTo_MessageOut
        )
#Else
        ''' Per Test
        If _flagCallOnetimeResetIncrement = False Then
            ' 1° Tentativo
            RefTo_MessageOut = "KO-903-Sequenza non valida-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"            ' <-- x test su questo signal
            retCode = ArgenteaFunctionsReturnCode.KO

        Else
            ' 2° tenttivo
            RefTo_MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 
            retCode = ArgenteaFunctionsReturnCode.OK

        End If
        ''' to remove:
#End If

        ' Riprendiamo la Risposta cos' come è stata
        ' data per il log di debug grezza
        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". BP: " & m_CurrentBarcodeScan & ". Output: " & RefTo_MessageOut)

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Tento prima di uscire di fare un Reset dell'increment
            ' remoto per vedere se possiamo prendere a richiamarlo.
            If _flagCallOnetimeResetIncrement = False Then
                If CallResetIncrement(m_LastResponseRawArgentea) = False Then
                    Return False
                Else
                    _flagCallOnetimeResetIncrement = True
                    Dim _Result As Boolean = CallInitialization(funcName)
                    _flagCallOnetimeResetIncrement = False
                    Return _Result
                End If
            Else
                ' 2° e ultimo tentativo altrimenti è fallito punto e basta
                ' ** NON INIZIALIZZATA e tentativo eventuale riallineamento FALLITO
                m_LastStatus = GLB_FAILED_INITIALIZATION
                m_LastErrorMessage = "Inizializzazione e tantivo di riallineamento fallito"
                LOG_Debug(getLocationString(funcName), "BP inizialization " & m_CurrentBarcodeScan & " unsuccessfuly call with message " & m_LastResponseRawArgentea.SuccessMessage)
                Return False
            End If

        End If

        ' Riprendiamo la Risposta da protocollo Argentea
        m_LastResponseRawArgentea = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
            _IncrementProgressiveCall()

            ' ** INIZIALIZZATA e corretamente chiamata ad Argentea
            LOG_Debug(getLocationString(funcName), "Inizialization " & m_CurrentBarcodeScan & " successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            ' Non inizializzata da parte di Argentea per
            ' errore remoto in risposta a questo codice.
            LOG_Debug(getLocationString(funcName), "Inizialization " & m_CurrentBarcodeScan & " remote failed on first call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

    ''' <summary>
    '''     Esegue un riallineamento del contatore
    '''     remoto sul servizio Argentea se la risposta
    '''     rispetto all'ultima chiamata mi dà un 903
    ''' </summary>
    ''' <param name="LastResponse">La risposta che si deve prendere in analisi per avere il comportamento rispetto a 903 come staus<see cref="ArgenteaFunctionReturnObject"/></param>
    ''' <returns>True/False se si deve ritentare dopo il riallineamento oppure False se ha fallito o se diverso da 903 per segnalazione e inoltro dell'errore</returns>
    Private Function CallResetIncrement(ByRef LastResponse As ArgenteaFunctionReturnObject) As Boolean
        Dim funcName As String = "CallResetIncrement"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Partiamo che non sia OK l'esito su chiamata remota Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        If LastResponse.CodeResult = "903" Then
#If DEBUG_SERVICE = 0 Then
            retCode = ArgenteaCOMObject.Riallineamento( RefTo_MessageOut )
#Else
            ''' Per Test
            RefTo_MessageOut = "OK--RIALLINEATO-----0---" ' <-- x test 
            retCode = ArgenteaFunctionsReturnCode.OK
            ''' to remove:
#End If

            ' Riprendiamo la Risposta cos' come è stata
            ' data per il log di debug grezza
            LOG_Debug(getLocationString(funcName), "ReturnCode of reset: " & retCode.ToString & ". BP: " & m_CurrentBarcodeScan & ". Output: " & RefTo_MessageOut)

            If retCode <> ArgenteaFunctionsReturnCode.OK Then

                ' Su risposta da COM  in  negativo
                ' in ogni formatto il returnString
                ' ma con la variante che già mi filla
                ' l'attributro ErrorMessage
                m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

                ' Non inizializzata su Errori di comunicazione
                ' o per risposta remota data da Argentea KO.
                LOG_Error(getLocationString(funcName), "Reset increment for BP with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)
                Return False

            Else

                ' Il reset è andato a buon fine possiamo riprendere dall'inizitialization
                Return True

            End If

        Else

            ' Altre Risposte Negative da Chiamata Precedente

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Activation check for BPC with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)
            Return False

        End If

    End Function

    ''' <summary>
    '''     Esegue la chiamata di Dematerializzazione secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    Private Function CallDematerialize(funcName As String) As StatusCode

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' CSV for Argumet return Status Call to Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallDematerialize = StatusCode.KO

#If DEBUG_SERVICE = 0 Then
        ' Active to first Argentea COM communication                                **** DEMATERIALIZZAZIONE
        retCode = ArgenteaCOMObject.DematerializzazioneBP(
                    GetCodifiqueReceipt(TypeCodifiqueProtocol.Dematerialization),
                    RefTo_MessageOut
                )
#Else
        ''' Per Test questo è il suio CSV
        RefTo_MessageOut = "OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--"    ' <-- x test 
        'RefTo_MessageOut = "KO-3-Buono pasto gia' rientrato-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"       ' <-- x test su questo signal
        'RefTo_MessageOut = "KO-903-Sequenza non valida-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"            ' <-- x test su questo signal
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
#End If

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Dematerialization for BP with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)

            ' Esco dal  flow immediatamente
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
            m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
            Return CallDematerialize

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            m_LastResponseRawArgentea = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If m_LastResponseRawArgentea.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                _IncrementProgressiveCall()

                ' Se la risposta argenta richiede un ulteriore 
                ' conferma allora procedo ad uscire per il flow.
                If m_LastResponseRawArgentea.RequireCommit Then

                    ' ** DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP dematirializated with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                    ' RICHIESTO CONFERMA
                    m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                    m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                    CallDematerialize = StatusCode.CONFIRMREQUEST
                    Return CallDematerialize
                Else

                    ' ** DEMATERIALIZZATO corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP dematerializated " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                    ' COMPLETATO
                    m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                    m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                    CallDematerialize = StatusCode.OK
                    Return CallDematerialize
                End If

            Else

                ' Non dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BP dematerializated " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' NON EFFETTUATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                CallDematerialize = StatusCode.KO
                Return CallDematerialize
            End If

        End If

    End Function

    ''' <summary>
    '''     Esegue la chiamata di Reverse da uno già Dematerializzato secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    Private Function CallReverseMaterializated(funcName As String) As StatusCode

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' CSV for Argumet return Status Call to Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallReverseMaterializated = StatusCode.KO

#If DEBUG_SERVICE = 0 Then
        ' Active to first Argentea COM communication                                **** ANNULLO BUONO GIA' MATERIALIZZATO
        retCode = ArgenteaCOMObject.ReverseTransactionDBP(
                    GetCodifiqueReceipt(TypeCodifiqueProtocol.Reverse),
                    RefTo_MessageOut
                )
#Else
        ''' Per Test
        RefTo_MessageOut = "OK-0 - BUONO STORNATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
#End If

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Reverse Dematerialization for BPC with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)

            ' Esco dal  flow immediatamente
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
            CallReverseMaterializated = StatusCode.KO
            Return CallReverseMaterializated
        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            m_LastResponseRawArgentea = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If m_LastResponseRawArgentea.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                _IncrementProgressiveCall()

                ' Se la risposta argenta richiede un ulteriore 
                ' conferma allora procedo ad uscire per il flow.
                If m_LastResponseRawArgentea.RequireCommit Then

                    ' ** REVERSE SU DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP reverse dematirializated with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                    ' RICHIESTO CONFERMA
                    m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                    m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                    CallReverseMaterializated = StatusCode.CONFIRMREQUEST
                    Return CallReverseMaterializated
                Else

                    ' ** REVERSE SU DEMATERIALIZZATO corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP reverse dematirializated " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                    ' COMPLETATO
                    m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                    m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                    CallReverseMaterializated = StatusCode.OK
                    Return CallReverseMaterializated
                End If

            Else

                ' Non reverse su dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BP reverse dematerializated " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' NON EFFETTUATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                CallReverseMaterializated = StatusCode.KO
                Return CallReverseMaterializated
            End If

        End If

    End Function

    ''' <summary>
    '''     Esegue una chiamata con protocollo di Conferma verso
    '''     Argentea per confirm su Dematerializzazione o Reverse.
    ''' </summary>
    ''' <param name="funcName">Nome della funzione chiamante</param>
    ''' <returns></returns>
    Private Function CallConfirmOperation(funcName As String, sConfirmOperation As String) As StatusCode

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' CSV for Argumet return Status Call to Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallConfirmOperation = StatusCode.KO

#If DEBUG_SERVICE = 0 Then

        ' Active to first Argentea COM communication                                **** CONFERMA DEMATERIALIZZAZIONE o REVERSE
        retCode = ArgenteaCOMObject.CommitTransactionDBP(
                    GetCodifiqueReceipt(TypeCodifiqueProtocol.Confirm),
                    RefTo_MessageOut
                )
#Else
        ''' Per Test
        RefTo_MessageOut = "OK-0 - BUONO CONFERMATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
#End If

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            m_LastResponseRawArgentea = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Confirm " & sConfirmOperation & " for BP with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)

            ' Esco dal  flow immediatamente
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
            CallConfirmOperation = StatusCode.KO
            Return CallConfirmOperation
        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            m_LastResponseRawArgentea = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If m_LastResponseRawArgentea.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                _IncrementProgressiveCall()

                ' ** CONFIRM su REVERSE o DEMATERIALIZZATO effettuata corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP confirm " & sConfirmOperation & " for " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' COMPLETATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                CallConfirmOperation = StatusCode.OK
                Return CallConfirmOperation
            Else

                ' Non confirm su reverse o dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BP confirm " & sConfirmOperation & " for " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' NON EFFETTUATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                m_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
                CallConfirmOperation = StatusCode.KO
                Return CallConfirmOperation
            End If

        End If

    End Function


    ''' <summary>
    '''     In service aggiorna il form visualizzato
    '''     in emulazione per l'attesa dei barcode.
    ''' </summary>
    Private Sub _updatePosForm()

        ' In modalità emulatore software del POS
        ' aggiorno il form con gli stessi dati.
        If m_TypeProxy = enTypeProxy.Service Then

            '
            ' In questa modalità avvio il form 
            ' preparandolo al totale e il pagabile.
            '
            If Not frmEmulation Is Nothing AndAlso TypeOf frmEmulation Is FormBuonoChiaro Then
                CType(frmEmulation, FormBuonoChiaro).Paid = m_PaidAmount
                CType(frmEmulation, FormBuonoChiaro).Payable = m_PayableAmount
            End If

        End If

    End Sub

    Private Function ValidationVoucherRequest(barcode As String) As Boolean
        'Logic Comunication Barcode at Argentea Supplier

        ValidationVoucherRequest = False


    End Function

#End Region

#Region "Functions per la gestione Exception e Error"

    ''' <summary>
    '''     Negli eventi utilizzare questo metodo per impostare
    '''     un eccezione non gestita o errori di cui si vuole che
    '''     il flow in corso sia interrotto regolarmente.
    ''' </summary>
    ''' <param name="funcname">Il nome della funzione che vuole gestire lo stato dell'errore</param>
    ''' <param name="status">Lo Status con cui si deve innescare l'eccezione</param>
    ''' <param name="errorMessage">Il Messaggio di Errore da mostrare</param>
    ''' <param name="viewMessageErr">Se si vuole visualizzare il messaggio di errore in corso</param>
    Friend Sub SetStatusInError(funcname As String, status As String, errorMessage As String, viewMessageErr As Boolean)

        ' Se per qualche motivo o perchè manca il file di trasformazione
        ' o per errori in esecuzione non applica il filtro esco dalla gestione.
        If viewMessageErr Then
            msgUtil.ShowMessage(m_TheModcntr, errorMessage, "LevelITCommonModArgentea_" + status, PosDef.TARMessageTypes.TPSTOP)
        End If

        ' Definisco lo stato generale del proxy in esecuzione
        ' come stato di errore per interrompere nel flow 
        ' eventuali prosegqui.
        m_ServiceStatus = enProxyStatus.InError

    End Sub

    ''' <summary>
    '''     Gestisce per segnalarle le eccezioni sui Throw
    '''     arrivati dal proxy locale per comodo all'interpretazione.
    ''' </summary>
    ''' <param name="funcname">Il nome della funzione che sta gestendo il try</param>
    ''' <param name="ex">L'eccezione che è arrivata dalla funzione di throw</param>
    Private Sub SetExceptionsStatus(funcname As String, ex As Exception)

        If ex.Message = GLB_ERROR_PARSING Then

            ' Se una funzione nel Flow mi ha dato picca
            ' ed ha sollevato questa eccezione di parsing
            ' sulla chiamat interna.
            m_LastStatus = GLB_ERROR_PARSING
            m_LastErrorMessage = "Errore di parsing sul protocollo Argentea (Chiamare assistenza)"

        ElseIf ex.Message = GLB_ERROR_FORM_INSTANCE Then

            ' Nell'instanziare il Form è successo qualche       *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            ' errore di valutazione sul tipo specifico.
            m_LastStatus = GLB_ERROR_FORM_INSTANCE
            m_LastErrorMessage = "Errore interno alla procedura --istance or object-- (Chiamare assistenza)"

        ElseIf ex.Message = GLB_ERROR_FORM_DATA Then

            ' Eventuali errori interni sull'iterazione con      *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            ' i controlli grigli e altro nel Form legato alla
            ' presentazione dei dati con errore non previsto
            m_LastStatus = GLB_ERROR_FORM_DATA
            m_LastErrorMessage = "Errore interno alla procedura --data values or stream-- (Chiamare assistenza)"

        Else

            ' Altro non previsto in questa funzione             *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
            m_LastErrorMessage = "Errore interno alla procedura --exception unexcpted-- (Chiamare assistenza)"
            ex = New Exception(GLB_ERROR_NOT_UNEXPECTED, ex)

        End If
        '
        If Not ex.InnerException Is Nothing Then
            LOG_Debug(getLocationString(funcname), m_LastErrorMessage + "--" + ex.InnerException.ToString())
        Else
            LOG_Debug(getLocationString(funcname), m_LastErrorMessage)
        End If
        '
        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
        '
    End Sub


#End Region

#Region "Functions Common and Argentea Specifique"

    ''' <summary>
    '''     Restituisce una stringa codificata ripresa dalla sessione 
    '''     Argentea secondo le sue specifiche di protocollo in input.
    ''' </summary>
    ''' <remarks>
    '''     Inizialization:
    '''     Dematerialization:
    '''             CouponCode:                 tag COUP_CODE                           -> code to identify coupon, tag ;
    '''             Codice Cassiere:            tag CouponCode BPSW2$[codicecassiere]   -> extra cash identifier that will be added As extra   data   In   the   form, . 
    '''             Progressive:                tag PROG                                -> progressive of cash counter operation;
    '''             IdCassa:                    tag CASSA                               -> RUPP as provided by Argentea, tag CASSA;
    '''             CodiceDevice:               tag COD_DEV                             -> numeric field. Usually it Is created With shop code plus cash counter code;
    '''                                                                                   (e.g if shop code Is 120 And cash counter Is 1 then you can pass 0012000001)
    '''             CodiceScontrino:            tag COD_SCON                            -> receipt number(numeric);
    '''             RFU_1                       tag RFU_2                               -> Internal Use
    '''             RFU_2                       tag RFU_2                               -> Internal Use
    '''     Reverse:
    '''             CouponCode:                 tag COUP_CODE                           -> code to identify coupon, tag ;
    '''             Progressive:                tag PROG                                -> progressive of cash counter operation;
    '''             IdCassa:                    tag CASSA                               -> RUPP as provided by Argentea, tag CASSA;
    '''             CodiceDevice:               tag COD_DEV                             -> numeric field. Usually it Is created With shop code plus cash counter code;
    '''                                                                                   (e.g if shop code Is 120 And cash counter Is 1 then you can pass 0012000001)
    '''             CodiceScontrino:            tag COD_SCON                            -> receipt number(numeric);
    '''             TransactionId               tag TRAN_ID     -> Transazione da parte Argentea
    '''             RFU_1                       tag RFU_2                               -> Internal Use
    '''             RFU_2                       tag RFU_2                               -> Internal Use
    '''     Confirm:
    '''            Progressive                  tag PROG        -> Progressivo di chiamata verso Argentea
    '''            IdCassa                      tag CASSA       -> Prefisso foarmato per Cliente-Cassa e non terminal
    '''            CouponCode                   tag COUP_CODE   -> Codice Coupon (Barcode)
    '''            TransactionId                tag TRAN_ID     -> Transazione da parte Argentea
    '''            RFU_1                        tag RFU_2       -> Internal Use
    '''            RFU_2                        tag RFU_2       -> Internal Use
    ''' </remarks>
    ''' <returns>String</returns>
    Public Function GetCodifiqueReceipt(TypeCodifique As TypeCodifiqueProtocol) As String
        Dim Result As String

        Result = ""
        m_RUPP = Me.GetPar_RUPP  ' Accertiamo il valore

        Select Case TypeCodifique
            Case TypeCodifiqueProtocol.Inizialization

                Result = "" ' Non Utilizzato

            Case TypeCodifiqueProtocol.Dematerialization

                Result = "COUP_CODE=" + m_CurrentBarcodeScan + "-BPSW2$=" + Me.Get_CodeOperatorID() + "-PROG=" _
                 + m_ProgressiveCall.ToString() + "-CASSA=" + m_RUPP + "-COD_DEV=" + Me.Get_CodeCashDevice +
                 "-COD_SCON=" + Me.Get_ReceiptNumber + "-RFU_1=-RFU_2=-"

            Case TypeCodifiqueProtocol.Reverse

                Result = "COUP_CODE=" + m_CurrentBarcodeScan + "-PROG=" _
                 + m_ProgressiveCall.ToString() + "-CASSA=" + m_RUPP + "-COD_DEV=" + Me.Get_CodeCashDevice +
                 "-COD_SCON=" + Me.Get_ReceiptNumber + "-TRAN_ID=" + m_CurrentTransactionID + "-RFU_1=-RFU_2=-"

            Case TypeCodifiqueProtocol.Confirm

                Result = "PROG=" + m_ProgressiveCall.ToString() + "-CASSA=" + m_RUPP + "-COUP_CODE=" + m_CurrentBarcodeScan + "-TRAN_ID=" + m_CurrentTransactionID

        End Select

        Return Result
    End Function


    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = Microsoft.VisualBasic.TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

#Region "Class DataResponse per il ResultData in risposta su Evento Collect"

    Private Interface IWResultDataList(Of T)
        ReadOnly Property Count As Integer
        Default Property Item(index As Integer) As T
        Sub Add(item As T)
        Sub Clear()
        Sub CopyTo(array() As T, arrayIndex As Integer)
        Sub Insert(index As Integer, item As T)
        Sub RemoveAt(index As Integer)
        Function Contains(item As T) As Boolean
        Function GetEnumerator() As IEnumerator(Of T)
        Function IndexOf(item As T) As Integer
        Function NewPaid(BarCode As String, TerminalId As String) As PaidEntry
        Function Remove(item As T) As Boolean
    End Interface

    ' Il writer per l'istanza interna corrente
    Private Shared WriterResultDataList As IWResultDataList(Of PaidEntry)


    ''' <summary>
    '''     Di supporto alla ClsPosArgentea conferisce
    '''     attributi dati in risposta alla comunicazione.
    ''' </summary>
    Public Class DataResponse

#Region "Membri Privati"

        ' All'ingresso serbo quelli pagati
        ' da sessioni precedenti
        Private _InitialBPPayed As Integer = 0                   ' <-- All'ingresso il conteggio dei BP già usati nelle sessioni di vendita precedenti
        Private _InitialTotalPayed As Decimal = 0                ' <-- All'ingresso il conteggio in valore già usato nelle sessioni di vendita precedenti
        Private _InitialTotalExcedeed As Decimal = 0             ' <-- All'ingresso il conteggio in valore già usato nelle sessioni di vendita precedenti

        '
        '   Data list del risultato dei 
        '   Barcode scansionati o dal tot<le
        '   che il servizio argenetea ha dato.
        '
        Private m_ListEntries As ResultDataList(Of PaidEntry)

#End Region

#Region ".ctor"

        ''' <summary>
        '''     .ctor
        ''' </summary>
        ''' <param name="TotCurrentBPUsed">Totale dei buoni usati nella sessione precedente eventuale</param>
        ''' <param name="TotValuePayedUsed">Totale in valore dei buoni usati per il pagamento nelle sessioni precedenti</param>
        ''' <param name="TotValueExcedeedUsed">Totale di un valore che è stato conteggiato in precedenza come resto</param>
        Public Sub New(ByVal TotCurrentBPUsed As Integer, ByVal TotValuePayedUsed As Decimal, ByVal TotValueExcedeedUsed As Decimal)

            ' Riserbo le iniziali quelli alla chiamata Totali che serviranno al consumer
            _InitialBPPayed = TotCurrentBPUsed
            _InitialTotalPayed = TotValuePayedUsed
            _InitialTotalExcedeed = TotValueExcedeedUsed

            ' Collection di risultati da riportare al consumer
            m_ListEntries = New ResultDataList(Of PaidEntry)()

            ' L'elemento a scrittura interna
            WriterResultDataList = m_ListEntries

        End Sub

#End Region

#Region "Properties"


        ''' <summary>
        '''     Tipo di BP elaborato in questa modalità Proxy
        ''' </summary>
        ''' <returns>Il tipo di PB elaborato <see cref="enTypeBP"/></returns>
        Protected Friend Overridable ReadOnly Property typeBPElaborated() As enTypeBP
            Get
                Return m_TypeBPElaborated_CS
            End Get
        End Property

        ''' <summary>
        '''     Numero totale di Buoni utilizzati dalla sessione
        '''     sul POS esterno per pagare il Totale sulla vendita
        '''     corrente.
        ''' </summary>
        ''' <returns>Numerico Integer</returns>
        Protected Friend Overridable ReadOnly Property totalBPUsed() As Integer
            Get
                Return m_TotalBPUsed_CS
            End Get
        End Property

        ''' <summary>
        '''     Il totale ottenuto dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sulla vendita corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalPayedWithBP() As Decimal
            Get
                Return m_TotalPayed_CS
            End Get
        End Property

        ''' <summary>
        '''     Il totale in eccesso su dei buoni transitati 
        '''     nella sessione del POS sulla vendita corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalExcedeedWithBP() As Decimal
            Get
                Return m_TotalValueExcedeed_CS
            End Get
        End Property

        ''' <summary>
        '''     Il set di risultati di tutti
        '''     i codici scansionati  dentro
        '''     il servizio. (per l'istanza a sola lettura esterna)
        ''' </summary>
        ''' <returns>Collection con tutti le scansioni che partecipano al pagamento in corso.</returns>
        Public ReadOnly Property PaidEntryBindingSource As IResultDataList(Of PaidEntry)
            Get
                Return m_ListEntries
            End Get
        End Property

        ''' <summary>
        '''     Restituisce se un elemento nell'nsieme ha un
        '''     determinato Barcode già inserito (che non può essere solitamente in questo contesto)
        ''' </summary>
        ''' <param name="m_CurrentBarcodeScan">Un EAN usato come BP per partecipare al pagamento</param>
        ''' <returns></returns>
        Public Function ContainsBarcode(m_CurrentBarcodeScan As String) As Boolean
            For Each itm As PaidEntry In m_ListEntries
                If itm.Barcode = m_CurrentBarcodeScan.Trim() Then
                    Return True
                End If
            Next
            Return False
        End Function


#End Region

#Region "IEnumerator ResultData"

        ''' <summary>
        '''     Collection a sola lettura
        '''     per il set di risultati.
        ''' </summary>
        ''' <typeparam name="T">Il Tipo di dato dentro la collection</typeparam>
        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Advanced)>
        Public Interface IResultDataList(Of T)
            ReadOnly Property Count As Integer
            Function Contains(item As T) As Boolean
            Function GetEnumerator() As IEnumerator(Of T)
            Function IndexOf(item As T) As Integer
        End Interface

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Protected Class ResultDataList(Of T)
            Implements IList(Of T)
            Implements IResultDataList(Of T)
            Implements IWResultDataList(Of T)

            Private _Count As Integer
            Private _Capacity As Integer = 64
            Private _BackingStore(0 To _Capacity - 1) As T

            Private Sub ResizeIfNeeded()
                If _Capacity = _Count Then
                    _Capacity *= 2
                    ReDim Preserve _BackingStore(0 To _Capacity - 1)
                End If
            End Sub

            Public Sub Add(ByVal item As T) Implements System.Collections.Generic.ICollection(Of T).Add, IWResultDataList(Of T).Add
                ResizeIfNeeded()

                _BackingStore(_Count) = item
                _Count += 1
            End Sub

            ''' <summary>
            '''     Aggiunge un elemento all'insieme
            '''     e restituisce il riferimento.
            ''' </summary>
            ''' <param name="item"></param>
            ''' <returns><see cref="T"/>L'elemento appena aggiunto</returns>
            Public Overridable Function NewPaid(ByVal BarCode As String, ByVal TerminalId As String) As PaidEntry Implements IWResultDataList(Of T).NewPaid
                ' Metodo valido solo per i tipi PaidEntry
                ' si ptrebbe inseire a default(T) ma non ho tempo cassarola.
                Dim NewElement As Object = New PaidEntry(BarCode, TerminalId)
                Me.Add(NewElement)
                Return NewElement
            End Function


            Public Sub Clear() Implements System.Collections.Generic.ICollection(Of T).Clear, IWResultDataList(Of T).Clear
                _Count = 0
            End Sub

            Public Function Contains(ByVal item As T) As Boolean Implements System.Collections.Generic.ICollection(Of T).Contains, IResultDataList(Of T).Contains, IWResultDataList(Of T).Contains
                Return IndexOf(item) <> -1
            End Function

            Public Sub CopyTo(ByVal array() As T, ByVal arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of T).CopyTo, IWResultDataList(Of T).CopyTo
                If array Is Nothing Then
                    Throw New ArgumentNullException("array")
                End If
                If arrayIndex + Me.Count > array.Length Then
                    Throw New ArgumentException("array")
                End If
                For i As Integer = 0 To Me.Count - 1
                    array(arrayIndex + i) = Me(i)
                Next i
            End Sub

            Public ReadOnly Property Count() As Integer Implements System.Collections.Generic.ICollection(Of T).Count, IResultDataList(Of T).Count, IWResultDataList(Of T).Count
                Get
                    Return _Count
                End Get
            End Property

            Public ReadOnly Property IsReadOnly() As Boolean Implements System.Collections.Generic.ICollection(Of T).IsReadOnly
                Get
                    Return False
                End Get
            End Property

            Public Function Remove(ByVal item As T) As Boolean Implements System.Collections.Generic.ICollection(Of T).Remove, IWResultDataList(Of T).Remove
                Dim index As Integer = Me.IndexOf(item)
                If index = -1 Then
                    Return False
                Else
                    RemoveAt(index)
                    Return True
                End If
            End Function

            Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of T) Implements System.Collections.Generic.IEnumerable(Of T).GetEnumerator, IResultDataList(Of T).GetEnumerator, IWResultDataList(Of T).GetEnumerator
                Return New MyListEnumerator(Me)
            End Function

            Public Function IndexOf(ByVal item As T) As Integer Implements System.Collections.Generic.IList(Of T).IndexOf, IResultDataList(Of T).IndexOf, IWResultDataList(Of T).IndexOf
                Return Array.IndexOf(_BackingStore, item)
            End Function

            Public Sub Insert(ByVal index As Integer, ByVal item As T) Implements System.Collections.Generic.IList(Of T).Insert, IWResultDataList(Of T).Insert
                ResizeIfNeeded()

                _Count += 1

                For i As Integer = Me.Count - 1 To index + 1 Step -1
                    Me(i) = Me(i - 1)
                Next i

                Me(index) = item
            End Sub

            Default Public Property Item(ByVal index As Integer) As T Implements System.Collections.Generic.IList(Of T).Item, IWResultDataList(Of T).Item
                Get
                    If index >= Me.Count Then
                        Throw New ArgumentOutOfRangeException("index")
                    End If
                    Return _BackingStore(index)
                End Get
                Set(ByVal value As T)
                    If index >= Me.Count Then
                        Throw New ArgumentOutOfRangeException("index")
                    End If
                    _BackingStore(index) = value
                End Set
            End Property

            Public Sub RemoveAt(ByVal index As Integer) Implements System.Collections.Generic.IList(Of T).RemoveAt, IWResultDataList(Of T).RemoveAt
                If index >= Me.Count Then
                    Throw New ArgumentOutOfRangeException("index")
                End If

                For i As Integer = index To Me.Count - 2
                    Me(i) = Me(i + 1)
                Next i

                _Count -= 1
            End Sub

            Private Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
                Return New MyListEnumerator(Me)
            End Function

            Private Class MyListEnumerator
                Implements IEnumerator(Of T)

                Private _Parent As ResultDataList(Of T)
                Private _Index As Integer = -1

                Friend Sub New(ByVal parent As ResultDataList(Of T))
                    _Parent = parent
                End Sub

                Public ReadOnly Property Current() As T Implements System.Collections.Generic.IEnumerator(Of T).Current
                    Get
                        If _Index = -1 Then
                            Throw New InvalidOperationException()
                        End If

                        Return _Parent(_Index)
                    End Get
                End Property

                Private ReadOnly Property Current1() As Object Implements System.Collections.IEnumerator.Current
                    Get
                        Return Me.Current
                    End Get
                End Property

                Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
                    If _Index < _Parent.Count - 1 Then
                        _Index += 1
                        Return True
                    Else
                        Return False
                    End If
                End Function

                Public Sub Reset() Implements System.Collections.IEnumerator.Reset
                    _Index = -1
                End Sub

                Public Sub Dispose() Implements IDisposable.Dispose
                    'Not required.
                End Sub

            End Class
        End Class

#End Region

    End Class

#End Region

End Class