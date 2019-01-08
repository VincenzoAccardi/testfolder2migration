Imports System
Imports ARGLIB = PAGAMENTOLib
Imports System.Collections
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos.EFT
Imports TPDotnet.Pos
Imports System.Windows.Forms

' 0=PRODUZIONE 1=TEST
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

    ''' <summary>
    '''     Collection usata per  le Transazioni
    '''     Argentea andati a buon  fine dopo la
    '''     chiamata alla funzione Argentea o da 
    '''     POS hardware
    ''' </summary>
    'Private _listBpCompletated As New Collections.Generic.Dictionary(Of String, BPType)(System.StringComparer.InvariantCultureIgnoreCase)
    Private _DataResponse As DataResponse

    '
    ' Per il servizio che usa un Form interno
    ' per l'inserimento manuale dei Buoni Pasto
    ' mi appoggio su un Form della cassa già presente.  ( Specializzato a FormEmulationArgentea )
    '
    Private frmEmulation As FormEmulationArgentea = Nothing           ' <-- Il Form di appoggio per servire il POS software sulla cassa corrente

#Region "DBParameters"

    'lRetailStoreID	 szWorkstationGroupID	szObject	szDllName	szKey	                                    szContents	lTechLayerAccessID	szLastUpdLocal
    ' 599            RITSSWFP	            Argentea	StPosMod	BP_ACCEPT_EXCEDEED_VALUES	                N	        1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BP_ACCORPATE_ON_TA	                        Y	        1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BP_MAX_BP_PAYABLES_SOME_SESSION	            7	        1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BP_RUPP_ARGENTEA	                        A1939338	1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BP_VIEW_FORM_FOR_MESSAGES_STATUS_POS	    N	        1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BP_VIEWS_MODE_ELEMENT_DELETETED	            0	        (null)	            (null)
    ' 599            RITSSWFP	            Argentea	StPosMod	BPE_VIEW_MESSAGES_ON_RETURN_FROM_POS	    1	        1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BPE_AUTOCLOSE_ON_COMPLETE_OPERATION	        Y	        (null)	            (null)
    ' 599            RITSSWFP	            Argentea	StPosMod	BPE_COPIES	                                1	        1	                20171018115327
    ' 599            RITSSWFP	            Argentea	StPosMod	BPE_PRINT_WITHIN_TA	                        Y	        1	                20171018115327

#End Region

#Region "CONST di INFO e ERRORE Private"

    ' Messaggeria per codifica segnalazioni ID di errore remoti 
    Private msgUtil As New TPDotnet.IT.Common.Pos.Common

    ' *** **** ****

    ' Notifiche

    ' Su errore quando l'opzione di girare il resto in eccesso
    ' assegno la costante di operazione non valida ai fine di pagamento.
    Private Const NOT_OPT_ERROR_VALUE_EXCEDEED As String = "Error-OPTION-PAYABLE-WITH-REST"

    ' Numero dei titoli di pagamento usati se
    ' il limite è impostato superato per la sessione.
    Private Const NOT_OPT_ERROR_NUMBP_EXCEDEED As String = "Error-OPTION-PAYABLE-NUMBP-EXCEDEED"

    ' BarCode già utilizzato in precedenza evitiamo
    ' di richiamare argentea per il controllo
    Private Const NOT_INFO_CODE_ALREADYINUSE As String = "Error-BARCODE-ALREADYINUSE"

    ' Quando l'importo delle righe è già stato raggiunto (con o 
    ' senza eccesso per eventuale ipotesi di resto) non procediamo.
    Private Const NOT_INFO_IMPORT_ALREADYCOMPLETED As String = "Error-IMPORT-ALREADY_COMPLETED"

    ' BarCode da rimuovere da quelli già scanditi in precedenza 
    ' non presente in elenco
    Private Const NOT_INFO_CODE_NOTPRESENT As String = "Error-BARCODE-NOTPRESENT"

    ' Su errori non bloccanti ma da segnalare all'operatore
    ' come stmapa scontrino non effettuata o altro usiamo questo.
    Private Const NOT_ERROR_PRINTER_FAILED As String = "Error-PRINTER_FAIL"

    ' Su errori non bloccanti ma da segnalare all'operatore
    ' come quando magari l'hardware non è collegato al sistema.
    Private Const NOT_ERROR_HARDWAREPOS_FAILED As String = "Error-HARDWAREPOS_FAIL"

    ' Su errori non bloccanti ma da segnalare all'operatore
    ' come quando magari l'emulatore software ha fallito nell'avvio.
    Private Const NOT_ERROR_SOFTWAREPOS_FAILED As String = "Error-OFTWAREPOS_FAIL"

    ' In fase di chiamata al servizio remoto di Argentea per la
    ' chiamata API Demtarelializzazione mi restituire NOT VALID
    ' questo è speciale perchè nel momento del setstatus riporta
    ' per la msgbox una codificazione dello status interno + quello remoto
    Private Const NOT_INFO_OPERATION_NOT_VALID_SPECIAL As String = "Error-SIGNAL_REMOTE_FAIL"

    ' Per gli errori interni di exception non classificati o classificati
    ' per la msgbox questa codifica lo riprende come errore interno
    Private Const NOT_INFO_ERROR_INTERNAL As String = "Error-SIGNAL_SOFTWARE_INTERNAL"

    ' Operatore

    ' Su chiamate a Nomi di Funzioni API remote
    ' che non sono gestite da questo proxy
    Private Const OPR_ERROR_API_NOT_VALID As String = "Error-API_NOT_VALID"

    ' Form POS

    ' Inizializzazione dell'emulatore POS
    Private Const NOT_INFO_POS_INIT As String = "Info-POS_INIT"

    ' Chiamata remota dell'emulatore POS
    Private Const NOT_INFO_POS_CALL As String = "Info-POS_CALL"

    ' Errore nella funzione per l'emulatore POS
    Private Const NOT_INFO_POS_ERROR As String = "Info-POS_ERROR"

    ' Al momento che l'utente ha consluso
    Private Const NOT_INFO_POS_DATA_OK As String = "Info-POS_DATA_OK"

    ' Al momento del collect
    Private Const NOT_INFO_POS_CLOSING As String = "Info-POS_CLOSING"

    ' Al momento del print ticket al close
    Private Const NOT_INFO_POS_PRINTING As String = "Info-POS_PRINT"

    ' Al momento del print ticket al close in caso di errore
    Private Const NOT_INFO_POS_PRINTERR As String = "Info-POS_PRINT_ERR"

    ' ** Eccezioni

    ' Se durante la comunicazione con il dispositivo POS
    ' collegato alla cassa arriva un errore di proto o signal
    Private Const GLB_FAILED_POS_HARDWARE As String = "Error-FAILED_HARDWARE"

    ' Se un tentativo di inizializzazione e uno di riallinemento
    ' e ritentativo di inzializzazione ha fallito abbiamo FALLITO punto e basta.
    Private Const GLB_FAILED_INITIALIZATION As String = "Error-FAILED_INIATIALIZATION"

    ' In fase di inizializzazione la risposta di errore
    ' remota da parte di Argentea riportata nei log.
    Private Const GLB_FAILED_RESETCOUNTER As String = "Error-FAILED_RESETCOUNTER"

    ' Se un tentativo di conferma dopo richiesta di conferma a demat o void
    ' e la risposta remota era in KO per qualche motivo.
    Private Const GLB_FAILED_CONFIRMATION As String = "Error-FAILED_CONFIRMATION"

    ' Se un tentativo di storno su materializzazione per un titolo
    ' e la risposta remota era in KO per qualche motivo.
    Private Const GLB_FAILED_VOIDDEMATERIALIZATION As String = "Error-FAILED_VOID"

    ' Se un tentativo di dematerializzazione di un titolo
    ' e la risposta remota era in  KO  per qualche motivo.
    Private Const GLB_FAILED_DEMATERIALIZATION As String = "Error-FAILED_DEMAT"

    ' La chiamata all'api close può aver dato esito negativo
    ' dato che avrebbe dovuto chiudere la sessione.
    Private Const GLB_OPT_ERROR_ON_API_CLOSE As String = "Error-API_CLOSE"

    ' ** Enumerazioni

    ''' <summary>
    '''     Stati interni per la Risposta 
    '''     tra funzioni nel triplo stato.
    ''' </summary>
    Private Enum StatusCode

        ''' <summary>
        '''     OK
        ''' </summary>
        OK = True

        ''' <summary>
        '''     KO
        ''' </summary>
        KO = False

        ''' <summary>
        '''     In risposta di OK
        '''     ma richiede conferma.
        ''' </summary>
        CONFIRMREQUEST = 9

        ' *-*-*-- Speciali Per resetcounter con gestione a chiamata ricorsiva

        ''' <summary>
        '''     Non è stato necessario fare  il  Reset
        '''     del contatore remoto può procedere con
        '''     la funzione in corso.
        ''' </summary>
        RESETCOUNTER_CONTINUE = 0

        ''' <summary>
        '''     E' stato effettuato il reset del contatore
        '''     remoto e l'esito è  riuscito  quindi  puoi 
        '''     procedere con la funzione in corso.
        ''' </summary>
        RESETCOUNTER_OK = 1

        ''' <summary>
        '''     E' stato effettuato il reset del contatore
        '''     remoto e l'esito non è riuscito quindi non 
        '''     si può procedere con la funzione in  corso
        '''     e segnaliamo l'utente con l'errore (Reset non Riuscito)
        ''' </summary>
        RESETCOUNTER_KO = 2

        ''' <summary>
        '''     Speciale, e' stato effettuato il Reset  e
        '''     prima di completare lo status la funzione
        '''     chiamante richiamerà ricorsivamente se stessa.
        ''' </summary>
        RESETCOUNTER_REQUEST_TO_RECALL = -1

        ''' <summary>
        '''     Speciale, nell'analisi della risposta dopo
        '''     le operazioni per restituire al  chiamante
        '''     lo status di m_LastStatusResult.Successfully
        ''' </summary>
        RESETCOUNTER_RETURN_TO_CALL = -2

    End Enum

    ''' <summary>
    '''     Nella scelta del giusto tipo di Parser
    '''     facciamo in modo che l'argomento identifichi
    '''     la corretta azione API
    ''' </summary>
    Public Enum enApiToCall

        ''' <summary>
        '''     Non stiamo chiamando nessuna azione api per il parser corretto
        '''     Nel parser verrà presa in considerazione la command dellla dll argentea
        ''' </summary>
        None

        ''' <summary>
        '''     API per la corretta chiamata al servizio prima della Command di demat della dll inizializzando il service remoto
        ''' </summary>
        Initialization

        ''' <summary>
        '''     API di serfvizio Argentea remoto per il reset del contatore di chiamate remoto
        ''' </summary>
        ResetCounter

        ''' <summary>
        '''     API di servizio quando alcuni tipi di dematerializzazione richiedono la conferma dopo il demat stesso
        ''' </summary>
        Confirmation

        ''' <summary>
        '''     API di servizio per richiedere informazioni intorno al codice di un BP o Coupon (Validità e Imposrto Facciale)
        ''' </summary>
        CheckBP

        ''' <summary>
        '''     API di servizio speciale da richiedere per il Payment di servizio (singolo pagamento alla volta)
        ''' </summary>
        SinglePayment

        ''' <summary>
        '''     API di servizio speciale da richiedere ad un terminale l'insieme dei Payment di servizio (più pagamento effettuati sul pos)
        ''' </summary>
        MultiplePayments

        ''' <summary>
        '''     API di servizio speciale da richiedere per il Void di servizio (singolo storno alla volta)
        ''' </summary>
        SingleVoid

        ''' <summary>
        '''     API di servizio speciale da richiedere ad un terminale l'insieme degli Storni di servizio (più storni effettuati sul pos)
        ''' </summary>
        MultipleVoids

        ''' <summary>
        '''     API di servizio speciale da richiedere ad un terminale per azionare l'utente ad inserire la carta ed avere le info relative al suo ammontare
        ''' </summary>
        InfoCardUser

        ''' <summary>
        '''     API per la corretta chiamata al servizio alla fine della Command di demat della dll concludendo sul service remoto la convalida
        ''' </summary>
        Close

    End Enum

    '
    ' Elementi necessari per il parsing
    ' sul protocollo previsto e di frazione
    ' per il risultato.
    '
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

    ''' <summary>
    '''     Comportamento ed Azione 
    '''     del Proxy in emulazione.
    ''' </summary>
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

    ''' <summary>
    '''     Globalmente alla sessione in corso evita di riaprire
    '''     nuovamente il Ticket di Sessione per la volta successiva.
    ''' </summary>
    Public Shared FLAG_STATIC_INITIALIZATED As Boolean

#End Region

#Region "Membri Privati"

    '
    ' Comportamento a secondo del proxy di utilizzo
    ' e status logico interno rispetto all'inizio e fine di vita.
    '
    Private m_TypeProxy As enTypeProxy                                          ' <-- Il Comportamento che deve assumere il Procy corrente
    Private m_CommandToCall As enCommandToCall = enCommandToCall.Payment        ' <-- Il comando che deve eseguire nella modalità in esecuzione (per default in pagamento)
    Private m_CurrentTransactionID As String = String.Empty                     ' <-- La Transazione in corso sulla TA GUID 
    Private m_CurrentAmountScalable As Decimal = 0                              ' <-- L'importo pagato/stornato in ingresso sul servizio da scalare
    Private m_TotalBPInUseOnCurDoc As Integer = 0                               ' <-- Il Numero di Titoli prima dell'istanza usati fino ad adesso sull'intero documento
    Private m_InitialPaymentsInTA As Decimal = 0                                ' <-- Il Totale sulla TA prima del void o del payment
    Private m_ServiceStatus As enProxyStatus = enProxyStatus.Uninitializated    ' <-- Lo stato iniziale ed in corso del flow del Proxy corrente
    Private m_SilentMode As Boolean = False                                     ' <-- Se mostrare all'utente i messaggi di errore e di avviso

    '
    ' Passati dal Chiamante per essere
    ' letti agiornati in uscita.
    '
    Private m_PaidAmount As Decimal                             ' <-- Il Pagato fino ad adesso all'entrata
    Private m_PayableAmount As Decimal                          ' <-- Il pagabile con le azioni del servizio
    '
    Private m_VoidAmount As Decimal                             ' <-- Lo Storno attuale fino ad adesso all'entrata
    Private m_VoidableAmount As Decimal                         ' <-- Lo Stornabile o lo stornato con le azioni del servizio
    '
    Private m_PrefillVoidable As Dictionary(Of String, PaidEntry) ' <-- Gestisce un possibile elenco di BP prefillato sul FORM di appoggio per gestire storni tramite operatore

    '
    '   Per le chiamate API dirette
    '
    Private m_CurrentApiNameToCall As String                    ' <-- Riporta l'ultima chiamata API tramite dll COM di argentea chiamata

    '
    ' Aggiornati per il Risultato
    '
    Shared m_Transaction_Identifier As String                   ' <-- In Comunicazione con il Dispostivo Hardware identifica l'id della transazione sulla Carta BP
    '
    Shared m_TypeBPElaborated_CS As enTypeBP                    ' <-- Il Tipo di BP o gruppo di BP elaborati nella sessione
    Shared m_TotalBPElaborated_CS As Integer                    ' <-- Il Numero totale di BP elaborati nella sessione sia in Demat che in Void
    '
    Shared m_TotalBPUsedToPay_CS As Integer                     ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento validi
    Shared m_TotalPayed_CS As Decimal                           ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente per il pagamento
    Shared m_TotalValueExcedeed_CS As Decimal                   ' <-- Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata
    '
    Shared m_TotalBPUsedToVoid_CS As Integer
    Shared m_TotalVoided_CS As Decimal                          ' <-- L'Accumulutore Globale al Proxy corrente nella sessione corrente per lo storno
    Shared m_TotalValueExtraVoidNotContabilizated As Decimal    ' <-- Situazione in cui si stanno dando più storni rispetto a quelli previsti
    '
    Shared m_TotalBPNotValid_CS As Integer
    Shared m_TotalInvalid_CS As Decimal                         ' <-- L'Accumulutore Globale al Proxy corrente nella sessione corrente per lo storno risultanti non validi

    ' Parziali durante il conteggio

    Private m_PartialBPUsedToPay_XS As Integer                  ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento validi
    Private m_PartialPayed_XS As Decimal                        ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente per il pagamento
    Private m_PartialValueExcedeed_XS As Decimal                ' <-- Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata
    '
    Private m_PartialBPUsedToVoid_XS As Integer
    Private m_PartialVoided_XS As Decimal                       ' <-- L'Accumulutore Globale al Proxy corrente nella sessione corrente per lo storno
    '
    Private m_PartialBPNotValid_XS As Integer
    Private m_PartialInvalid_XS As Decimal                      ' <-- L'Accumulutore Globale al Proxy corrente nella sessione corrente per lo storno risultanti non validi

    ' Conteggio Parziale di Funzione
    Dim _PartialTransactValue As Decimal = 0

    '
    ' Variabili private
    '
    Private m_LastStatus As String                                      ' <-- Ultimo Status di Costante per errore in STDOUT
    Private m_LastErrorMessage As String                                ' <-- Ultimo Messaggio di errore STDOUT
    Private m_LastResponseRawArgentea As ArgenteaFunctionReturnObject   ' <-- Ultima risposta di Argentea per STDOUT (di utilità al reprint dello scontrino)
    Private m_LastCrcTransactionID As String = String.Empty             ' <-- Il Crc di risposta in risposta arrivato da Argentea

    '
    ' Status interni e ultime letture
    '
    Private m_FirstCall As Boolean = False                  ' <-- Inizializzazione alla prima chiamata dal Form di scansione per i Barcode vs Argentea
    Private m_CurrentBarcodeScan As String = String.Empty   ' <-- Ultimo Barcode scansionato
    Private m_CurrentValueOfBP As Decimal                   ' <-- Valore facciale dell'n Barcode di BP scansionato

    'Private m_CurrentTerminalID As String = Nothing        ' <-- In Pos Hardware identifica il POS usato in Software l'ID del WebService

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
    '   Variabili private per la gestione interna
    '   in emulazione software del servizio Argentea
    '
    Public Shared m_ProgressiveCall As Integer = 1          '   <--     Il progressive call Privato (Relativo a tutte le chiamate in sequenza richieste dal protocollo)"
    Private m_RUPP As String = Nothing                      '   <--     Il RUPP necessario per comunicare con il servizio Remoto

    ''' :: Flag per il tentativo di Reset del Counter remoto presso Argentea
    Private _flagCallOnetimeResetIncrement As Boolean = False
    ''' :: Log durante delle chiamate di ogni errore o stato precedente
    Private m_LogErrors As Dictionary(Of Integer, tLogErr)

    '
    ' Parametri Globali di applicazione predefiniti
    ' nel contesto backStore per Argentea.
    '
    Private st_Parameters_Argentea As ArgenteaParameters    ' <-- Riprende dal modello statico tutti i parametri globali nel contesto corrente dedicati ad Argentea

    '
    ' I Primi due parametri per la visualizzazione
    ' o meno del WaitScreen e quello in caso di una
    ' modalità Hardware POS se mostrare i msgbox
    ' degli stati di errore che arrivano dal POS hardware
    '
    Private m_OPT_ShowWaitScreenLevel As Integer    '   ->  st_Parameters_Argentea.BP_ShowWaitScreenLevel
    Private m_OPT_ShowMessageBoxLevel As Integer    '   ->  st_Parameters_Argentea.BP_ShowMessageBoxLevel

    '
    ' Il Numero massimo di BP da  poter  usare
    ' come pagamento nella stessa sessione, ma
    ' corretto in entrata con il Numero Massimo
    ' espresso nella TA ripreso dalle altre sessioni
    '
    Private m_OPT_MaxNumBPSomeSession As Integer    '   ->  st_Parameters_Argentea.BP_MaxBPPayableSomeSession


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
            'If m_RUPP <> Nothing And m_RUPP <> "" Then m_RUPP = st_Parameters_Argentea.BPRupp
            If st_Parameters_Argentea.BPRupp.Trim = "" Then

                ' Sollevo eccezione
                Throw New ExceptionProxyArgentea("GetPar", ExceptionProxyArgentea.LOC_PAR_NOT_CONFIGURATED, "Parametro necessario per il Proxy mancante -- Parametro RUPP --")

            End If
            Return st_Parameters_Argentea.BPRupp
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
    ''' <param name="TotalPayableInSession">Il Totale pagato fino adesso prima di effettuare l'eggiornamento dai dati del proxy corrente.</param>
    ''' <param name="TotalBPInUseToDoc">Il Numero totale di Titoli BP usati fino adesso sul documento (Valido per eventuale blocco su un numero predefinitio di Titoli utilizzabili per il Doc)</param>
    Protected Friend Sub New(
                             ByRef theModCntr As ModCntr,
                             ByRef taobj As TA,
                             TypeBehavior As enTypeProxy,
                             ByVal CurrentTransactionID As String,
                             ByVal TotalPayableInSession As Decimal,
                             Optional TotalBPInUseToDoc As Integer = 0,
                             Optional CurrentTotalInTA As Decimal = Decimal.MinValue
                             )

        Dim funcName As String = "ClsProxyArgentea.New"

        '
        ' Gli oggetti di base
        '
        m_taobj = taobj
        m_TheModcntr = theModCntr

        ' Tipo BEHAVIOR
        m_TypeProxy = TypeBehavior

        ' Se non è stato Passato il Totale della
        ' TA come rimanenza di pagamento  prende
        ' comunque a riferimento l'importo di sessione
        ' utilizzato come importo a scalare.
        If CurrentTotalInTA = Decimal.MinValue Then
            m_InitialPaymentsInTA = TotalPayableInSession   ' O prende il Totale pagabile Come importo totale di TA
        Else
            m_InitialPaymentsInTA = CurrentTotalInTA        ' O prende il Totale vero della TA corrente
        End If

        ' Dati fondamentali
        m_CurrentTransactionID = CurrentTransactionID       ' l'ID della Transazione GUID sulla TA
        m_CurrentAmountScalable = TotalPayableInSession     ' l'amount da usare come da pagare o da stornare quindi lo scalabile per questa azione
        m_TotalBPInUseOnCurDoc = TotalBPInUseToDoc          ' Il Numero di buoni presenti sull'intero documento fino ad adesso
        m_LogErrors = New Dictionary(Of Integer, tLogErr)   ' Log degli stati di volta in volta sul corso della sessione

        ' Preparo il DataResponse e la Vista sul Form
        _UpdateResultData("INITIALIZE", m_InitialPaymentsInTA, Nothing)

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

            ' I Primi due parametri per la visualizzazione
            ' o meno del WaitScreen e quello in caso di una
            ' modalità Hardware POS se mostrare i msgbox
            ' degli stati di errore che arrivano dal POS hardware
            m_OPT_ShowWaitScreenLevel = st_Parameters_Argentea.BP_ShowWaitScreenLevel
            m_OPT_ShowMessageBoxLevel = st_Parameters_Argentea.BP_ShowMessageBoxLevel

            '
            ' Riprendo dai parametri il numero massimo di
            ' titoli utilizzabili nella stessa sessione, ma
            ' sarà considerato insieme e sommato al numero
            ' di elementi presenti nelle altre sesioni al
            ' momento che verrà utilizzato.
            '
            m_OPT_MaxNumBPSomeSession = st_Parameters_Argentea.BP_MaxBPPayableSomeSession


        Catch ex As Exception

            ' Signal (come errore di stampa ma non bloccante)
            _SetOperationStatus(funcName,
                ExceptionProxyArgentea.LOC_ERROR_INSTANCE_PARAMETERS,
                "Non è stato possibile caricare i parametri applicativi per eseguire il servizio Argentea",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Bloccante
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_INSTANCE_PARAMETERS, "Errore nel caricare i parametri applicativi per eseguire il servizio Argentea -- Parametri Argentea non presenti --")

        Finally

        End Try

    End Sub

    ''' <summary>
    '''     In presenza di situazioni di Storno
    '''     può essere presente un  Prefill  di
    '''     tutti i BP facenti parte della TA o
    '''     di un elemento raggruppato come Media.
    '''     ( DATI IN ENTRATA SULLA SESSIONE )
    ''' </summary>
    Private Sub CalculateAndFillMultiItemsInitialize()
        Dim funcName As String = "CalculateAndFillMultiItemsInitialize"

        If Not m_PrefillVoidable Is Nothing Then
            Dim ItemNew As PaidEntry

            frmEmulation.Show()

            ' Faciamo questo test perchè se dovesse
            ' nell'stanza essere statos usato un importo
            ' da scalare diverso rispetto a quello che
            ' stiamo calcolando con il prefill passato
            ' allora vuol dire che c'è stato qualche
            ' errore di interpretazione della voce da stornare.
            Dim Old_AmountScalable As Decimal = Math.Abs(m_CurrentAmountScalable)
            m_CurrentAmountScalable = 0

            ' Il Prefill Riassegna i 
            ' valori iniziali di entrata
            m_TotalVoided_CS = 0
            m_VoidableAmount = 0
            m_TotalInvalid_CS = 0

            ' Aggiorno il DataResponse e la vista (Iniziale)
            _UpdateResultData("PREFILL_INIT", Old_AmountScalable, Nothing)

            For Each itm As KeyValuePair(Of String, PaidEntry) In m_PrefillVoidable

                '
                ' Questo è un Elemento che va visualizzato
                ' sul form e immesso nel dataResult per 
                ' una corretta gestione in uscita.
                '
                ItemNew = itm.Value

                ' Riformatto il Valore Stringato in modo corretto
                Dim CValueDecimal As Decimal = 0
                Dim FValueDecimal As Decimal = 0
                If m_TypeProxy = enTypeProxy.Service Then
                    CValueDecimal = ItemNew.DecimalValue / FractParsing
                    FValueDecimal = ItemNew.DecimalFaceValue / FractParsing
                ElseIf m_TypeProxy = enTypeProxy.Pos Then
                    CValueDecimal = ItemNew.DecimalValue
                    FValueDecimal = ItemNew.DecimalFaceValue
                End If
                ItemNew.Value = CValueDecimal.ToString("###,##0.00")
                ItemNew.FaceValue = FValueDecimal.ToString("###,##0.00")

                ' Aggiornamento dello stornabile e del pagato e del non contabilizzato nelle sessioni precedenti
                If Not ItemNew.Voided And Not ItemNew.Invalid Then

                    ' BP Usati come pagamento
                    m_PartialPayed_XS += CValueDecimal
                    m_PartialBPUsedToPay_XS += 1
                    m_VoidableAmount += CValueDecimal           ' Usato Come stornabile in modalità VOID
                    _PartialTransactValue -= CValueDecimal
                    m_CurrentAmountScalable += CValueDecimal

                ElseIf Not ItemNew.Invalid Then

                    ' BP Usati come storno
                    m_PartialVoided_XS += CValueDecimal
                    m_PartialBPUsedToVoid_XS += 1
                    m_VoidableAmount += CValueDecimal           ' Usato Come stornabile in modalità VOID
                    _PartialTransactValue += CValueDecimal
                    'm_CurrentAmountScalable -= CValueDecimal
                Else

                    ' BP Non Validi e non Contabilizzati
                    m_PartialBPNotValid_XS += 1
                    m_PartialInvalid_XS += CValueDecimal
                    '_PartialTransactValue -= CValueDecimal

                End If

                If m_CurrentAmountScalable < 0 Then
                    m_PartialValueExcedeed_XS = Math.Abs(m_CurrentAmountScalable)
                End If

                ' Aggiungo al dataResult per il calcolo in
                ' uscita da usare per aggiornare nella TA i MetaData
                WriterResultDataList.Add(ItemNew)

                ' Aggiorno il DataResponse e la vista
                _UpdateResultData("PREFILL_UPDATE", _PartialTransactValue, ItemNew) ' _PartialTransactValue

            Next

            ' Questo solo Per il POS Software (Che usa il contatore un elemento alla volta)
            If m_TypeProxy = enTypeProxy.Service Then
                _PartialTransactValue = m_CurrentAmountScalable
            Else
                m_VoidableAmount = m_CurrentAmountScalable
            End If

            If Old_AmountScalable <> Math.Abs(_PartialTransactValue) Then

                ' Sollevo l'eccezione (Prefill non congruente con il Totale passato)
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_PROXY_ALREADY_RUNNING, "Errore azione Proxy -- Totale storno non valido rispetto al Prefill in corso --")

            End If

            ' Completo il DataResponse e la vista ( Nella Vista Inserisco sempre l'amount completo con tutto gli storni )
            _UpdateResultData("PREFILL_END", m_VoidableAmount, Nothing) ' _PartialTransactValue

        End If

    End Sub

    ''' <summary>
    '''     Refilla dopo la chiamata Proxy
    '''     i Dati per la Risposta ed il Form
    '''     ( FASE INIZIALE IN DMAT )
    ''' </summary>
    ''' <param name="RefTo_Transaction_Identifier"></param>
    Private Sub CalculateAndFillMultiItemsInPayment(RefTo_Transaction_Identifier As String)
        Dim funcName As String = "CalculateAndFillMultiItemsInPayment"

        ' Status sul Form emulazione POS Hardware
        frmEmulation.SetStatus(PictureMultiStatusControlExpanse.enStatustype.Ok)
        '
        ' A differenza del Software  Creo  voci
        ' di TA quanti sono stati inoltrati nel
        ' dispositivo.
        '
        m_TotalBPUsedToPay_CS = m_LastResponseRawArgentea.NumBPEvalutated                           ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento
        If m_LastResponseRawArgentea.Amount > m_CurrentAmountScalable Then                          ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente
            m_TotalValueExcedeed_CS = m_LastResponseRawArgentea.Amount - m_CurrentAmountScalable    ' <-- ?? TODO:: Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata
        Else
            m_TotalValueExcedeed_CS = 0
        End If

        ' Il Pagabile da questo momento è.:
        _PartialTransactValue = m_CurrentAmountScalable

        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form
        _UpdateResultData("FILL_DATA_RESPONSE_INIT", _PartialTransactValue, Nothing)

        ' Riprendo l'elenco riportato dall'hardware
        ' per ogni taglio e colloco ricopiandolo il 
        ' pezzo interessato
        For Each itm As Object In m_LastResponseRawArgentea.ListBPsEvaluated

            ' Questo dall'hardware 
            ' portiamo un code contatore
            Dim paidValue As Decimal = itm.Value
            Dim faceValue As Decimal = itm.Value

            m_CurrentBarcodeScan = itm.Key
            'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

            ' Aggiungo in una collection specifica in uso
            ' interno l'elemento Buono appena accodato in
            ' modo univoco rispetto al suo BarCode.
            Dim ItemNew As PaidEntry = WriterResultDataList.NewPaid(
                m_CurrentBarcodeScan,
                Value:=paidValue.ToString("###,##0.00"),
                FaceValue:=faceValue.ToString("###,##0.00"),
                Emitter:=RefTo_Transaction_Identifier,
                IdTransactionCrc:=m_LastResponseRawArgentea.TerminalID
            )

            ' Numero di elementi elaborati
            m_TotalBPElaborated_CS += 1

            ' Aggiorniamo i parziali
            m_PartialBPUsedToPay_XS += 1

            ' Il Pagato Incrementa come Importo Pagato
            m_PartialPayed_XS += paidValue
            ' Il Pagabile decrementa il Suo Pagambile
            _PartialTransactValue -= paidValue
            ' E se si eccede Produce il Resto
            If m_PartialPayed_XS > m_CurrentAmountScalable Then
                m_PartialValueExcedeed_XS += paidValue
            End If

            ' Aggiorno i Dati in corso per il ResulData e la Vista sul Form
            _UpdateResultData("FILL_DATA_RESPONSE_UPDATE", _PartialTransactValue, ItemNew)

        Next

        ' ** OK --> ATTESA COMPLETATA e corretamente chiamata vs Hardware Terminal POS
        LOG_Debug(getLocationString(funcName), "BP filled form Data From POS comunication " & m_LastResponseRawArgentea.SuccessMessage)

    End Sub

    ''' <summary>
    '''     Chiamata successiva ad un azione
    '''     di Void per la rapp. dei dati.
    '''     ( Completa l'Azione Prima del Chiudi )
    ''' </summary>
    ''' <param name="RefTo_Transaction_Identifier"></param>
    Private Sub CalculateAndFillMultiItemsInReturns(RefTo_Transaction_Identifier As String)
        Dim funcName As String = "CalculateAndFillMultiItemsInReturns"

        ' Status sul Form emulazione POS Hardware
        frmEmulation.SetStatus(PictureMultiStatusControlExpanse.enStatustype.Ok)

        'System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(100)

        '
        ' A differenza del Software  Creo  voci
        ' di TA quanti sono stati inoltrati nel
        ' dispositivo.
        '               --> Conteggio reale per il Record
        Dim NumBPEEvaluated As Integer = m_LastResponseRawArgentea.NumBPEvalutated        ' <-- Il Numero dei buoni che mi torna la chiamata
        Dim TotBPEEvaluated As Decimal = m_LastResponseRawArgentea.Amount                 ' <-- Il Totale che mi torna la chiamata

        ' Resetto inizialmente i dati sul form .
        m_TotalVoided_CS = 0
        m_VoidableAmount = 0
        m_TotalInvalid_CS = 0
        'm_CurrentAmountScalable
        m_VoidableAmount = m_PartialPayed_XS - m_PartialVoided_XS

        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form
        _UpdateResultData("FILL_DATA_VOID_RESPONSE_INIT", m_VoidableAmount, Nothing)

        ' Mi riserbo l'elenco originale
        'Dim qCopyOrigin As DataResponse.IResultDataList(Of PaidEntry) = CType(CType(WriterResultDataList, ICloneable).Clone(), DataResponse.IResultDataList(Of PaidEntry))
        Dim pp As Object = CType(WriterResultDataList, ICloneable).Clone()
        Dim qCopyOrigin As DataResponse.IResultDataList(Of PaidEntry) = CType(pp, DataResponse.IResultDataList(Of PaidEntry))

        ' Riprendo l'elenco riportato dall'hardware
        ' per ogni taglio e colloco ricopiandolo il 
        ' pezzo interessato ad essere decurtato
        For Each itm As Object In m_LastResponseRawArgentea.ListBPsEvaluated

            ' Questo dall'hardware 
            ' portiamo un code contatore
            _PartialTransactValue = itm.Value

            m_CurrentBarcodeScan = itm.Key
            'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

            ' Aggiungo in una collection specifica in uso
            ' interno l'elemento Buono appena accodato in
            ' modo univoco rispetto al suo BarCode.
            Dim ItemNew As PaidEntry = WriterResultDataList.NewPaid(
                m_CurrentBarcodeScan,
                Value:=_PartialTransactValue.ToString("###,##0.00"),
                FaceValue:=_PartialTransactValue.ToString("###,##0.00"),
                Emitter:=RefTo_Transaction_Identifier,
                IdTransactionCrc:=m_LastResponseRawArgentea.TerminalID
            )

            '
            ' Se l'elemtno nell'elenco di questa sessione
            ' esiste allora può essere stornato, altrimenti
            ' non potrà essere stornato.
            '
            Dim _Count_Elements_With_SomeFaceValue_InList As Integer = qCopyOrigin.CountElementsWithSomeFaceValue(ItemNew.FaceValue, False)
            Dim _Count_Elements_With_SomeFaceValue_Voided As Integer = qCopyOrigin.CountElementsWithSomeFaceValue(ItemNew.FaceValue, True)

            ' Numero di elementi elaborati
            m_TotalBPElaborated_CS += 1

            If _Count_Elements_With_SomeFaceValue_InList = 0 Then

                ' L'elemento non esite quindi è invalido
                ItemNew.Invalid = True
                ItemNew.InfoExtra = "Element not valid for void"

            Else

                If _Count_Elements_With_SomeFaceValue_InList = _Count_Elements_With_SomeFaceValue_Voided Then

                    ' L'elemento eistte ma tutti i pezzi sono stati stornati quindi è invalido questo da inserire
                    ItemNew.Invalid = True
                    ItemNew.InfoExtra = "Element not valid for void excedeed exisestant"

                ElseIf _Count_Elements_With_SomeFaceValue_Voided > 0 Then

                    '
                    ' L'elemento esiste già tra i vecchi 
                    ' elementi stornati e non lo conteggio
                    '
                    ItemNew.Voided = True
                    ItemNew.InfoExtra = "Element already voided"

                Else

                    '
                    '
                    '
                    ItemNew.Voided = True
                    ItemNew.InfoExtra = "Element current voided"

                End If

            End If

            Dim CValueDecimal As Decimal = ItemNew.DecimalValue

            ' Aggiornamento dello stornabile e del pagato e del non contabilizzato nelle sessioni precedenti
            If Not ItemNew.Voided And Not ItemNew.Invalid Then
                ' BP Usati come pagamento
                m_PartialPayed_XS += CValueDecimal
                m_PartialBPUsedToPay_XS += 1
                m_VoidableAmount += CValueDecimal           ' Usato Come stornabile in modalità VOID
                If ItemNew.InfoExtra = "Element already voided" Then
                    _PartialTransactValue += CValueDecimal
                End If
            ElseIf Not ItemNew.Invalid Then
                ' BP Usati come storno
                m_PartialVoided_XS += CValueDecimal
                m_PartialBPUsedToVoid_XS += 1
                m_VoidableAmount -= CValueDecimal           ' Usato Come stornabile in modalità VOID
            Else
                ' BP Non Validi e non Contabilizzati
                m_PartialBPNotValid_XS += 1
                m_PartialInvalid_XS += CValueDecimal
                _PartialTransactValue -= CValueDecimal

            End If

            If (m_PartialPayed_XS - m_PartialVoided_XS) < 0 Then
                m_PartialValueExcedeed_XS = -(m_PartialPayed_XS - m_PartialVoided_XS)
            End If

            ' Aggiungo al dataResult per il calcolo in
            ' uscita da usare per aggiornare nella TA i MetaData
            'WriterResultDataList.Add(ItemNew)

            '
            ' In Prefill il Payable viene
            ' aggiornato dall'insieme prefillato appunto
            '
            m_CurrentAmountScalable = m_PartialPayed_XS

            ' Aggiorno il DataResponse e la vista
            _UpdateResultData("FILL_DATA_VOID_RESPONSE_UPDATE", _PartialTransactValue, ItemNew)

        Next

    End Sub

    ''' <summary>
    '''     Richiesta per le Info del Ticket
    '''     dove recuper i dati per i BPE a disposizione dell'utente.
    '''     ( Azione richiesta in sessione Hardware )
    ''' </summary>
    Private Sub CalculateAndFillMultiItemsInfoTicket(RefTo_Transaction_Identifier As String)
        Dim funcName As String = "CalculateAndFillMultiItemsInfoTicket"

        ' Status sul Form emulazione POS Hardware
        frmEmulation.SetStatus(PictureMultiStatusControlExpanse.enStatustype.Ok)

        'System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(100)

        '
        ' A differenza del Software  Creo  voci
        ' di TA quanti sono presenti nel Ticket
        ' dell'utente inserito nel dispositivo.
        '               --> Conteggio reale per il Record
        Dim NumBPEEvaluated As Integer = m_LastResponseRawArgentea.NumBPEvalutated        ' <-- Il Numero dei buoni che mi torna la chiamata
        Dim TotBPEEvaluated As Decimal = m_LastResponseRawArgentea.Amount                 ' <-- Il Totale che mi torna la chiamata

        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form
        _UpdateResultData("FILL_DATA_INFO_RESPONSE_INIT", 0, Nothing)

        ' Aggiungo sulla vista da qui il Numero dei BP della Ticket Card 
        ' e il totoale complessivo dei buoni pagabili dall'utente
        'frmEmulation.ForInfoCardView(NumBPEEvaluated,TotBPEEvaluated)

        Dim _PartialTotalCard As Decimal

        ' Riprendo l'elenco riportato dall'hardware
        ' per ogni taglio e colloco ricopiandolo il 
        ' pezzo interessato ad essere decurtato
        For Each itm As Object In m_LastResponseRawArgentea.ListBPsEvaluated

            ' Questo dall'hardware 
            ' portiamo un code contatore
            _PartialTotalCard = itm.Value

            ' Aggiungo in una collection specifica in uso
            ' interno l'elemento Buono appena accodato in
            ' modo univoco rispetto al suo BarCode.
            Dim ItemNew As PaidEntry = New PaidEntry(
                "BPE",
                _PartialTotalCard.ToString("###,##0.00"),
                _PartialTotalCard.ToString("###,##0.00"),
                RefTo_Transaction_Identifier,
                m_LastResponseRawArgentea.TerminalID
            )

            ' Aggiorno il DataResponse e la vista
            _UpdateResultData("FILL_DATA_INFO_RESPONSE_UPDATE", _PartialTotalCard, ItemNew)

        Next

        ' Completo il DataResponse e la vista
        _UpdateResultData("FILL_DATA_INFO_RESPONSE_END", _PartialTotalCard, Nothing)


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
            '_UpdateFormEmulation()
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
            '_UpdateFormEmulation()
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
            '_UpdateFormEmulation()
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
            '_UpdateFormEmulation()
        End Set
    End Property

    ''' <summary>
    '''     Aggiorna se in questo momento
    '''     è visualizzato il form di emulazione
    '''     i dati con  i rispettivi valori sui 
    '''     controli del form per la modalità e lo
    '''     scopo in corso.
    ''' </summary>
    Private Sub _UpdateFormEmulation()
        If Not frmEmulation Is Nothing Then
            If m_CommandToCall = enCommandToCall.Payment Then
                frmEmulation.UpdateDataValues(m_PayableAmount,, m_PaidAmount)
            ElseIf m_CommandToCall = enCommandToCall.Void Then
                frmEmulation.UpdateDataValues(m_VoidableAmount,, ,,, m_VoidAmount)
            End If
        End If
    End Sub

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

            '
            ' Controllo che il numero di BP in totale 
            ' globali alla Vendita corrente rispetto a
            ' l'opzione di utilizzo sia non superato.
            '
            If m_OPT_MaxNumBPSomeSession <> 0 And ((m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc) <= 0) And
                Not m_CommandToCall = enCommandToCall.Void Then

                ' Signal (Numero buoni dematerializzabili superato rispetto a quelli permessi)
                _SetOperationStatus(
                    funcName,
                    NOT_OPT_ERROR_NUMBP_EXCEDEED,
                    "Il numero di titoli di pagamento per questa vendita è stato superato!!",
                    PosDef.TARMessageTypes.TPINFORMATION, True
                )

                ' Status in errore
                m_ServiceStatus = enProxyStatus.InError

                Return ' Immediato

            End If

            ' Istanza della Lib Argentea MONETICA
            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()

            ' Preparo la classe per il Set
            ' di risultati da restituire.
            '_DataResponse = New DataResponse()
            If Not m_PrefillVoidable Is Nothing Then
                Me.PrefillVoidable = m_PrefillVoidable
            End If

            ' Flag locale che stato attivo
            m_bWaitActive = True

            ' BEHAVIOR
            If m_TypeProxy = enTypeProxy.Service Then

                ' Preparo la risposta (Elaborazione per BP Cartacei)
                m_TypeBPElaborated_CS = enTypeBP.TicketsRestaurant

                '
                ' Avvio il servizio con la gestione
                ' del POS software tramite service.
                ' Rimane in idle sul form attivo di emulazione.
                '
                StartPosSoftware()

            Else

                ' Preparo la risposta (Elaborazione per Card con BP a tagli)
                m_TypeBPElaborated_CS = enTypeBP.TicketsCard

                '
                ' Avvio il pos locale con la gestione
                ' del POS hardware tramite terminale.
                '
                StartPosHardware()

            End If

            ' Flag locale che stato attivo
            m_bWaitActive = False

        Else
            ' Sollevo l'eccezione
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_PROXY_ALREADY_RUNNING, "Errore nell'esecuzione del Proxy -- Proxy attualmente già in esecuzione --")
        End If

    End Sub

    ''' <summary>
    '''     Chiamate Call ad API di servizio su
    '''     service Argentea remoto.
    ''' </summary>
    ''' <param name="APItoCALL">Il nome dell'api da richiamare presso l'endpoint di argentea</param>
    ''' <returns>Restituisce immeditamente lo stato OK KO o Error di Proxy Argentea <see cref="enProxyStatus"/></returns>
    Friend Function CallAPI(APItoCALL As String, ParamArray Arguments() As KeyValuePair(Of String, Object)) As enProxyStatus
        Dim funcName As String = "ProxyArgentea.CallAPI"
        Dim Response As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        m_CurrentApiNameToCall = APItoCALL

        ' ARGS
        ' Istanza della Lib Argentea MONETICA
        ArgenteaCOMObject = Nothing
        ArgenteaCOMObject = New ARGLIB.argpay()


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
                    Response = _API_Close()

                Case "SINGLEVOID"

                    ' Chiamata al metodo remoto   VOID
                    ' per stornare un titolo già validato
                    ' sempre che la transazione lo permetta.
                    Response = _API_SingleVoid(Arguments)

                Case Else

                    ' Signal (Api non valida)
                    _SetOperationStatus(funcName,
                        OPR_ERROR_API_NOT_VALID,
                        "Non è presente un API sul servizio di Argentea con questo nome",
                    PosDef.TARMessageTypes.TPSTOP, True
                    )

                    Return enProxyStatus.InError

            End Select

            If Response = ArgenteaFunctionsReturnCode.OK Then
                Return enProxyStatus.OK
            ElseIf Response = ArgenteaFunctionsReturnCode.KO Then
                Return enProxyStatus.KO
            End If

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
            _ClearWaitScreen()

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

                ' Mostra la finestra di Wait avviata al connect
                _ShowWaitScreen(2, "Connessione", "Connessione in corso")

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
        _ClearWaitScreen()

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
        _flagCallOnetimeResetIncrement = False
        'm_ProgressiveCall = 1

        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        ' Totalizzatori di sessione
        _ResetSessionCountersFinals()

        ' Totalizzatori Parziali
        _ResetSessionCountersPartials()

        '
        m_CurrentTransactionID = String.Empty
        m_CurrentAmountScalable = 0
        m_InitialPaymentsInTA = 0
        m_TotalBPInUseOnCurDoc = 0
        ' 
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
        Dim actApiCall As enApiToCall
        Dim funcName As String = "_API_Close"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Partiamo che non sia OK l'esito su chiamata remota Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' Chiusura di un BP (non si potrà più fare un annullo)
        actApiCall = enApiToCall.Close
        metdName = "CloseTicketBC"

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

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' L'Iniziailizzazione deve essere chiamata una volta sola nel contesto
            ' della sessione in corso, se la sessione sarà conlcusa sarà  chiamata
            ' nuovamente.
            FLAG_STATIC_INITIALIZATED = False

            ' ** OK --> API CLOSE corretamente chiamata ad Argentea
            LOG_Debug(getLocationString(funcName), "API " & m_CurrentApiNameToCall & " successfuly on response with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            ' Set dell'errore sullo status della risposta
            m_LastStatus = GLB_OPT_ERROR_ON_API_CLOSE
            m_LastErrorMessage = "L'operazione di chiudere la sessione in corso ha dato un KO " + m_LastResponseRawArgentea.ErrorMessage

            ' ** KO --> Risposta da API KO quindi non validata sul sistema remoto.
            LOG_Debug(getLocationString(funcName), "API " & m_CurrentApiNameToCall & " remote failed on response from service argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

    ''' <summary>
    '''     API Void Single presso Argentea
    ''' </summary>
    ''' <param name="Arguments">
    '''     Come parametri obbligatori per questa API è necessario passare.:
    ''' 
    '''         Key = BarCode               Value = Un valore stringa identificante il Barcode da stornare
    '''         Key = IdCrcTransaction      Value = Un valore stringa identificante la transazione associata alla ripostasta da argentea a questo barcode quando è stato dematerilizzato
    ''' 
    ''' </param>
    ''' <returns><see cref="ArgenteaFunctionsReturnCode"/></returns>
    Private Function _API_SingleVoid(ParamArray Arguments() As KeyValuePair(Of String, Object)) As ArgenteaFunctionsReturnCode

        Dim sender As FormEmulationArgentea = Nothing ' New FormEmulationArgentea() ' Form fittizio


        ' 0 = "BarCode" 1 = "IdCrcTransaction" 2 = "Value"
        Dim barcode As String = Arguments(0).Value
        Dim refToCrcTransactionId As String = Arguments(1).Value
        Dim TotValueBP As String = Arguments(2).Value

        'msgUtil.ShowMessage(m_TheModcntr, "Barcode = " + barcode + " Trans. = " + refToCrcTransactionId, "LevelITCommonModArgentea_", PosDef.TARMessageTypes.TPINFORMATION)

        _DataResponse = New DataResponse(TotValueBP, 1, TotValueBP, 0, 0, 0, 0, 0)

        ' L'elemento da rimuovere (Essendo singolo sono rivlevanti solo ID Transactione e barcode)
        Dim ItemNew As PaidEntry = WriterResultDataList.NewPaid(
            barcode,
            Value:=TotValueBP,
            FaceValue:=TotValueBP,
            Emitter:="",
            CodeIssuer:="",
            NameIssuer:="",
            IdTransactionCrc:=refToCrcTransactionId
        )

        ' Per questa API sfrutto la gestione interna del PROXY 
        ' con l'evento sull'emulatore software dei VOID da gruppo.
        m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
        Me.BarcodeRemoveHandler(sender, barcode)
        m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

        If m_LastStatus Is Nothing Then
            Return ArgenteaFunctionsReturnCode.OK
        Else
            Return ArgenteaFunctionsReturnCode.KO
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
    Private Sub CreateInstanceFormForEmulationOrWaitAction()
        Dim funcName As String = "CreateInstanceFormForEmulationOrWaitAction"

        ' Istanza del form di appoggio  ad uso 
        ' operatore per l'inserimento per ogni
        ' BP che deve partecipare al pagamento.
        frmEmulation = m_TheModcntr.GetCustomizedForm(GetType(FormEmulationArgentea), NO_STRETCH) ' STRETCH_TO_SMALL_WINDOW)
        If frmEmulation Is Nothing Then

            ' Sollevo l'eccezione
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

        Else

            ' Tipizzazione dello scopo

            If Not m_TheModcntr Is Nothing Then

                '
                '   Formato delle stringhe decimali (##.###)
                '
                frmEmulation.FormatData = m_TheModcntr.getFormatString4Price()

            End If

        End If

        Try

            ' 
            ' Riporto come property al form  da 
            ' visualizzare per una sua gestione
            ' interna il Controller e la Transazione
            '
            frmEmulation.theModCntr = m_TheModcntr
            frmEmulation.taobj = m_taobj

            '
            ' Resetto gli Hanlders atomici
            '
            ResetHandlers()

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
                    _SetFormForUseServiceVoid(FormEmulationArgentea.enScopeUseForm.VOID)

                ElseIf m_TypeProxy = enTypeProxy.Pos Then

                    ' Se per usare la Modalità Pos in uso con il terminale 
                    _SetFormForUsePosVoid(FormEmulationArgentea.enScopeUseForm.VOID)

                End If

            ElseIf m_CommandToCall = enCommandToCall.Payment Then

                ' Se gestiamo la modalità Pagamento con BP delle API di Argentea

                If m_TypeProxy = enTypeProxy.Service Then

                    ' Se per usare la Modalità Service in emulazione Software
                    _SetFormForUseServicePayment(FormEmulationArgentea.enScopeUseForm.DEMAT)

                ElseIf m_TypeProxy = enTypeProxy.Pos Then

                    ' Se per usare la Modalità Pos in uso con il terminale Hardware
                    _SetFormForUsePosPayment(FormEmulationArgentea.enScopeUseForm.DEMAT)

                End If

            End If


        Catch ex As Exception

            ' Sollevo l'eccezione
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_INSTANCE, "Errore nell'istanziare il form nell'emulatore -- Contattare Assistenza --", ex)

        End Try

    End Sub

    ''' <summary>
    '''     Mostra il form di attesa con l'azione da  parte
    '''     di un operatore, gestisce eventi e contesti con
    '''     l'operatore.
    ''' </summary>
    Private Sub ShowAndIdleOnFormForAction(ByVal NotIdle As Boolean)

        '
        ' Mostro il form per la gestione e comunicazione
        ' con il Servizio remoto di convalida su azioni di
        ' Dematerializzazione e Storno.
        '
        frmEmulation.Show() ' non modal VB Dialog

        ' Dispongo le proprietà del Form  Cassa
        ' ripreso nel Controller globale per la
        ' preparazione a non prendere lo  status
        ' attivo durante la scansione dove si sta 
        ' operando con il controllo locale del form 
        ' che ha la textbox per prendere i codici EAN
        m_TheModcntr.DialogActiv = True
        m_TheModcntr.DialogFormName = frmEmulation.Text
        m_TheModcntr.SetFuncKeys((False))

        '
        ' (Idle) sul Form Locale
        ' Finestra di dialogo avviata e rimango in idle 
        ' finchè l'operatore non finisce le azioni necessarie.
        '
        frmEmulation.bDialogActive = True

        ' (Idle)
        If Not NotIdle Then

            ' Form in attesa

            ' Aggiorna all'entrata i totali
            ' sul form di emulazione.
            '_UpdateFormEmulation()

            Do While frmEmulation.bDialogActive = True
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()
            Loop

        Else

            ' Form nascosto (Uscita immediata)
            For x As Int16 = 1 To 5

                System.Threading.Thread.Sleep(50)
                'System.Windows.Forms.Application.DoEvents()

            Next

        End If

    End Sub

    ''' <summary>
    '''     Trasforma il form per rimuovere tutti i buoni di tipo
    '''     cartace gestiti dall'emulatore software di POS
    ''' </summary>
    Private Sub _SetFormForUseServiceVoid(Scope As FormEmulationArgentea.enScopeUseForm)

        ' Inizializza il Form i Controlli
        ' e la lingua per il testo sui controlli.
        frmEmulation.InitializeFormSoftware(Scope)

        ' In Base alla proprietà di Tp .NEt
        'definisamo questo comportamento per
        'visualizzare o meno gli elementi eliminati
        frmEmulation.ModeViewElementsDeleteted = st_Parameters_Argentea.BP_ViewModeElementsDeleteted

        ' Definisce il comportamento sulla 
        ' vista per i totalizzatori e in 
        ' base a questo parametro saranno 
        ' visti gli elementi di totale in vari modi
        frmEmulation.ModeViewTotalsAndPartials = st_Parameters_Argentea.BP_ViewModeTotalsAndPartials    ' BP_VIEV_MODE_TOTALS_AND_PARTIALS

        ' In Base alla proprietà di Tp .NEt
        ' definiamo questo comportamento per
        ' visualizzare o meno il tasto OK 
        ' per chiudere il form a fine operazione ( In Software sarà sempre False)
        frmEmulation.AutoCloseOnCompleteOperation = False

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere un taglio in
        ' base a dove clicca.
        '
        '
        AddHandler frmEmulation.BarcodeRead, AddressOf BarcodeReadVoidHandler
        AddHandler frmEmulation.BarcodeCheck, AddressOf BarcodeCheckValidCodeHandler
        AddHandler frmEmulation.BarcodeRemove, AddressOf BarcodeRemoveVoidHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        AddHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler

    End Sub

    ''' <summary>
    '''     Trasforma il form per utilizzarlo solo uso  copione
    '''     intanto che c'è la comunicazione con l'hardware con
    '''     un unica label riepilogativa.
    ''' </summary>
    Private Sub _SetFormForUsePosVoid(Scope As FormEmulationArgentea.enScopeUseForm)

        ' Inizializza il Form i Controlli
        ' e la lingua per il testo sui controlli.
        frmEmulation.InitializeFormHardware(Scope)

        ' In Base alla proprietà di Tp .NEt
        ' definiamo questo comportamento per
        ' visualizzare o meno gli elementi eliminati
        frmEmulation.ModeViewElementsDeleteted = st_Parameters_Argentea.BP_ViewModeElementsDeleteted

        ' Definisce il comportamento sulla 
        ' vista per i totalizzatori e in 
        ' base a questo parametro saranno 
        ' visti gli elementi di totale in vari modi
        frmEmulation.ModeViewTotalsAndPartials = st_Parameters_Argentea.BP_ViewModeTotalsAndPartials    ' BP_VIEV_MODE_TOTALS_AND_PARTIALS

        ' In Base alla proprietà di Tp .NEt
        ' definiamo questo comportamento per
        ' visualizzare o meno il tasto OK 
        ' per chiudere il form a fine operazione
        frmEmulation.AutoCloseOnCompleteOperation = st_Parameters_Argentea.BPE_AutoCloseOnCompleteOperation

        ' Definisce di mostrare un tasto per ritentare
        ' il richiamao al POS Hardware a condizione che
        ' l'autocomplete sia a false e che ci sia un errore
        ' di tentata comunicazione in corso
        frmEmulation.ReattemptOperationOnErrors = st_Parameters_Argentea.BPE_ReattemptToCompleteOperation

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere  i tagli e gli
        ' importi dall'azione utente sul POS Hardware.
        '
        AddHandler frmEmulation.ReattemptOperation, AddressOf ConnectHardwareVoidHandler
        AddHandler frmEmulation.InfoTicketCall, AddressOf ConnectHardwareInfoTicketHandler
        AddHandler frmEmulation.ReloadTransactionCall, AddressOf ConnectHardwareReloadLastTransactionHandler


        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        AddHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler


    End Sub

    ''' <summary>
    '''     Trasforma il form per aggiungere buoni di tipo
    '''     cartace gestiti dall'emulatore software di POS.
    ''' </summary>
    Private Sub _SetFormForUseServicePayment(Scope As FormEmulationArgentea.enScopeUseForm)

        ' Inizializza il Form i Controlli
        ' e la lingua per il testo sui controlli.
        frmEmulation.InitializeFormSoftware(Scope)

        ' In Base alla proprietà di Tp .NEt
        'definisamo questo comportamento per
        'visualizzare o meno gli elementi eliminati
        frmEmulation.ModeViewElementsDeleteted = st_Parameters_Argentea.BP_ViewModeElementsDeleteted

        ' Definisce il comportamento sulla 
        ' vista per i totalizzatori e in 
        ' base a questo parametro saranno 
        ' visti gli elementi di totale in vari modi
        frmEmulation.ModeViewTotalsAndPartials = st_Parameters_Argentea.BP_ViewModeTotalsAndPartials    ' BP_VIEV_MODE_TOTALS_AND_PARTIALS

        ' In Base alla proprietà di Tp .NEt
        ' definiamo questo comportamento per
        ' visualizzare o meno il tasto OK 
        ' per chiudere il form a fine operazione ( In Software sarà sempre False)
        frmEmulation.AutoCloseOnCompleteOperation = False

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di rimuovere un taglio in
        ' base a dove clicca.
        '
        AddHandler frmEmulation.BarcodeRead, AddressOf BarcodeReadHandler
        AddHandler frmEmulation.BarcodeCheck, AddressOf BarcodeCheckValidCodeHandler
        AddHandler frmEmulation.BarcodeRemove, AddressOf BarcodeRemoveHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        AddHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler

    End Sub

    ''' <summary>
    '''     Trasforma il form per utilizzarlo solo uso  copione
    '''     intanto che c'è la comunicazione con l'hardware con
    '''     un unica label riepilogativa.
    ''' </summary>
    Private Sub _SetFormForUsePosPayment(Scope As FormEmulationArgentea.enScopeUseForm)

        ' Inizializza il Form i Controlli
        ' e la lingua per il testo sui controlli.
        frmEmulation.InitializeFormHardware(Scope)

        ' In Base alla proprietà di Tp .NEt
        'definisamo questo comportamento per
        'visualizzare o meno gli elementi eliminati
        frmEmulation.ModeViewElementsDeleteted = st_Parameters_Argentea.BP_ViewModeElementsDeleteted

        ' Definisce il comportamento sulla 
        ' vista per i totalizzatori e in 
        ' base a questo parametro saranno 
        ' visti gli elementi di totale in vari modi
        frmEmulation.ModeViewTotalsAndPartials = st_Parameters_Argentea.BP_ViewModeTotalsAndPartials    ' BP_VIEV_MODE_TOTALS_AND_PARTIALS

        ' In Base alla proprietà di Tp .NEt
        ' definiamo questo comportamento per
        ' visualizzare o meno il tasto OK 
        ' per chiudere il form a fine operazione
        frmEmulation.AutoCloseOnCompleteOperation = st_Parameters_Argentea.BPE_AutoCloseOnCompleteOperation

        ' Definisce di mostrare un tasto per ritentare
        ' il richiamao al POS Hardware a condizione che
        ' l'autocomplete sia a false e che ci sia un errore
        ' di tentata comunicazione in corso
        frmEmulation.ReattemptOperationOnErrors = st_Parameters_Argentea.BPE_ReattemptToCompleteOperation

        '
        ' Preparo ad accettare l'handler degli eventi gestiti
        ' durante l'azione utente di aggiungere i tagli e gli
        ' importi dall'azione utente sul POS Hardware.
        '
        AddHandler frmEmulation.ReattemptOperation, AddressOf ConnectHardwareDematHandler
        AddHandler frmEmulation.InfoTicketCall, AddressOf ConnectHardwareInfoTicketHandler
        AddHandler frmEmulation.ReloadTransactionCall, AddressOf ConnectHardwareReloadLastTransactionHandler

        '
        ' Evento chiave all'ok del form o alla chiusura del pos
        ' per il collect dei dati in risposta al chiamante.
        '
        AddHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler

    End Sub

    Private Sub ResetHandlers()

        RemoveHandler frmEmulation.BarcodeRead, AddressOf BarcodeReadHandler
        RemoveHandler frmEmulation.BarcodeCheck, AddressOf BarcodeCheckValidCodeHandler
        RemoveHandler frmEmulation.BarcodeRemove, AddressOf BarcodeRemoveHandler
        '
        RemoveHandler frmEmulation.ReattemptOperation, AddressOf ConnectHardwareDematHandler
        RemoveHandler frmEmulation.ReloadTransactionCall, AddressOf ConnectHardwareReloadLastTransactionHandler
        RemoveHandler frmEmulation.InfoTicketCall, AddressOf ConnectHardwareInfoTicketHandler
        '
        RemoveHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler

    End Sub


#End Region

#Region "Statico all'applicativo Parser Functions Privates"

    ''' <summary>
    '''     Definisce e ricava pr il protocollo Argentea  nella  funzione
    '''     helper globale quale è la chiamata su cui nel CSV di risposta
    '''     di Argentea richiamare per il Parsing e quela il carattere di
    '''     separazione nella stringa e quale la frazione sui valori.
    ''' </summary>
    ''' <param name="ApiToCall">Il tipo di azione <see cref="enApiToCall"/> per il parser su cui è da ricavare il corrispondente nell'helper <see cref="InternalArgenteaFunctionTypes"/> la funzione parser di protocollo e i corrispondenti separatori e frazioni </param>
    ''' <returns>Restituisce il Parser Type (Tupla 1° <see cref="InternalArgenteaFunctionTypes"/> riflette in base all'api il tipo da parsare sull'helper, il 2° il Carattere separatore usato nel protocollo, il 3° la Frazione usata nel protocollo)</returns>
    Friend Shared Function GetSplitAndFormatModeForParsing(ApiToCall As enApiToCall) As Tuple(Of InternalArgenteaFunctionTypes, Char, Integer)

        ' Behavior
        Dim _ParsingMode As InternalArgenteaFunctionTypes = InternalArgenteaFunctionTypes.Initialization_AG
        Dim _ParseSplitMode As Char = "-"
        Dim _ParseFractMode As Integer = 100

        ' Se stiamo usando il Parsing per le chiamate
        ' API allora usiamo questo.
        If ApiToCall <> enApiToCall.None Then

            _ParseSplitMode = "-"
            _ParseFractMode = 100

            Select Case ApiToCall
                Case enApiToCall.Initialization
                    _ParsingMode = InternalArgenteaFunctionTypes.Initialization_AG
                Case enApiToCall.ResetCounter
                    _ParsingMode = InternalArgenteaFunctionTypes.ResetCounter_AG
                Case enApiToCall.Confirmation
                    _ParsingMode = InternalArgenteaFunctionTypes.Confirmation_AG
                Case enApiToCall.Close
                    _ParsingMode = InternalArgenteaFunctionTypes.Close_AG
                Case enApiToCall.CheckBP
                    _ParsingMode = InternalArgenteaFunctionTypes.Check_BP
                Case enApiToCall.SingleVoid
                    _ParsingMode = InternalArgenteaFunctionTypes.SingleVoid_BP
                Case enApiToCall.SinglePayment
                    _ParsingMode = InternalArgenteaFunctionTypes.SinglePaid_BP
                Case enApiToCall.MultiplePayments
                    _ParseSplitMode = ";"
                    _ParsingMode = InternalArgenteaFunctionTypes.MultiPaid_BP
                Case enApiToCall.MultipleVoids
                    _ParseSplitMode = ";"
                    _ParsingMode = InternalArgenteaFunctionTypes.MultiVoid_BP
                Case enApiToCall.InfoCardUser
                    _ParseSplitMode = ";"
                    _ParsingMode = InternalArgenteaFunctionTypes.MultiItemsIC_BP
            End Select

        End If

        Return Tuple.Create(Of InternalArgenteaFunctionTypes, Char, Integer)(_ParsingMode, _ParseSplitMode, _ParseFractMode)

    End Function

#End Region

#Region " Parsing _reflection metadata"

    ''' <summary>
    '''     Esegue il parsing del protocollo su una risposta di Argentea
    '''     per formulare il success o l'unseccessfull con il  messaggio 
    '''     ripreso dalla codifica del protocollo come risposta.
    ''' </summary>
    ''' <param name="func_Name">Il nome della funzione del Proxy di Argentea che ha sollevato questa eccezione</param>
    ''' <param name="Method_Name">Il nome del metodo sulla dll di Argentea da cui si sta ricevendo la response</param>
    ''' <param name="ret_Code">Il returno code che ha restituito la dll all'uscita</param>
    ''' <param name="Ref_MessageOut">Il messaggio dalla dll di argentea che è stato restituito</param>
    Private Function _ParseResponseAndMapToThisResult(func_Name As String, Method_Name As String, Api_Called As ClsProxyArgentea.enApiToCall, ret_Code As Integer, Ref_MessageOut As String) As ArgenteaFunctionReturnObject

        ' ** SOCKET COM --> Su GENERAL Errore di comunicazione protocollo o interno della dll (potrebbe risolversi anche in eccezione sul parsing della decodifica dell'errore su com )
        If ret_Code <> ArgenteaFunctionsReturnCode.OK Then

            ' ** KO --> Codificato Errore Socket 9001
            LOG_Error(func_Name, "Error COM Protocol  .: " & Ref_MessageOut + " RETCODE " + CStr(ret_Code))
            Ref_MessageOut = "ERRORE SOCKET"

        End If

        ' ** DECODIFIQUE PROTOCOL --> Tupla con i primi due eventaulmente per errore Comunicazione o di Parsing oppure con la codifica corretta.
        Dim Response As Tuple(Of Boolean, Boolean, Boolean, String, String, ArgenteaFunctionReturnObject) = ExceptionProxyArgentea.ParseProtocolForMapResponse(
            Api_Called, ret_Code, Ref_MessageOut, func_Name, Method_Name
        )

        ' ** NAT ERR COMUNICAZIONE o SISTEMA o PARSING --> Su GENERAL Errore di parsing sulla risposta nel protocollo  di risposta
        If Response.Item1 Or Response.Item2 Or Response.Item3 Then
            ' Errori di comunicazione o di Parsing o di Sistema sono Bloccanti (Uscita forzata)
            Throw New ExceptionProxyArgentea(func_Name, Method_Name, Api_Called, ret_Code, Ref_MessageOut, Response)
        End If

        ' ** ULTIMO CRC di risposta in risposta da argentea valido
        If Response.Item6.TerminalID <> String.Empty Then
            m_LastCrcTransactionID = Response.Item6.TerminalID
        End If

        ' ** NAT OK/KO CODIFICATO --> Parsing dela risposta (OK/KO) effettuato e nattato con successo
        Return Response.Item6                                   ' Response Parsed Object

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

                    ' Signal (Tentativo di stampa della ricevuta fallito)
                    _SetOperationStatus(funcName,
                        NOT_ERROR_PRINTER_FAILED,
                        "Il tentativo di stampare la ricevuta ha fallito con la comunicazione alla stampante!",
                        PosDef.TARMessageTypes.TPWARNING
                    )

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
    Private Sub StartPosHardware()
        Dim funcName As String = "StartPosHardware"

        ' Entrap sull'idle
        Try

            '
            ' Resetto i contatori iniziali
            ' per l'aggiornamento di stato.
            '
            m_TotalPayed_CS = 0
            m_TotalInvalid_CS = 0
            m_TotalVoided_CS = 0
            m_TotalInvalid_CS = 0

            '
            ' 1° STep ( Preparazione Form ed eventi sui controlli ) 
            '       Creo l'istanza adatta al tipo di azione di  un form 
            '       in emulazione POS Sofwware per il Service.
            '
            ' Prepara avviando a seconda della funzione per il Void o per il Pay
            '       --> _SetFormForUseServiceVoid(FormEmulationArgentea.enScopeUseForm.VOID)
            '           '*
            '           --> AddHandler frmEmulation.ReattemptClick, AddressOf ConnectHardwareVoidHandler
            '           '*
            '       --> _SetFormForUseServicePayment(FormEmulationArgentea.enScopeUseForm.DEMAT)
            '           '*
            '           --> AddHandler frmEmulation.ReattemptClick, AddressOf ConnectHardwareDematHandler
            '           '*
            '       *.*--> AddHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler ( Completa i Totali)
            CreateInstanceFormForEmulationOrWaitAction()

            ' Aggiorno il DataResponse e la vista (Iniziale)
            _UpdateResultData("INITIALIZE_EMULATOR_HARDWARE", m_CurrentAmountScalable, Nothing)

            ' Status
            m_ServiceStatus = enProxyStatus.InRunning

            ' In questa modalità avvio il waitwscreen
            ' modale a pieno schermo per  attendere 
            ' le operazioni dal Pos Locale collegato.
            _ShowWaitScreen(3, "Carico Terminale Locale", "Attesa azione utente")

            '
            '   2° STep ( Controllo se in VOID e prefillo sessione precedente)
            '           Aggiusta l'importo scalabile e riporta i dati iniziali.
            '
            If Not m_PrefillVoidable Is Nothing Then

                Me.PrefillVoidable = m_PrefillVoidable

                ' Signal Initialization Status
                _SetOperationStatusForm(
                        NOT_INFO_POS_INIT,
                        "Refill Data...",
                        FormEmulationArgentea.InfoStatus.Wait
                    )

                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(200)

                ' Carico i Dati Inziali Prefillati dalla chiamata
                Me.CalculateAndFillMultiItemsInitialize()

            End If

            ' Tolgo eventuali Wait Screen
            _ClearWaitScreen()

            '********************************************************
            '
            ' (Idle) --> senza AutoClose (ReadWrite)
            ShowAndIdleOnFormForAction(False)
            '
            '********************************************************

            ' Dichiaro come concluso correttamente tutto
            If m_ServiceStatus = enProxyStatus.InRunning Then

                ' Se era rimasto in Running e non InError
                ' tutto è filato liscio e torno con OK
                m_ServiceStatus = enProxyStatus.OK

                ' Signal Finish Status
                _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

            Else

                ' Signal (come errore di mancato collegamento all'hardware pos ma non bloccante)
                _SetOperationStatus(
                    funcName,
                    NOT_ERROR_SOFTWAREPOS_FAILED,
                    "Emulatore POS non avviato in modo corretto, controllare le impostazioni e lo stato!",
                    PosDef.TARMessageTypes.TPERROR
                )

                ' Scrive una riga di Log per monitorare....
                LOG_Info(getLocationString(funcName), m_ServiceStatus)

            End If

        Catch ex As Exception

            ' Signal Initialization Status
            _SetOperationStatusForm(
                NOT_INFO_POS_ERROR,
                "POS in error...",
                FormEmulationArgentea.InfoStatus.Error
            )

            '
            ' Alla chiusura della chiamata Errori possibili
            ' dal Proxy Argentea Gestiti come SOCKET o CHIAMATA 
            ' o interni al flow corrente come Errori non Gestiti Uknowed
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

            ' Con Ritorno a Status InError
            m_ServiceStatus = enProxyStatus.InError

        Finally

            Try

                '
                '  Provo a chiudere il Form del POS
                ' software se siamo in questa modalità.
                '
                If Not frmEmulation Is Nothing Then
                    m_TheModcntr.DialogActiv = False
                    m_TheModcntr.DialogFormName = ""
                    m_TheModcntr.SetFuncKeys((True))
                    m_TheModcntr.EndForm()
                    frmEmulation.Close()
                    frmEmulation = Nothing
                End If
                '
            Catch ex As Exception

                '
                ' Errori Interni UKNOWED da rimarcare
                '
                SetExceptionsStatus(funcName, ex, m_SilentMode)

                ' Con Ritorno a Status InError
                m_ServiceStatus = enProxyStatus.InError

            Finally

                ' Chiudo per memory leak con argentea
                ArgenteaCOMObject = Nothing

                ' Effettuo un Dispose forzato per 
                ' la chiusura del form su eccezioni.
                If Not frmEmulation Is Nothing Then
                    frmEmulation.Dispose()
                    frmEmulation = Nothing
                End If

            End Try

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla command che riceve il click per collegarsi
    '''     al pos hardware che detterà l'elenco dei BPE da consumare.
    ''' </summary>
    ''' <param name="sender">Il controllo command</param>
    Private Sub ConnectHardwareDematHandler(ByRef sender As Object, e As EventArgs)
        Dim funcName As String = "ConnectHardwareDematHandler"

        '_
        '
        '    Tipi di Repsonse su Hardware Argentea
        '    
        '    ->  OK su ADD BPE (CallDematerialize)      :::         OK;TRANSAZIONE ACCETTATA;5|2|1020|3|414;104;PELLEGRINI;  PAGAMENTO BUONI PASTO EFFEFFUATA
        '    ->  KO su ADD BPE (CallDematerialize)      :::         KO;    DATI NON RICEVUTI    ;;;;
        '_
        '

        Try

            If Not TypeOf sender Is FormEmulationArgentea Then

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

            '
            ' Notifica interrogazione
            '

            ' Signal Initialization Status
            _SetOperationStatusForm(
                NOT_INFO_POS_CALL,
                "Interrogazione in corso...",
                FormEmulationArgentea.InfoStatus.Wait
            )

            '
            ' ::Opzione:: Max BP pagabili per vendita.: 
            '
            '       Controllo che il numero di BP in totale 
            '       globali alla Vendita corrente rispetto a
            '       l'opzione di utilizzo sia non superato.
            '
            If m_OPT_MaxNumBPSomeSession <> 0 And (((WriterResultDataList.Count + 1) > (m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc)) Or ((m_TotalBPUsedToPay_CS + 1) > (m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc))) Then

                ' Signal (Numero buoni dematerializzabili superato rispetto a quelli permessi)
                _SetOperationStatus(funcName,
                    NOT_OPT_ERROR_NUMBP_EXCEDEED,
                    "Il numero di titoli di pagamento per questa vendita è stato superato!!",
                    PosDef.TARMessageTypes.TPINFORMATION, True
                )

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                Return

            End If

            '''     Controllo per il Buono in Corso se il valore
            '''     supera l'importo pagabile, e se lo è  allora
            '''     riporto lo status del pagamento generale come
            '''     in eccesso (sempre se l'opzione lo permette)
            '''     e se lo permette riporto per l'ultimo buono in 
            '''     corso il totale di facciata diverso dal valore effettivo
            '''     di pagato (per scrivefe un media di resto all'uscita)
            If m_PartialValueExcedeed_XS <> 0 Then

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                ' Signal (Importo già raggiunto)
                _SetOperationStatus(
                    funcName,
                    NOT_INFO_IMPORT_ALREADYCOMPLETED,
                    "L'importo da pagare è già stato completato per questa vendita completare con inoltro!!",
                    PosDef.TARMessageTypes.TPINFORMATION, True
                )

                Return

            End If

            ' Ripulisco eventuali wait screen in corso
            _ClearWaitScreen()

            ' Mostriamo il Wait
            _ShowWaitScreen(2, "Operazione in corso", "Demat in corso presso il Pos Argentea")

            ' Chiama per l'Hardware per il Payment
            ' e riportare l'insieme dei  pagamenti
            ' al totalizzatore. ( Idle --> Attesa POS )
            Dim _CallDematerialize As StatusCode = Me.CallMultiplePaymentsOnPosHardware(funcName)

            ' Una Volta richiamata la demateriliazzione 
            ' ed eventuale conferma ed hanno dato esito positivo.
            If _CallDematerialize = StatusCode.OK Then

                ' Signal sul Form OK
                _SetOperationStatusForm(
                    NOT_INFO_POS_DATA_OK,
                    "POS data ok!!",
                    FormEmulationArgentea.InfoStatus.OK, 9
                )
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()

                '
                '   Riprendo il Risultato e lo riporto
                '   come attuato e verificato dal POS Hardware.
                '
                CalculateAndFillMultiItemsInPayment(m_Transaction_Identifier)

                ' Cancelliamo Attesa precedente
                _ClearWaitScreen()

                '
                ' Non Abbiamo finito, dobbiamo ancora vedere
                ' se abbiamo superato l'importo da pagare  e
                ' se accettata l'opzione di avere un resto o 
                ' meno rispetto al pagato con buoni.
                '
                If Not CheckIfDematExcedeedTotToPay(sender) Then
                    Return
                End If

            Else

                ' Errata Dematerializzione dei BP su Card
                ' data dalla risposta argentea quindi su POS.
                _ClearWaitScreen()

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                _SetOperationStatus(
                    funcName,
                    NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                    m_LastErrorMessage,
                    PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                )

                Return

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Signal Initialization Status
            System.Threading.Thread.Sleep(200)
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla command che riceve il click per collegarsi
    '''     al pos hardware che detterà l'elenco dei BPE da stornare.
    ''' </summary>
    ''' <param name="sender">Il controllo command</param>
    Private Sub ConnectHardwareVoidHandler(ByRef sender As Object, e As EventArgs)
        Dim funcName As String = "ConnectHardwareVoidHandler"

        '_
        '
        '    Tipi di Repsonse su Hardware Argentea
        '    
        '    ->  OK su ADD BPE (CallDematerialize)      :::         OK;TRANSAZIONE ACCETTATA;5|2|1020|3|414;104;PELLEGRINI;  STORNO BUONI PASTO EFFEFFUATA
        '    ->  KO su ADD BPE (CallDematerialize)      :::         KO;    DATI NON RICEVUTI    ;;;;
        '_
        '

        Try

            If Not TypeOf sender Is FormEmulationArgentea Then

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

            '
            ' Nascondiamo il tentativo di connession
            ' al dispositivo pos hardware.
            '
            'frmEmulation.cmdReattempt.Visible = False

            ' Signal Initialization Status
            _SetOperationStatusForm(
                NOT_INFO_POS_CALL,
                "Interrogazione in corso...",
                FormEmulationArgentea.InfoStatus.Wait
            )

            '
            ' ::Opzione:: Max BP pagabili per vendita.: 
            '
            '       Controllo che il numero di BP in totale 
            '       globali alla Vendita corrente rispetto a
            '       l'opzione di utilizzo sia non superato.     ** IN QUESTO CASO ESCLUDO PERCHé IN VOID
            '
            If False Then 'm_OPT_MaxNumBPSomeSession <> 0 And (((WriterResultDataList.Count + 1) > (m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc)) Or ((m_TotalBPUsedToPay_CS + 1) > (m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc))) Then

                ' Signal (Numero buoni dematerializzabili superato rispetto a quelli permessi)
                _SetOperationStatus(funcName,
                    NOT_OPT_ERROR_NUMBP_EXCEDEED,
                    "Il numero di titoli di pagamento per questa vendita è stato superato!!",
                    PosDef.TARMessageTypes.TPINFORMATION, True
                )

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                Return

            End If

            '''     Controllo per il Buono in Corso se il valore
            '''     supera l'importo pagabile, e se lo è  allora
            '''     riporto lo status del pagamento generale come
            '''     in eccesso (sempre se l'opzione lo permette)
            '''     e se lo permette riporto per l'ultimo buono in 
            '''     corso il totale di facciata diverso dal valore effettivo
            '''     di pagato (per scrivefe un media di resto all'uscita)
            If m_PartialValueExcedeed_XS <> 0 Then

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                ' Signal (Importo già raggiunto)
                _SetOperationStatus(
                    funcName,
                    NOT_INFO_IMPORT_ALREADYCOMPLETED,
                    "L'importo da stornare è già stato completato per questa azione completare con inoltro!!",
                    PosDef.TARMessageTypes.TPINFORMATION, True
                )

                Return

            End If

            ' Ripulisco eventuali wait screen in corso
            _ClearWaitScreen()

            ' Mostriamo il Wait
            _ShowWaitScreen(2, "Operazione in corso", "Void in corso presso il Pos Argentea")

            ' Chiama per  l'Hardware per il Void
            ' e riportare l'insieme degli storni
            ' al totalizzatore. ( Idle --> Attesa POS )
            Dim _CallVoid As StatusCode = Me.CallMultipleVoidOnPosHardware(funcName)

            ' Una Volta richiamata la demateriliazzione 
            ' ed eventuale conferma ed hanno dato esito positivo.
            If _CallVoid = StatusCode.OK Then

                ' Signal sul Form OK
                _SetOperationStatusForm(
                    NOT_INFO_POS_DATA_OK,
                    "POS data ok!!",
                    FormEmulationArgentea.InfoStatus.OK, 9
                )
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()

                '
                '   Riprendo il Risultato e lo riporto
                '   come attuato e verificato dal POS Hardware.
                '
                CalculateAndFillMultiItemsInReturns(m_Transaction_Identifier)

                ' Cancelliamo Attesa precedente
                _ClearWaitScreen()

                '
                ' Non Abbiamo finito, dobbiamo ancora vedere
                ' se abbiamo superato l'importo da pagare  e
                ' se accettata l'opzione di avere un resto o 
                ' meno rispetto al pagato con buoni.
                '
                If Not CheckIfDematExcedeedTotToPay(sender) Then
                    Return
                End If

            Else

                ' Errata Void su Dematerializzazione già Esistente
                ' data dalla risposta argentea quindi errore POS.
                _ClearWaitScreen()

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                ' Signal (KO remoto su Void + Status POS per le codifiche da db personalizzate)
                _SetOperationStatus(
                    funcName,
                    NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                    m_LastErrorMessage,
                    PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                )

                Return

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Signal Initialization Status
            System.Threading.Thread.Sleep(200)
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla command che riceve il click per collegarsi
    '''     al pos hardware che chiederà di ricevere le info sulla card dell'utente.
    ''' </summary>
    ''' <param name="sender">Il controllo command</param>
    Private Sub ConnectHardwareInfoTicketHandler(ByRef sender As Object, e As EventArgs, ByRef InError As Boolean)
        Dim funcName As String = "ConnectHardwareInfoTicketHandler"

        '_
        '
        '    Tipi di Repsonse su Hardware Argentea
        '    
        '    ->  OK su ADD BPE (CallInfo)       :::         OK;RICHIESTA ACCETTATA;5|2|1020|3|414;104;PELLEGRINI;  BUONI PASTO PRESENTI
        '    ->  KO su ADD BPE (CallInfo)       :::         KO;    DATI NON RICEVUTI    ;;;;
        '_
        '

        Try

            If Not TypeOf sender Is FormEmulationArgentea Then

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

            ' Signal Initialization Status
            _SetOperationStatusForm(
                NOT_INFO_POS_CALL,
                "Interrogazione in corso...",
                FormEmulationArgentea.InfoStatus.Wait
            )

            ' Ripulisco eventuali wait screen in corso
            _ClearWaitScreen()

            ' Mostriamo il Wait
            _ShowWaitScreen(2, "Operazione in corso", "Richiesta info in corso presso il Pos Argentea")

            ' Chiama per l'Hardware per il Payment
            ' e riportare l'insieme dei  pagamenti
            ' al totalizzatore. ( Idle --> Attesa POS )
            Dim _CallGetInfo As StatusCode = Me.CallInfoOnPosHardware(funcName)

            ' Una Volta richiamata la demateriliazzione 
            ' ed eventuale conferma ed hanno dato esito positivo.
            If _CallGetInfo = StatusCode.OK Then

                ' Signal sul Form OK
                _SetOperationStatusForm(
                    NOT_INFO_POS_DATA_OK,
                    "POS data ok!!",
                    FormEmulationArgentea.InfoStatus.OK, 9
                )
                System.Threading.Thread.Sleep(100)
                'System.Windows.Forms.Application.DoEvents()

                '
                '   Riprendo il Risultato e lo riporto
                '   come attuato e verificato dal POS Hardware.
                '
                CalculateAndFillMultiItemsInfoTicket(m_Transaction_Identifier)

                ' Cancelliamo Attesa precedente
                _ClearWaitScreen()

            Else

                ' Errata Chiamata Info della Card Utente
                ' data dalla risposta argentea quindi su POS.
                _ClearWaitScreen()

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                _SetOperationStatus(
                    funcName,
                    NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                    m_LastErrorMessage,
                    PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                )

                InError = True
                Return

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

            '
            ' Da Riportare al chiamante
            '
            InError = True

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Signal Initialization Status
            System.Threading.Thread.Sleep(200)
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla command che riceve il click per collegarsi
    '''     al pos hardware che chiederà di recuperare l'ultima transazione effettuata dell'utente.
    ''' </summary>
    ''' <param name="sender">Il controllo command</param>
    Private Sub ConnectHardwareReloadLastTransactionHandler(ByRef sender As Object, e As EventArgs, ByRef InError As Boolean)
        Dim funcName As String = "ConnectHardwareReloadLastTransactionHandler"

        InError = True

    End Sub


#End Region

#Region "** SERVICE LOCALE POS FORM -> Con i suoi Handler per il Collect dalle azioni del FORM Locale **"

    ''' <summary>
    '''     Avvia mostrando il Form per il servizio
    '''     POS locale che attende le scansioni dei
    '''     barcode provenienti dall'operatore.
    ''' </summary>
    Private Sub StartPosSoftware()
        Dim funcName As String = "StartPosSoftware"

        ' Entrap sull'idle
        Try

            '
            ' Resetto i contatori iniziali
            ' per l'aggiornamento di stato.
            '
            m_TotalPayed_CS = 0
            m_TotalInvalid_CS = 0
            m_TotalVoided_CS = 0
            m_TotalInvalid_CS = 0

            '
            ' 1° STep ( Preparazione Form ed eventi sui controlli ) 
            '       Creo l'istanza adatta al tipo di azione di  un form 
            '       in emulazione POS Sofwware per il Service.
            '
            ' Prepara avviando a seconda della funzione per il Void o per il Pay
            '       --> _SetFormForUseServiceVoid(FormEmulationArgentea.enScopeUseForm.VOID)
            '           '*
            '           --> AddHandler frmEmulation.BarcodeRead, AddressOf BarcodeReadVoidHandler
            '           --> AddHandler frmEmulation.BarcodeRemove, AddressOf BarcodeRemoveVoidHandler
            '           '*
            '       --> _SetFormForUseServicePayment(FormEmulationArgentea.enScopeUseForm.DEMAT)
            '           '*
            '           --> AddHandler frmEmulation.BarcodeRead, AddressOf BarcodeReadVoidHandler
            '           --> AddHandler frmEmulation.BarcodeRemove, AddressOf BarcodeRemoveVoidHandler
            '           '*
            '       *.*--> AddHandler frmEmulation.BarcodeCheck, AddressOf BarcodeCheckVoidHandler
            '       *.*--> AddHandler frmEmulation.FormClosed, AddressOf CloseOperationHandler ( Completa i Totali)
            CreateInstanceFormForEmulationOrWaitAction()

            ' Aggiorno il DataResponse e la vista (Iniziale)
            _UpdateResultData("INITIALIZE_EMULATOR_SOFTWARE", m_CurrentAmountScalable, Nothing)

            ' Status
            m_ServiceStatus = enProxyStatus.InRunning

            ' In questa modalità avvio il waitwscreen
            ' modale a pieno schermo per  attendere 
            ' le operazioni dal Pos Locale collegato.
            _ShowWaitScreen(3, "Carico Terminale Locale", "Attesa azione utente")

            '
            '   2° STep ( Controllo se in VOID e prefillo sessione precedente)
            '           Aggiusta l'importo scalabile e riporta i dati iniziali.
            '
            If Not m_PrefillVoidable Is Nothing Then

                Me.PrefillVoidable = m_PrefillVoidable

                ' Signal Initialization Status
                _SetOperationStatusForm(
                        NOT_INFO_POS_INIT,
                        "Refill Data...",
                        FormEmulationArgentea.InfoStatus.Wait
                    )

                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(200)

                ' Carico i Dati Inziali Prefillati dalla chiamata
                Me.CalculateAndFillMultiItemsInitialize()

                'For Each itm As PaidEntry In m_PrefillVoidable.Values
                ' Aggiungo l'elemento al controllo Griglia
                'frmEmulation.AddItemOnGrid(itm)
                'Next

            End If

            ' Tolgo eventuali Wait Screen
            _ClearWaitScreen()

            '********************************************************
            '
            ' (Idle) --> senza AutoClose (ReadWrite)
            ShowAndIdleOnFormForAction(False)
            '
            '********************************************************

            ' Dichiaro come concluso correttamente tutto
            If m_ServiceStatus = enProxyStatus.InRunning Then

                ' Se era rimasto in Running e non InError
                ' tutto è filato liscio e torno con OK
                m_ServiceStatus = enProxyStatus.OK

                ' Signal Finish Status
                _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

            Else

                ' Signal (come errore di mancato collegamento all'hardware pos ma non bloccante)
                _SetOperationStatus(
                    funcName,
                    NOT_ERROR_SOFTWAREPOS_FAILED,
                    "Emulatore POS non avviato in modo corretto, controllare le impostazioni e lo stato!",
                    PosDef.TARMessageTypes.TPERROR
                )

                ' Scrive una riga di Log per monitorare....
                LOG_Info(getLocationString(funcName), m_ServiceStatus)

            End If

        Catch ex As Exception

            ' Signal Initialization Status
            _SetOperationStatusForm(
                NOT_INFO_POS_ERROR,
                "POS in error...",
                FormEmulationArgentea.InfoStatus.Error
            )

            '
            ' Alla chiusura della chiamata Errori possibili
            ' dal Proxy Argentea Gestiti come SOCKET o CHIAMATA 
            ' o interni al flow corrente come Errori non Gestiti Uknowed
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

            ' Con Ritorno a Status InError
            m_ServiceStatus = enProxyStatus.InError

        Finally

            Try

                '
                '  Provo a chiudere il Form del POS
                ' software se siamo in questa modalità.
                '
                If Not frmEmulation Is Nothing Then
                    m_TheModcntr.DialogActiv = False
                    m_TheModcntr.DialogFormName = ""
                    m_TheModcntr.SetFuncKeys((True))
                    m_TheModcntr.EndForm()
                    frmEmulation.Close()
                    frmEmulation = Nothing
                End If
                '
            Catch ex As Exception

                '
                ' Errori Interni UKNOWED da rimarcare
                '
                SetExceptionsStatus(funcName, ex, m_SilentMode)

                ' Con Ritorno a Status InError
                m_ServiceStatus = enProxyStatus.InError

            Finally

                ' Chiudo per memory leak con argentea
                ArgenteaCOMObject = Nothing

                ' Effettuo un Dispose forzato per 
                ' la chiusura del form su eccezioni.
                If Not frmEmulation Is Nothing Then
                    frmEmulation.Dispose()
                    frmEmulation = Nothing
                End If

            End Try

        End Try
    End Sub

    ''' <summary>
    '''     Handler sulla textbox che riceve in input i barcode in ingresso
    '''     per le azioni di interrogazione di Stato dei buoni pasto cartacei.
    ''' </summary>
    ''' <param name="sender">Il controllo textbox</param>
    ''' <param name="barcode">Il barcode stringato nell'evento come parametro di handler</param>
    Protected Overridable Sub BarcodeCheckHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeCheckHandler"
        Dim Inizializated As Boolean = True
        Dim faceValue As Decimal = 0
        Dim paidValue As Decimal = 0

        '_
        '
        '    Tipi di Repsonse su Protoccolo Argentea
        '    
        '    ->  OPEN TICKET (CallInitialization)       :::         OK--TICKET APERTO-----0--- 
        '    ->  OK su CHK BPC (CallDematerialize)      :::         OK-0 - BUONO VALIDO -68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202-- 
        '    ->  KO su CHK BPC (CallDematerialize)      :::         KO-903 - PROGRESSIVO FUORI SEQUENZA-----0--- 
        '_
        '

        Try

            If TypeOf sender Is FormEmulationArgentea Then

                ' Signal Initialization Status
                _SetOperationStatusForm(
                    NOT_INFO_POS_CALL,
                    "Interrogazione in corso...",
                    FormEmulationArgentea.InfoStatus.Wait
                )

                ' Catturiamo subito il Barcode
                m_CurrentBarcodeScan = barcode

                '
                ' La prima chiamata Apre la sessione su host remoto Argentea
                '
                If Not m_FirstCall Then

                    ' Chiama per  la  Dematirializzazione
                    ' e incrementa di uno il numero delle
                    ' chiamate interne.
                    _ShowWaitScreen(2, "Inizializzazione", "Inizizializzazione Argentea")
                    Inizializated = Me.CallInitialization(funcName)
                    m_FirstCall = True

                End If

                _ClearWaitScreen()

                If Inizializated Then

                    ' Mostriamo il Wait
                    _ShowWaitScreen(2, "Operazione in corso", "Check in corso presso il Provider Argentea")

                    ' Chiama per  l'Interrogazione di stato
                    ' di un determinato buono pasto da
                    ' controlllare.
                    Dim _CallCheckTicket As StatusCode = Me.CallCheckTicket(funcName)
                    Dim _CallConfirmation As StatusCode = StatusCode.OK

                    ' Vediamo se richiesta la conferma su Demat
                    If _CallCheckTicket = StatusCode.CONFIRMREQUEST Then

                        ' Chiama per  la  Risposta di stato del
                        ' BP al forma princiaple
                        _CallConfirmation = Me.CallConfirmOperation(funcName, "interrogation")

                        If _CallConfirmation = StatusCode.KO Then

                            ' Log locale (Non confermato in demat)
                            LOG_Info(funcName, "Transaction Check Status on Argentea ::KO:: ON CONFIRM")

                        End If

                    End If

                    ' Una Volta richiamata la demateriliazzione 
                    ' ed eventuale conferma ed hanno dato esito positivo.
                    If _CallCheckTicket = StatusCode.OK And _CallConfirmation = StatusCode.OK Then

                        ' Signal sul Form OK
                        _SetOperationStatusForm(
                            NOT_INFO_POS_DATA_OK,
                            "POS data ok!!",
                            FormEmulationArgentea.InfoStatus.OK, 9
                        )
                        System.Threading.Thread.Sleep(100)
                        System.Windows.Forms.Application.DoEvents()

                        ' Riprendo per il BP i valori
                        ' e il pagato reale
                        faceValue = m_CurrentValueOfBP
                        paidValue = m_CurrentValueOfBP + m_PartialValueExcedeed_XS

                        ' Aggiungo in una collection specifica in uso
                        ' interno l'elemento Buono appena accodato in
                        ' modo univoco rispetto al suo BarCode.
                        Dim ItemNew As PaidEntry = New PaidEntry(
                            m_CurrentBarcodeScan,
                            paidValue.ToString("###,##0.00"),
                            faceValue.ToString("###,##0.00"),
                            m_LastResponseRawArgentea.Provider
                        )

                        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form
                        _UpdateResultData("SINGLEP_CHECK", _PartialTransactValue, ItemNew)

                        ' Cancelliamo Attesa precedente
                        _ClearWaitScreen()

                    Else

                        ' Errato Contorllo stato Buono o Confirm su stato Buono
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        _ClearWaitScreen()

                        ' Signal Error Ritardato Status
                        _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                        ' Signal (KO remoto su controllo + Status remoto per le codifiche da db personalizzate)
                        _SetOperationStatus(
                            funcName,
                            NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                            m_LastErrorMessage,
                            PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                        )

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    _ClearWaitScreen()

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction CheckBP Argentea ::KO:: Remote")

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                        m_LastErrorMessage,
                        PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                    )

                End If

            Else

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Riporto la firstcall a false
            ' per le istanze successive.
            m_FirstCall = False

            ' Svuoto il controllo del barcode
            'If Not formBC Is Nothing Then formBC.txtBarcode.Text = String.Empty

            ' Signal Initialization Status
            System.Threading.Thread.Sleep(200)
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

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

        '_
        '
        '    Tipi di Repsonse su Protoccolo Argentea
        '    
        '    ->  OPEN TICKET (CallInitialization)       :::         OK--TICKET APERTO-----0--- 
        '    ->  OK su ADD BPC (CallDematerialize)      :::         OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202-- 
        '    ->  KO su ADD BPC (CallDematerialize)      :::         KO-903 - PROGRESSIVO FUORI SEQUENZA-----0--- 
        '_
        '

        Try

            If TypeOf sender Is FormEmulationArgentea Then

                ' Signal Initialization Status
                _SetOperationStatusForm(
                    NOT_INFO_POS_CALL,
                    "Interrogazione in corso...",
                    FormEmulationArgentea.InfoStatus.Wait
                )

                ' Catturiamo subito il Barcode
                m_CurrentBarcodeScan = barcode

                '
                ' ::Opzione:: Max BP pagabili per vendita.: 
                '
                '       Controllo che il numero di BP in totale 
                '       globali alla Vendita corrente rispetto a
                '       l'opzione di utilizzo sia non superato.     
                '
                If m_OPT_MaxNumBPSomeSession <> 0 And (((WriterResultDataList.Count + 1) > (m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc)) Or ((m_TotalBPUsedToPay_CS + 1) > (m_OPT_MaxNumBPSomeSession - m_TotalBPInUseOnCurDoc))) Then

                    ' Signal (Numero buoni dematerializzabili superato rispetto a quelli permessi)
                    _SetOperationStatus(funcName,
                        NOT_OPT_ERROR_NUMBP_EXCEDEED,
                        "Il numero di titoli di pagamento per questa vendita è stato superato!!",
                        PosDef.TARMessageTypes.TPINFORMATION, True
                    )
                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    Return

                End If

                ' Controllo se nell'elenco è già stato considerato il BarCode
                If _DataResponse.ContainsBarcode(m_CurrentBarcodeScan) Then

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (Buono già usato in questo pagamento per questa vendita)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_CODE_ALREADYINUSE,
                        "Il barcode per questo titolo di pagamento è già stato usato per questa vendita!!",
                        PosDef.TARMessageTypes.TPINFORMATION, True
                    )

                    Return

                End If

                '''     Controllo per il Buono in Corso se il valore
                '''     supera l'importo pagabile, e se lo è  allora
                '''     riporto lo status del pagamento generale come
                '''     in eccesso (sempre se l'opzione lo permette)
                '''     e se lo permette riporto per l'ultimo buono in 
                '''     corso il totale di facciata diverso dal valore effettivo
                '''     di pagato (per scrivefe un media di resto all'uscita)
                If m_PartialValueExcedeed_XS <> 0 Then

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (Importo già raggiunto)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_IMPORT_ALREADYCOMPLETED,
                        "L'importo da pagare è già stato completato per questa vendita completare con inoltro!!",
                        PosDef.TARMessageTypes.TPINFORMATION, True
                    )

                    Return

                End If

                '
                ' La prima chiamata Apre la sessione su host remoto Argentea
                '
                If Not m_FirstCall Then

                    ' Chiama per  la  Dematirializzazione
                    ' e incrementa di uno il numero delle
                    ' chiamate interne.
                    _ShowWaitScreen(2, "Inizializzazione", "Inizizializzazione Argentea")
                    Inizializated = Me.CallInitialization(funcName)
                    m_FirstCall = True

                End If

                _ClearWaitScreen()

                If Inizializated Then

                    ' Mostriamo il Wait
                    _ShowWaitScreen(2, "Operazione in corso", "Demat in corso presso il Provider Argentea")

                    ' Chiama per  la  Dematirializzazione
                    ' e incrementa di uno il numero delle
                    ' chiamate interne.
                    Dim _CallDematerialize As StatusCode = Me.CallDematerialize(funcName)
                    Dim _CallConfirmation As StatusCode = StatusCode.OK

                    ' Vediamo se richiesta la conferma su Demat
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
                    If _CallDematerialize = StatusCode.OK And _CallConfirmation = StatusCode.OK Then

                        ' Signal sul Form OK
                        _SetOperationStatusForm(
                            NOT_INFO_POS_DATA_OK,
                            "POS data ok!!",
                            FormEmulationArgentea.InfoStatus.OK, 9
                        )
                        System.Threading.Thread.Sleep(100)
                        System.Windows.Forms.Application.DoEvents()

                        ' Riprendo per il BP i valori
                        ' e il pagato reale
                        faceValue = m_CurrentValueOfBP
                        paidValue = m_CurrentValueOfBP + m_PartialValueExcedeed_XS

                        ' Aggiungo in una collection specifica in uso
                        ' interno l'elemento Buono appena accodato in
                        ' modo univoco rispetto al suo BarCode.
                        Dim ItemNew As PaidEntry = WriterResultDataList.NewPaid(
                            m_CurrentBarcodeScan,
                            Value:=paidValue.ToString("###,##0.00"),
                            FaceValue:=faceValue.ToString("###,##0.00"),
                            Emitter:=m_LastResponseRawArgentea.Provider,
                            CodeIssuer:=m_LastResponseRawArgentea.CodeIssuer,
                            NameIssuer:=m_LastResponseRawArgentea.NameIssuer,
                            IdTransactionCrc:=m_LastResponseRawArgentea.TerminalID
                        )

                        ' Per l'azione sull'elenco corrente mi
                        ' riprendo il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        m_PartialPayed_XS += m_CurrentValueOfBP
                        m_PartialBPUsedToPay_XS += 1                         ' <-- Conteggio numero di bpc usati in local per ogni ingresso sulla vendita
                        m_TotalBPElaborated_CS += 1
                        _PartialTransactValue += paidValue

                        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form
                        _UpdateResultData("SINGLEP_UPDATE", _PartialTransactValue, ItemNew)

                        ' Cancelliamo Attesa precedente
                        _ClearWaitScreen()

                        ' Non Abbiamo finito, dobbiamo ancora vedere
                        ' se abbiamo superato l'importo da pagare  e
                        ' se accettata l'opzione di avere un resto o 
                        ' meno rispetto al pagato con buoni.
                        If Not CheckIfDematExcedeedTotToPay(sender) Then
                            Return
                        End If

                    Else

                        ' Errata Dematerializzione o Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        _ClearWaitScreen()

                        ' Signal Error Ritardato Status
                        _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                        ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                        _SetOperationStatus(
                            funcName,
                            NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                            m_LastErrorMessage,
                            PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                        )

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    _ClearWaitScreen()

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Remote")

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                        m_LastErrorMessage,
                        PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                    )

                End If

            Else

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Riporto la firstcall a false
            ' per le istanze successive.
            m_FirstCall = False

            ' Svuoto il controllo del barcode
            'If Not formBC Is Nothing Then formBC.txtBarcode.Text = String.Empty

            ' Signal Initialization Status
            System.Threading.Thread.Sleep(200)
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

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

            If TypeOf sender Is FormEmulationArgentea Then

                ' Signal Initialization Status
                _SetOperationStatusForm(
                    NOT_INFO_POS_CALL,
                    "Interrogazione in corso...",
                    FormEmulationArgentea.InfoStatus.Wait
                )

                ' Cattueriamo subito il Barcode
                m_CurrentBarcodeScan = barcode

                'Controllo se nell'elenco è già stato considerato il BarCode ' Not m_FlagUndoBPCForExcedeed And
                If Not _DataResponse.ContainsBarcode(m_CurrentBarcodeScan, True) Then

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (Titolo non presente tra quelli usati)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_CODE_NOTPRESENT,
                        "Il titolo non è presente in elenco tra quelli usati!! ",
                        PosDef.TARMessageTypes.TPINFORMATION, True
                    )
                    Return

                End If

                If Inizializated Then

                    _ShowWaitScreen(2, "Operazione in Corso", "Annullo di un titolo già dematerializzato in corso")

                    ' Chiama per l'anullamento di uno già 
                    ' Dematirializzato  e  incrementa  di 
                    ' uno il numero delle il numero delle 
                    ' chiamate interne.
                    Dim _IDSpecifiqueTransaction As String = _DataResponse.GetIDTransactionOfThisBarCode(m_CurrentBarcodeScan)
                    Dim _CallReverseMaterializated As StatusCode = Me.CallReverseMaterializated(funcName, _IDSpecifiqueTransaction)
                    Dim _CallConfirmation As StatusCode = StatusCode.OK

                    If _CallReverseMaterializated = StatusCode.CONFIRMREQUEST Then

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
                    If _CallReverseMaterializated <> StatusCode.KO And _CallConfirmation = StatusCode.OK Then

                        ' Signal sul Form OK
                        _SetOperationStatusForm(
                            NOT_INFO_POS_DATA_OK,
                            "Operazione completata!!",
                            FormEmulationArgentea.InfoStatus.OK, 9
                        )
                        System.Threading.Thread.Sleep(200)

                        ' Rimuovo dalla collection specifica in uso
                        ' interno l'elemento Buono da annullare individuandolo
                        ' in modo univoco rispetto al suo BarCode con cui era 
                        ' stato registrato all'aggiunta dell'handler di ADD.
                        For Each itm As PaidEntry In WriterResultDataList
                            If itm.Barcode = m_CurrentBarcodeScan Then
                                'WriterResultDataList.Remove(itm)
                                itm.Voided = True
                                Exit For
                            End If
                        Next

                        ' Argormento Opzione per Opzione 
                        ' su Flow operatore se non accetta
                        ' Sulla griglia e il form non deve
                        ' fare altro dato che non è stato aggiunto.
                        If m_FlagUndoBPCForExcedeed Then
                            Return
                        End If

                        ' Aggiornamento paraziali
                        m_TotalBPElaborated_CS += 1
                        m_PartialBPUsedToVoid_XS += 1

                        ' Per il Form in azione corrente mi
                        ' aggiorno il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        faceValue = faceValue
                        paidValue = m_CurrentValueOfBP
                        m_PartialVoided_XS += paidValue
                        _PartialTransactValue -= paidValue
                        m_PartialValueExcedeed_XS = Math.Min((m_CurrentAmountScalable - (m_PartialPayed_XS - m_PartialVoided_XS)), 0) ' - m_CurrentValueOfBP

                        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form
                        _UpdateResultData("SINGLEP_REMOVE_SELECTED", _PartialTransactValue, Nothing)

                    Else

                        ' Errata Reverse per Dematerializzione o Reverse Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        _ClearWaitScreen()

                        ' Log locale
                        LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Reverse Demat Argentea ::KO:: Local")

                        ' Signal Error Ritardato Status
                        _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                        ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                        _SetOperationStatus(
                            funcName,
                            NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                            m_LastErrorMessage,
                            PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                        )

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    _ClearWaitScreen()

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Reverse Demat Argentea ::KO:: Not Intializated")

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                        m_LastErrorMessage,
                        PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                    )

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Svuoto il controllo del barcode
            'If Not m_FlagUndoBPCForExcedeed And Not formBC Is Nothing Then formBC.txtBarcode.Text = String.Empty

            ' Signal Initialization Status
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla gridbox che riceve in input il barcode selezionato
    '''     per le azioni di controllo di validità del buono pasto cartaceo o coupon.
    ''' </summary>
    ''' <param name="sender">Istanza del form sugll'handler degli eventi</param>
    ''' <param name="barcode">Il Barcode all'evneto sulla scansione</param>
    Protected Overridable Sub BarcodeCheckValidCodeHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeCheckValidCodeHandler"
        Dim Inizializated As Boolean = True
        Dim faceValue As Decimal = 0
        Dim paidValue As Decimal = 0

        '_
        '
        '    Tipi di Repsonse su Protoccolo Argentea
        '    
        '    ->  OPEN TICKET (CallInitialization)       :::         OK--TICKET APERTO-----0--- 
        '    ->  OK su CHK BPC (CallCheckValid)         :::         Buono-000-Buono Valido-529-8897456-12345687-201809201733577-ARGENTEA-
        '    ->  KO su CHK BPC (CallCheckValid)         :::         Buono-2 - Buono non Valido già utilizzato-------- 
        '    ->  KO su CHK BPC (CallCheckValid)         :::         Buono-903 - PROGRESSIVO FUORI SEQUENZA-------- 
        '_
        '

        Try

            If TypeOf sender Is FormEmulationArgentea Then

                ' Signal Initialization Status
                _SetOperationStatusForm(
                    NOT_INFO_POS_CALL,
                    "Interrogazione in corso...",
                    FormEmulationArgentea.InfoStatus.Wait
                )

                ' Cattueriamo subito il Barcode
                m_CurrentBarcodeScan = barcode

                If Inizializated Then

                    _ShowWaitScreen(2, "Operazione in Corso", "Controllo di un titolo di pagamento valido in corso")

                    ' Chiama per il  controllo di un buono
                    ' o Coupon dal suo codice e incrementa  
                    ' di uno il numero delle chiamate interne. 
                    Dim _CallCheckValidItem As StatusCode = Me.CallCheckTicket(funcName)
                    Dim _CallConfirmation As StatusCode = StatusCode.OK

                    If _CallCheckValidItem = StatusCode.CONFIRMREQUEST Then

                        ' Chiama  per  conferma  di Controllo sul
                        ' codice per questo titolo  di  pagamento
                        ' e  incrementa di uno il numero delle chiamate interne.
                        _CallConfirmation = Me.CallConfirmOperation(funcName, "check")

                        If _CallConfirmation = StatusCode.KO Then

                            ' Log locale (Non confermato in check per controllo validità)
                            LOG_Info(funcName, "Transaction for Check Valid BP on Argentea ::KO:: ON CONFIRM")

                        End If

                    End If

                    ' Una Volta richiamata l'api di controllo del titolo
                    ' ed eventuale conferma ed hanno dato esito positivo
                    ' In Questo specifico caso il Check può asumere OK o KO 
                    If _CallConfirmation = StatusCode.OK Then

                        ' Signal sul Form OK
                        _SetOperationStatusForm(
                            NOT_INFO_POS_DATA_OK,
                            "Operazione completata!!",
                            FormEmulationArgentea.InfoStatus.OK, 9
                        )
                        System.Threading.Thread.Sleep(200)

                        ' Per il Form in azione corrente mi
                        ' aggiorno il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        faceValue = m_CurrentValueOfBP 'faceValue
                        paidValue = m_CurrentValueOfBP
                        Dim ElementChecked As PaidEntry = New PaidEntry(barcode, paidValue, faceValue, m_LastResponseRawArgentea.Type, m_LastResponseRawArgentea.TerminalID)
                        ElementChecked.InfoExtra = m_LastResponseRawArgentea.Type

                        If _CallCheckValidItem = StatusCode.OK Then

                            ' Aggiorno La vista con il ritultato su OK del BPC
                            _UpdateResultData("SINGLEP_CHECKED_VALID", faceValue, ElementChecked)

                        Else ' KO

                            ' Aggiorno La vista con il ritultato su KO del BPC
                            _UpdateResultData("SINGLEP_CHECKED_INVALID", faceValue, ElementChecked)

                        End If

                    Else

                        ' Errata Interrogazione per Controllo di Validità del Titolo al Confirm su Check
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        _ClearWaitScreen()

                        ' Log locale
                        LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Check BP Confirm Argentea ::KO:: Local")

                        ' Signal Error Ritardato Status
                        _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                        ' Signal (KO remoto su check valid bp + Status remoto per le codifiche da db personalizzate)
                        _SetOperationStatus(
                            funcName,
                            NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                            m_LastErrorMessage,
                            PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                        )

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    _ClearWaitScreen()

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Check Valid BP Argentea ::KO:: Not Intializated")

                    ' Signal Error Ritardato Status
                    _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                    ' Signal (KO remoto su check valid bp + Status remoto per le codifiche da db personalizzate)
                    _SetOperationStatus(
                        funcName,
                        NOT_INFO_OPERATION_NOT_VALID_SPECIAL,
                        m_LastErrorMessage,
                        PosDef.TARMessageTypes.TPERROR, True  ' <-- Lo status remoto
                    )

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Chiudo sempre il waitscreen
            _ClearWaitScreen()

            '
            ' Errori Interni UKNOWED da rimarcare
            '
            SetExceptionsStatus(funcName, ex, m_SilentMode)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            _ClearWaitScreen()

            ' Svuoto il controllo del barcode
            'If Not m_FlagUndoBPCForExcedeed And Not formBC Is Nothing Then formBC.txtBarcode.Text = String.Empty

            ' Signal Initialization Status
            _SetOperationStatusForm(String.Empty, String.Empty, FormEmulationArgentea.InfoStatus.Flush)

        End Try

    End Sub

    ''' <summary>
    '''     Handler sulla gridbox che riceve in input il barcode selezionato
    '''     per le azioni di storno del buono pasto cartaceo.
    ''' </summary>
    ''' <param name="sender">Istanza del form sugll'handler degli eventi</param>
    ''' <param name="barcode">Il Barcode all'evneto sulla scansione</param>
    Protected Overridable Sub BarcodeReadVoidHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeReadVoidHandler"

        ' Immediatamente annullo verso il sistema argnetea l'operazione
        ' Per rimuoverlo tramite il metodo stesso per l'annullo
        m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
        Me.BarcodeRemoveHandler(sender, barcode)
        m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

        If m_LastStatus Is Nothing Then

            ' Se filato liscio e la void è stata effetuata
            _updateVoidedForm(sender)

        End If

    End Sub

    ''' <summary>
    '''     Handler sulla gridbox che riceve in input la riga con barcode selezionato
    '''     per le azioni di storno del buono pasto cartaceo.
    ''' </summary>
    ''' <param name="sender">Istanza del form sugll'handler degli eventi</param>
    ''' <param name="barcode">Il Barcode all'evneto sulla scansione</param>
    Protected Overridable Sub BarcodeRemoveVoidHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeRemoveVoidHandler"

        ' Immediatamente annullo verso il sistema argnetea l'operazione
        ' Per rimuoverlo tramite il metodo stesso per l'annullo
        m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
        Me.BarcodeRemoveHandler(sender, barcode)
        m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

        If m_LastStatus Is Nothing Then

            ' Se filato liscio e la void è stata effetuata
            m_PartialBPUsedToVoid_XS += 1           ' lo aggiorno qui
            _updateVoidedForm(sender)

        End If

    End Sub

    ''' <summary>
    '''     Completa l'aggiornamento dei dati sul form
    '''     se la void è andata a buon fine.
    '''     ( Usato per appoggio a BarcodeRemoveHandler flaggato per uscire prima)
    ''' </summary>
    ''' <param name="sender">Il form di storno</param>
    Private Sub _updateVoidedForm(ByRef sender As Object)
        Dim funcName As String = "_updateVoidedForm"
        Dim _revoke As Boolean = False

        '
        ' Rimuovo dalla collection specifica in uso
        ' interno l'elemento Buono da annullare individuandolo
        ' in modo univoco rispetto al suo BarCode con cui era 
        ' stato registrato all'aggiunta dell'handler di ADD.
        '
        For Each itm As PaidEntry In WriterResultDataList
            If itm.Barcode = m_CurrentBarcodeScan Then
                itm.Voided = True  ' Etitchettato come --> VOIDED -> Stornato
                _revoke = True
                Exit For
            End If
        Next

        If Not _revoke Then

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Signal (Titolo non presente in elenco)
            _SetOperationStatus(
                funcName,
                NOT_INFO_CODE_NOTPRESENT,
                "Il titolo non è presente in elenco tra quelli utilizzzati!!",
                PosDef.TARMessageTypes.TPINFORMATION, True
            )

            Return

        End If

        ' Se è stato rimosso correttametne procediamo
        m_PartialBPUsedToPay_XS -= 1                        ' <-- Conteggio numero di bpc usati in local per ogni rimosso

        ' Per il Form in azione corrente mi
        ' aggiorno il Totale da Pagare rispetto a
        ' quelli già in elenco
        m_VoidableAmount -= m_CurrentValueOfBP
        m_PartialVoided_XS += m_CurrentValueOfBP

        ' Aggiorno i Dati iniziali per il ResulData e la Vista sul Form (m_CurrentBarcodeScan)
        _UpdateResultData("SINGLEP_REMOVE_BARCODE", _PartialTransactValue, Nothing)

    End Sub


    Private Function CheckIfDematExcedeedTotToPay(sender As Object) As Boolean
        Dim funcName As String = "CheckIfDematExcedeedTotToPay"

        '
        ' ::Opzione:: Operatività.: 
        '       Se il Totale in ingresso è minore rispetto 
        '       al valore di facciata del Buono Pasto una volta
        '       ottenuto dalla materializzazione, opto per troncare su totale.
        Dim OptAcceptExceeded As Boolean = st_Parameters_Argentea.BP_AcceptExcedeedValues

        ' Mi conteggio eventuali eccessi su pagato
        m_PartialValueExcedeed_XS = 0

        'SU OK
        m_PartialValueExcedeed_XS = Math.Min((m_CurrentAmountScalable - m_PartialPayed_XS), 0) ' - m_CurrentValueOfBP

        'Su Opzione accetta Valore in eccesso per resto
        If OptAcceptExceeded Then

            '
            '       --> Accetta eccesso su Totale da Pagare
            '               Alla fine scrive il media don i due riporti
            '               concludendo il pagamento a totale.
            '
            If m_PartialValueExcedeed_XS < 0 Then

                ' Log locale
                LOG_Debug(funcName, m_LastStatus + " - Excedeed Rest -- Transaction Dematerialize Argentea ::OK:: Excedeed")

                ' Mostro il Tasto per stampare lo scontrino
                frmEmulation.EnableOperationToPrintReceiptOnExcedeed = True

            Else

                ' Mostro la label Rimanenza normale
                frmEmulation.EnableOperationToPrintReceiptOnExcedeed = True

            End If

            Return True

        Else

            '
            '       --> Non Accetta eccesso su Totale da Pagare
            '               Richiama Argentea per fare l'annullo
            '               alla demateriliazzazione fatta in precedenza
            '

            If m_PartialValueExcedeed_XS < 0 Then

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "-- Transaction Dematerialize Argentea ::KO:: Excedeed")

                ' Signal Error Ritardato Status
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

                ' Signal (Eccedenza sul totale)
                _SetOperationStatus(
                    funcName,
                    NOT_OPT_ERROR_VALUE_EXCEDEED,
                    "Il Valore del Titolo di Pagamento eccede il valore rispetto al totale (non è possibile dare resto)!!",
                    PosDef.TARMessageTypes.TPERROR, True
                )

                ' Immediatamente annullo verso il sistema argnetea l'operazione
                ' Per rimuoverlo tramite il metodo stesso per l'annullo
                m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
                Me.BarcodeRemoveHandler(sender, m_CurrentBarcodeScan)
                m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

                ' Ripristiniamo l'importo d'eccedenza
                m_PartialValueExcedeed_XS = 0

                ' Torno all'inseirmento eventualemnete per optare su altri 
                ' buoni pasto cliente con importi appropriati al completamento.
                Return False

            Else

                Return True

            End If

        End If

    End Function

#End Region

#Region "** VISTA LOCALE POS FORM -> Funzioni private per aggiornare il form e completare lo stato del ResultData"

    ''' <summary>
    '''     Aggiorna il set di rusltati  pronti
    '''     per restituirli al Chiamante, e sul
    '''     form di visualizzazione.
    ''' </summary>
    Private Sub _UpdateResultData(Mode As String, CurrentTransact As Decimal, ItemData As PaidEntry)

        Dim _UpdateFormEmulator As Boolean = False
        Dim _UpdateTotals As Boolean = False

        Try

            Select Case Mode

                Case "INITIALIZE"

                    '
                    ' Crea l'Istanza della Risposta da
                    ' dare in Uscita a completamento.
                    '
                    _DataResponse = New DataResponse(
                        CurrentTransact,                '   Il Totale sulla TA prima di entrare
                        m_PartialBPUsedToPay_XS,        '   N° elaborati
                        m_PartialPayed_XS,              '   Pagato
                        m_PartialValueExcedeed_XS,      '   L'eccesso usato come resto in Pay
                        m_PartialBPUsedToVoid_XS,       '   Il Numero di Elementi Elaborati come Voci di Storno
                        m_PartialVoided_XS,             '   L'Importo Stornato fino adesso
                        m_PartialBPNotValid_XS,         '   Il Numero di Elementi Elaborati come Voci non Valide (BPE)
                        m_PartialInvalid_XS             '   L'Importo Scartato fino adesso (Non Validi scartati in BPE)
                    )

                    _UpdateFormEmulator = True

                Case "FILL_DATA_RESPONSE_INIT"
                    ' Data Response

                    ' Blocco l'importo scalabile per aggiornamento
                    frmEmulation.LockAmountScalable = True
                    frmEmulation.Payable = m_CurrentAmountScalable
                    frmEmulation.InitialAmountScalableOnSession = 0

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "FILL_DATA_RESPONSE_UPDATE"
                    ' Data Response

                    ' Per l'Hardware Mode Riempio l'elenco
                    frmEmulation.AddItemBPE(ItemData)

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "FILL_DATA_VOID_RESPONSE_INIT"
                    ' Void

                    ' Blocco l'importo scalabile per aggiornamento
                    frmEmulation.LockAmountScalable = True
                    frmEmulation.Payable = m_CurrentAmountScalable

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "FILL_DATA_VOID_RESPONSE_UPDATE"
                    ' Void

                    ' Per l'Hardware Mode Riempio l'elenco
                    frmEmulation.AddItemBPE(ItemData)

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "PREFILL_INIT"

                    ' Resetto i valori iniziali per la response.
                    _DataResponse.UpdateInitialValues(
                            CurrentTransact,          '   Il Totale sulla TA prima di entrare
                            0,          '   N° elaborati
                            0,          '   Pagato
                            0,          '   L'eccesso usato come resto in Pay
                            0,
                            0,          '   Stornato
                            0,
                            0           '   Non Validi scartati in BPE
                        )

                    ' Blocco l'importo scalabile per aggiornamento
                    frmEmulation.LockAmountScalable = True
                    frmEmulation.Payable = CurrentTransact

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "PREFILL_UPDATE"

                    ' Aggiorno quelli per la response iniziali
                    _DataResponse.UpdateInitialValues(
                            CurrentTransact,          '   Il Totale sulla TA prima di entrare
                            m_PartialBPUsedToPay_XS,        '   N° elaborati
                            m_PartialPayed_XS,              '   Pagato
                            m_PartialValueExcedeed_XS,      '   L'eccesso usato come resto in Pay
                            m_PartialBPUsedToVoid_XS,
                            m_PartialVoided_XS,             '   Stornato
                            m_PartialBPNotValid_XS,
                            m_PartialInvalid_XS             '   Non Validi scartati in BPE
                        )

                    If m_TypeProxy = enTypeProxy.Service Then

                        ' Aggiungo l'elemento al controllo Lista
                        frmEmulation.AddItemOnGrid(ItemData)

                    Else

                        ' Aggiungo l'elemento al controllo Griglia
                        frmEmulation.AddItemBPE(ItemData, True)

                    End If

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "PREFILL_END"

                    ' Blocco l'importo scalabile per aggiornamento
                    frmEmulation.LockAmountScalable = True
                    frmEmulation.Payable = CurrentTransact

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True


                Case "FILL_DATA_INFO_RESPONSE_INIT"

                    ' Ripulisco l'elenco
                    frmEmulation.ClearBPE()

                    ' Non Aggiorno i Dati sulla Vista per i totlaizzatori
                    _UpdateFormEmulator = False

                Case "FILL_DATA_INFO_RESPONSE_UPDATE"

                    ' Non Aggiorno i Dati sulla Vista per i totlaizzatori
                    _UpdateFormEmulator = False

                    ' Aggiorno solo la vista elenco
                    frmEmulation.AddItemBPE(ItemData,, True)

                Case "FILL_DATA_INFO_RESPONSE_END"

                    ' Non Aggiorno i Dati sulla Vista per i totlaizzatori
                    _UpdateFormEmulator = False

                Case "INITIALIZE_EMULATOR_SOFTWARE"

                    ' Rispulisce la Griglia BPE
                    'frmEmulation.ClearBPE()

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "INITIALIZE_EMULATOR_HARDWARE"

                    ' Rispulisce la Griglia BPE
                    frmEmulation.ClearBPE()

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "SINGLEP_CHECKED_VALID"

                    ' In questo caso segnaliamo che il BP 
                    ' controllato è valido.
                    frmEmulation.UpdateValidBPCResultCheck(ItemData, True)

                    ' Non aggiorno altro
                    _UpdateTotals = False
                    _UpdateFormEmulator = False

                Case "SINGLEP_CHECKED_INVALID"

                    ' In questo caso segnaliamo che il BP 
                    ' controllato non è valido.
                    frmEmulation.UpdateValidBPCResultCheck(ItemData, False)

                    ' Non aggiorno altro
                    _UpdateTotals = False
                    _UpdateFormEmulator = False


                Case "SINGLEP_UPDATE"

                    ' Aggiungo l'elemento al controllo Griglia
                    frmEmulation.AddItemOnGrid(ItemData)

                    ' Blocco l'importo scalabile per aggiornamento
                    frmEmulation.LockAmountScalable = True
                    frmEmulation.Payable = m_CurrentAmountScalable

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "SINGLEP_REMOVE_SELECTED"

                    ' Sul Form rimuovo dalla griglia l'elemento
                    frmEmulation.RemoveCurrentItemSelectedOnGrid()

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "SINGLEP_REMOVE_BARCODE"

                    ' Sul Form rimuovo dalla griglia l'elemento
                    frmEmulation.RemoveItemOnGridByKey(m_CurrentBarcodeScan)

                    ' Aggiorno i Dati sulla Vista
                    _UpdateFormEmulator = True

                Case "END_PAY"

                    _UpdateTotals = True

                Case "END_VOID"

                    _UpdateTotals = True

                Case "END"

            End Select

            ' X il Form
            If _UpdateFormEmulator Then

                If Not frmEmulation Is Nothing Then

                    Dim NumItemsScalablePayed As Integer = m_PartialBPUsedToPay_XS
                    Dim NumItemsScalableVoided As Integer = m_PartialBPUsedToVoid_XS
                    If Mode = "SINGLEP_REMOVE_SELECTED" Then
                        ' Questo quando l'elemento viene eliminato sulla lista quando in PAY
                        NumItemsScalablePayed -= m_PartialBPUsedToVoid_XS
                    ElseIf Mode = "SINGLEP_REMOVE_BARCODE" Then
                        ' Questo quando l'elemento viene eliminato sulla lista quando in PAY
                        'NumItemsScalablePayed -= m_PartialBPUsedToVoid_XS
                    End If

                    ' Aggiorno i Dati di rilievo da mostrare sulla vista
                    frmEmulation.UpdateDataValues(
                            CurrentTransact,                '   Il Totale sulla TA o Quello Scalabile da visualizzare
                            NumItemsScalablePayed,                      '   N° elaborati Paganti 
                            m_PartialPayed_XS - m_PartialVoided_XS,     '   Pagato  
                            m_PartialValueExcedeed_XS,      '   L'eccesso usato come resto in Pay
                            NumItemsScalableVoided,         '   Il Numero di Elementi Elaborati come Voci di Storno
                            m_PartialVoided_XS,             '   L'Importo Stornato fino adesso
                            m_PartialBPNotValid_XS,         '   Il Numero di Elementi Elaborati come Voci non Valide (BPE)
                            m_PartialInvalid_XS             '   L'Importo Scartato fino adesso (Non Validi scartati in BPE)
                        )

                End If

            End If

            ' X il DataResult
            If _UpdateTotals Then

                m_TotalBPUsedToPay_CS = 0
                m_TotalPayed_CS = 0
                m_TotalBPUsedToVoid_CS = 0
                m_TotalVoided_CS = 0
                m_TotalBPNotValid_CS = 0
                m_TotalInvalid_CS = 0

                For Each qItm As PaidEntry In WriterResultDataList

                    If Not qItm.Invalid And Not qItm.Voided Then

                        ' Payed
                        m_TotalBPUsedToPay_CS += 1              '   N° elaborati
                        m_TotalPayed_CS += qItm.DecimalValue    '   Pagato ( In Payment ci possono anche essere degli storni iniziali come in BPC )

                    ElseIf Not qItm.Invalid Then

                        ' Voided
                        m_TotalBPUsedToVoid_CS += 1             '   Il Numero di Elementi Elaborati come Voci di Storno
                        m_TotalVoided_CS += qItm.DecimalValue   '   L'Importo Stornato fino adesso

                    Else

                        ' Not Valid
                        m_TotalBPNotValid_CS += 1               '   Il Numero di Elementi Elaborati come Voci non Valide (BPE)
                        m_TotalInvalid_CS += qItm.DecimalValue  '   L'Importo Scartato fino adesso (Non Validi scartati in BPE)

                    End If

                Next

                ' Eccesso per resto ( Solo in Pagamento )
                If m_CommandToCall = enCommandToCall.Payment Then
                    If (m_TotalPayed_CS - m_TotalVoided_CS) > m_InitialPaymentsInTA Then
                        m_TotalValueExcedeed_CS = m_InitialPaymentsInTA - (m_TotalPayed_CS - m_TotalBPUsedToVoid_CS)      '   L'eccesso usato come resto in Pay
                    End If
                End If

                If m_TypeProxy = enTypeProxy.Service Then

                    m_TotalPayed_CS += m_TotalVoided_CS ' Ci penserà la property a dare il gisuto risultato
                End If

            End If


        Catch ex As Exception

            ' Sollevo l'eccezione
            Throw New ExceptionProxyArgentea("UPDATERESULTDATA", ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore Update form come Form compatibile per l'evento -- Contattare Assistenza --", ex)

        End Try

        System.Threading.Thread.Sleep(50)
        'System.Windows.Forms.Application.DoEvents()

    End Sub

    ''' <summary>
    '''     Resetta a 0 I Contatori
    '''     interni per i totali di 
    '''     Sessione Iniziali/Finali.
    ''' </summary>
    Private Sub _resetSessionCountersFinals()

        m_TotalPayed_CS = 0
        m_TotalBPUsedToPay_CS = 0
        m_TotalValueExcedeed_CS = 0
        '
        m_TotalBPUsedToVoid_CS = 0
        m_TotalVoided_CS = 0
        m_TotalValueExtraVoidNotContabilizated = 0
        '
        m_TotalBPNotValid_CS = 0
        m_TotalInvalid_CS = 0
        '
        m_TotalBPElaborated_CS = 0

    End Sub

    ''' <summary>
    '''     Resetta a 0 I Contatori
    '''     interni per i totali di 
    '''     Sessione Parziale.
    ''' </summary>
    Private Sub _ResetSessionCountersPartials()

        m_PartialBPUsedToPay_XS = 0
        m_PartialPayed_XS = 0
        m_PartialValueExcedeed_XS = 0
        '
        m_PartialBPUsedToVoid_XS = 0
        m_PartialVoided_XS = 0
        '
        m_PartialBPNotValid_XS = 0
        m_PartialInvalid_XS = 0

    End Sub

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
            ' 1° Step ( Riporto i Parziali al Totale)
            '       Una volta completato i  dati  parziali
            '       aggiornano i dati per la response data
            '       ( PRIMA DEL CLOSE HANDLER )
            '


            ' Status Form Message
            _SetOperationStatusForm(NOT_INFO_POS_CLOSING, "POS Closing...", FormEmulationArgentea.InfoStatus.Warning, 8)
            System.Threading.Thread.Sleep(50)

            '
            ' 2° Step ( Print di eventuali scontrini pos )
            '       Print Last Receipt Solo in pagamento 
            '       e solo per quelli POS Hardware
            '
            If m_CommandToCall = enCommandToCall.Payment And m_TypeProxy = enTypeProxy.Pos Then

                If Not m_LastResponseRawArgentea Is Nothing And Not m_ServiceStatus = enProxyStatus.InError Then

                    ' Status Form Message
                    _SetOperationStatusForm(NOT_INFO_POS_PRINTING, "Print receipt bp...", FormEmulationArgentea.InfoStatus.Warning, 9)
                    System.Threading.Thread.Sleep(50)

                    PrintReceipt(m_LastResponseRawArgentea)

                End If

            End If

        Catch ex As Exception

            ' Status Form Message
            _SetOperationStatusForm(NOT_INFO_POS_PRINTERR, "Print receipt error...", FormEmulationArgentea.InfoStatus.Warning, 9)
            System.Threading.Thread.Sleep(100)

            ' Log locale (Errore di reprint dello scontrino non bloccante)
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Printer for print recipient in proxy Argentea: Hardware Output")
            LOG_ErrorInTry(getLocationString("ProxyArgentea"), ex)

        End Try

        ' Status Form Message
        _SetOperationStatusForm(NOT_INFO_POS_CLOSING, "POS Closing...", FormEmulationArgentea.InfoStatus.Wait, 8)

        ' (Happy Ending)
        Try

            '
            ' 3° Step ( Evento in Chiusura )
            '       Evento chiave di chiusura
            '
            If m_CommandToCall = enCommandToCall.Payment Then

                ' Preparo il DataResponse e la Vista sul Form
                _UpdateResultData("END_PAY", m_InitialPaymentsInTA, Nothing)

                '
                '   Evento collect data in chiusura PAY
                '
                If Not m_ServiceStatus = enProxyStatus.InError Then
                    RaiseEvent Event_ProxyCollectDataTotalsAtEnd(Me, _DataResponse)
                Else
                    RaiseEvent Event_ProxyCollectDataTotalsAtEnd(Me, Nothing)
                End If

            ElseIf m_CommandToCall = enCommandToCall.Void Then

                ' Preparo il DataResponse e la Vista sul Form
                _UpdateResultData("END_VOID", m_CurrentAmountScalable, Nothing)

                '
                '   Evento collect data in chiusura VOID
                '
                If Not m_ServiceStatus = enProxyStatus.InError Then
                    RaiseEvent Event_ProxyCollectDataVoidedAtEnd(Me, _DataResponse)
                Else
                    RaiseEvent Event_ProxyCollectDataVoidedAtEnd(Me, Nothing)
                End If

            End If


        Catch ex As Exception

            ' Signal Error Ritardato Status
            _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error...", FormEmulationArgentea.InfoStatus.Error, 8)

            ' Intercettiamo l'errore per il  contesto  probabilmente
            ' erchè il consumer non l'ha fatto per suo conto, quindi
            ' rimane che per noi il consumer con i dati non è aggioranto.

            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_ON_EVENT_DATA, "Errore nell'evento durante il collect dei dati al consumer del Proxy -- Consumer in errore --", ex)

        Finally

            ' Chiudo a questo punto il form
            frmEmulation.bDialogActive = False

        End Try

    End Sub

#End Region

#Region "Functions Private per Emulation Pos >>Software Service mode<<"

    ''' <summary>
    '''     Classe di appoggio per i log delle chiamate
    ''' </summary>
    Private Class tLogErr

        Public ReadOnly Property Status As String
        Public ReadOnly Property ErrorMessage As String
        Public ReadOnly Property ResponseRawArgentea As ArgenteaFunctionReturnObject

        Sub New(_Status As String, _ErrorMessage As String, _ResponseRawArgentea As ArgenteaFunctionReturnObject)
            Status = _Status
            ErrorMessage = _ErrorMessage
            ResponseRawArgentea = _ResponseRawArgentea
        End Sub

    End Class

    ''' <summary>
    '''     Inizializza la Sessione verso Argentea
    '''     e parte la numerazione interna delle chiamate
    '''     da 1
    ''' </summary>
    ''' <returns>Il codice di stato Riuscito Non RIuscito Interno <see cref="StatusCode"/></returns>
    Private Function CallInitialization(_funcName As String) As Boolean
        Dim actApiCall As enApiToCall
        Dim funcName As String = "CallInitialization"
        Dim metdName As String = "n/d"
        Dim _SignalErrorAndExitResetCounter As Boolean = False

        ' Se nel contesto della sessione per questo scontrino
        ' è già stato iniziato un nuovo Ticket non lo richiediamo.
        If FLAG_STATIC_INITIALIZATED Then
            Return True
        End If

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Status Corrente
        CallInitialization = False
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing


        ' Prima operazione di Avvio per il demat
        actApiCall = enApiToCall.Initialization
        metdName = "OpenTicketBC"

#If DEBUG_SERVICE = 0 Then

        retCode = ArgenteaCOMObject.OpenTicketBC(
            m_ProgressiveCall,
            Get_ReceiptNumber,
            Get_CodeCashDevice,
            RefTo_MessageOut
        )

        ''msgUtil.ShowMessage(m_TheModcntr, "INIZIAZLIZZAZIONE " + RefTo_MessageOut, "LevelITCommonModArgentea_", PosDef.TARMessageTypes.TPINFORMATION)

#Else

        ''' Per Test
        If Not _flagCallOnetimeResetIncrement Then
            ' 1° Tentativo
            RefTo_MessageOut = "KO-903-PROGRESSIVO FUORI SEQUENZA-----0---"            ' <-- x test su questo signal
            'RefTo_MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 
            retCode = ArgenteaFunctionsReturnCode.OK ' .OK
        Else
            ' 2° tenttivo
            'RefTo_MessageOut = "KO-903-ALTRO ERRORE-----0---"            ' <-- x test su questo signal
            RefTo_MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 
            retCode = ArgenteaFunctionsReturnCode.OK ' .OK
        End If
        ''' to remove:

#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Controllo se è necessario riallineare prima di segnalare eventuali KO
        Dim RequestResetCounter As StatusCode = CheckIfNecessaryToResetIncrement(m_LastResponseRawArgentea)
        Dim IsNecessaryToResetCounter As StatusCode = CheckResponseIfIsNecessaryToResetCounter(RequestResetCounter, funcName, "")
        If IsNecessaryToResetCounter = StatusCode.OK Then ' True Then
            Return False
        ElseIf IsNecessaryToResetCounter = StatusCode.RESETCOUNTER_RETURN_TO_CALL Then ' 2
            Return m_LastResponseRawArgentea.Successfull
        Else ' --> If IsNecessaryToResetCounter = StatusCode.KO Then ' False Then
            ''
        End If

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' L'Iniziailizzazione deve essere chiamata una volta sola nel contesto
            ' della sessione in corso, se la sessione sarà conlcusa sarà  chiamata
            ' nuovamente.
            FLAG_STATIC_INITIALIZATED = True


            ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
            _IncrementProgressiveCall()

            ' ** OK --> INIZIALIZZATA e corretamente chiamata ad Argentea
            LOG_Debug(getLocationString(funcName), "Inizialization " & m_CurrentBarcodeScan & " successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            m_LastStatus = GLB_FAILED_INITIALIZATION
            m_LastErrorMessage = "Inizializzazione fallita per KO remoto with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            ' ** KO --> Non inizializzata da parte di Argentea per errore remoto in risposta a questo codice.
            LOG_Debug(getLocationString(funcName), "Inizialization " & m_CurrentBarcodeScan & " remote failed on initialization argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

    ''' <summary>
    '''     Esegue la chiamata di Check di un BP secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    ''' <returns>Il codice di stato Riuscito Non RIuscito Interno <see cref="StatusCode"/></returns>
    Private Function CallCheckTicket(_funcName As String) As StatusCode
        Dim actApiCall As enApiToCall
        Dim funcName As String = "CallDematerialize"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Status Corrente
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallCheckTicket = StatusCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        ' Chiamata per la Dematerializzazione del BP
        actApiCall = enApiToCall.CheckBP
        metdName = "CheckAvailablityBP"

#If DEBUG_SERVICE = 0 Then

        ' Active to first Argentea COM communication                                **** CHECK SIPONIBILITA'
        retCode = ArgenteaCOMObject.CheckAvailablityBP(
                    GetCodifiqueReceipt(TypeCodifiqueProtocol.Dematerialization),
                    RefTo_MessageOut
                )

#Else


        ''' Per Test questo è il suio CSV
        'RefTo_MessageOut = "KO-3-Buono pasto gia' rientrato-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"       ' <-- x test su questo signal
        RefTo_MessageOut = "Buono-000-Buono Valido-529-8897456-12345687-201809201733577-ARGENTEA-"            ' <-- x test su questo signal
        RefTo_MessageOut = "Coupon-039-Coupon NonValido-529----SCONTIA-"            ' <-- x test su questo signal
        RefTo_MessageOut = "Coupon-000-Coupon Valido-529---201809201733577-SCONTIA-"            ' <-- x test su questo signal
        ''' to remove:
        retCode = ArgenteaFunctionsReturnCode.OK
#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Controllo se è necessario riallineare prima di segnalare eventuali KO
        Dim RequestResetCounter As StatusCode = CheckIfNecessaryToResetIncrement(m_LastResponseRawArgentea)
        Dim IsNecessaryToResetCounter As StatusCode = CheckResponseIfIsNecessaryToResetCounter(RequestResetCounter, funcName, "")
        If IsNecessaryToResetCounter = StatusCode.OK Then ' True Then
            Return False
        ElseIf IsNecessaryToResetCounter = StatusCode.RESETCOUNTER_RETURN_TO_CALL Then ' 2
            Return m_LastResponseRawArgentea.Successfull
        Else ' --> If IsNecessaryToResetCounter = StatusCode.KO Then ' False Then
            ''
        End If

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
            _IncrementProgressiveCall()

            ' Se la risposta argenta richiede un ulteriore 
            ' conferma allora procedo ad uscire per il flow.
            If m_LastResponseRawArgentea.RequireCommit Then

                ' ** OK --> CONTROLLO in check corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP check with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' RICHIESTO CONFERMA
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                Return StatusCode.CONFIRMREQUEST

            Else

                ' ** OK --> CHECK corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP check valid " & m_CurrentBarcodeScan & " return with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' COMPLETATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                Return StatusCode.OK

            End If

        Else

            ' ** KO --> Non controllato da risposta Argentea per errore remoto in relazione a questo codice.
            LOG_Debug(getLocationString(funcName), "BP check " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

            ' NON EFFETTUATO
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)

            ' SIGNAL
            m_LastStatus = GLB_FAILED_DEMATERIALIZATION
            m_LastErrorMessage = "Conferma su check fallita per KO remoto with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            Return StatusCode.KO

        End If

    End Function

    ''' <summary>
    '''     Esegue la chiamata di Dematerializzazione secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    ''' <returns>Il codice di stato Riuscito Non RIuscito Interno <see cref="StatusCode"/></returns>
    Private Function CallDematerialize(_funcName As String) As StatusCode
        Dim actApiCall As enApiToCall
        Dim funcName As String = "CallDematerialize"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Status Corrente
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallDematerialize = StatusCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        ' Chiamata per la Dematerializzazione del BP
        actApiCall = enApiToCall.SinglePayment
        metdName = "DematerializzazioneBP"

#If DEBUG_SERVICE = 0 Then

        ' Active to first Argentea COM communication                                **** DEMATERIALIZZAZIONE
        retCode = ArgenteaCOMObject.DematerializzazioneBP(
                    GetCodifiqueReceipt(TypeCodifiqueProtocol.Dematerialization),
                    RefTo_MessageOut
                )

#Else

        ''' Per Test questo è il suio CSV
        RefTo_MessageOut = "OK-0 - BUONO VALIDATO CON SUCCESSO-" + m_CurrentTransactionID + "-700-ARGENTEA-" + m_CurrentTransactionID + "-0-202--"    ' <-- x test 
        'RefTo_MessageOut = "KO-3-Buono pasto gia' rientrato-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"       ' <-- x test su questo signal
        'RefTo_MessageOut = "KO-903-Sequenza non valida-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"            ' <-- x test su questo signal
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Controllo se è necessario riallineare prima di segnalare eventuali KO
        Dim RequestResetCounter As StatusCode = CheckIfNecessaryToResetIncrement(m_LastResponseRawArgentea)
        Dim IsNecessaryToResetCounter As StatusCode = CheckResponseIfIsNecessaryToResetCounter(RequestResetCounter, funcName, "")
        If IsNecessaryToResetCounter = StatusCode.OK Then ' True Then
            Return False
        ElseIf IsNecessaryToResetCounter = StatusCode.RESETCOUNTER_RETURN_TO_CALL Then ' 2
            Return m_LastResponseRawArgentea.Successfull
        Else ' --> If IsNecessaryToResetCounter = StatusCode.KO Then ' False Then
            ''
        End If

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
            _IncrementProgressiveCall()

            ' Se la risposta argenta richiede un ulteriore 
            ' conferma allora procedo ad uscire per il flow.
            If m_LastResponseRawArgentea.RequireCommit Then

                ' ** OK --> DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP dematirializated with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' RICHIESTO CONFERMA
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                Return StatusCode.CONFIRMREQUEST

            Else

                ' ** OK --> DEMATERIALIZZATO corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP dematerializated " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' COMPLETATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                Return StatusCode.OK

            End If

        Else

            ' ** KO --> Non dematerializzato da risposta Argentea per errore remoto in relazione a questo codice.
            LOG_Debug(getLocationString(funcName), "BP dematerializated " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

            ' NON EFFETTUATO
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)

            ' SIGNAL
            m_LastStatus = GLB_FAILED_DEMATERIALIZATION
            m_LastErrorMessage = "Conferma su dematerializzazione fallita per KO remoto with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            Return StatusCode.KO

        End If

    End Function

    ''' <summary>
    '''     Esegue la chiamata di Reverse da uno già Dematerializzato secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    ''' <returns>Il codice di stato Riuscito Non RIuscito Interno <see cref="StatusCode"/></returns>
    Private Function CallReverseMaterializated(_funcName As String, IdTransactionToReverse As String) As StatusCode
        Dim actApiCall As enApiToCall
        Dim funcName As String = "CallReverseMaterializated"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Status Corrente
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallReverseMaterializated = StatusCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        ' Annullo di una dematerializzazione già fatta in corso di sessione prima del Close
        actApiCall = enApiToCall.SingleVoid
        metdName = "ReverseTransactionDBP"

#If DEBUG_SERVICE = 0 Then

        ' Active to first Argentea COM communication                                **** ANNULLO BUONO GIA' MATERIALIZZATO
        retCode = ArgenteaCOMObject.ReverseTransactionDBP(
                    GetCodifiqueReceipt(TypeCodifiqueProtocol.Reverse,IdTransactionToReverse),
                    RefTo_MessageOut
                )
#Else

        ''' Per Test
        Dim itm As String = GetCodifiqueReceipt(TypeCodifiqueProtocol.Reverse, IdTransactionToReverse)
        If False Then
            RefTo_MessageOut = "OK-0 - BUONO STORNATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
        Else
            If Not _flagCallOnetimeResetIncrement Then
                RefTo_MessageOut = "KO-903-PROGRESSIVO FUORI SEQUENZA-----0---"            ' <-- x test su questo signal
            Else
                'RefTo_MessageOut = "KO-0 - BUONO GIa' STORNATO -68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
                RefTo_MessageOut = "OK-0 - BUONO STORNATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
            End If
        End If
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Controllo se è necessario riallineare prima di segnalare eventuali KO
        Dim RequestResetCounter As StatusCode = CheckIfNecessaryToResetIncrement(m_LastResponseRawArgentea)
        Dim IsNecessaryToResetCounter As StatusCode = CheckResponseIfIsNecessaryToResetCounter(RequestResetCounter, funcName, IdTransactionToReverse)
        If IsNecessaryToResetCounter = StatusCode.OK Then ' True Then
            Return False
        ElseIf IsNecessaryToResetCounter = StatusCode.RESETCOUNTER_RETURN_TO_CALL Then ' 2
            Return m_LastResponseRawArgentea.Successfull
        Else ' --> If IsNecessaryToResetCounter = StatusCode.KO Then ' False Then
            ''
        End If

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
            _IncrementProgressiveCall()

            ' Se la risposta argenta richiede un ulteriore 
            ' conferma allora procedo ad uscire per il flow.
            If m_LastResponseRawArgentea.RequireCommit Then

                ' ** OK --> REQUEST CONFIRM FOR REVERSE SU DEMATERIALIZZATO richiesto da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP reverse dematirializated with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' RICHIESTO CONFERMA
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                Return StatusCode.CONFIRMREQUEST

            Else

                ' ** OK --> REVERSE SU DEMATERIALIZZATO correttamente effettuato su chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP reverse dematirializated " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

                ' COMPLETATO
                m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
                'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID

                Return StatusCode.OK

            End If

        Else

            ' ** KO --> Non reverse su dematerializzato da risposta Argentea per errore remoto in relazione a questo codice.
            LOG_Debug(getLocationString(funcName), "BP reverse dematerializated " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

            ' NON EFFETTUATO
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)

            ' SIGNAL
            m_LastStatus = GLB_FAILED_VOIDDEMATERIALIZATION
            m_LastErrorMessage = "Conferma su storno per questa dematerializzazione fallita per KO remoto with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            Return StatusCode.KO

        End If

    End Function

    ''' <summary>
    '''     Esegue una chiamata con protocollo di Conferma verso
    '''     Argentea per confirm su Dematerializzazione o Reverse.
    ''' </summary>
    ''' <param name="funcName">Nome della funzione chiamante</param>
    ''' <returns>Il codice di stato Riuscito Non RIuscito Interno <see cref="StatusCode"/></returns>
    Private Function CallConfirmOperation(_funcName As String, sConfirmOperation As String) As StatusCode
        Dim actApiCall As enApiToCall
        Dim funcName As String = "CallConfirmOperation"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Status Corrente
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallConfirmOperation = StatusCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        ' Conferma (se richiesto nella chiamata precedente) di una dematerializzazione
        actApiCall = enApiToCall.Confirmation
        metdName = "CommitTransactionDBP"

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

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Controllo se è necessario riallineare prima di segnalare eventuali KO
        Dim RequestResetCounter As StatusCode = CheckIfNecessaryToResetIncrement(m_LastResponseRawArgentea)
        Dim IsNecessaryToResetCounter As StatusCode = CheckResponseIfIsNecessaryToResetCounter(RequestResetCounter, funcName, sConfirmOperation)
        If IsNecessaryToResetCounter = StatusCode.OK Then ' True Then
            Return False
        ElseIf IsNecessaryToResetCounter = StatusCode.RESETCOUNTER_RETURN_TO_CALL Then ' 2
            Return m_LastResponseRawArgentea.Successfull
        Else ' --> If IsNecessaryToResetCounter = StatusCode.KO Then ' False Then
            ''
        End If

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
            _IncrementProgressiveCall()

            ' ** OK --> CONFIRM su REVERSE o DEMATERIALIZZATO effettuata corretamente da chiamata ad Argentea
            LOG_Debug(getLocationString(funcName), "BP confirm " & sConfirmOperation & " for " & m_CurrentBarcodeScan & " successfuly on call with message " & m_LastResponseRawArgentea.SuccessMessage)

            ' COMPLETATO
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)
            'm_CurrentTerminalID = m_LastResponseRawArgentea.TerminalID
            CallConfirmOperation = StatusCode.OK
            Return CallConfirmOperation

        Else

            ' ** KO --> Non confirm su reverse o dematerializzato da risposta Argentea per errore remoto in relazione a questo codice.
            LOG_Debug(getLocationString(funcName), "BP confirm " & sConfirmOperation & " for " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

            ' NON EFFETTUATO
            m_CurrentValueOfBP = m_LastResponseRawArgentea.GetAmountValue(m_ProtoFractMode)

            ' SIGNAL
            m_LastStatus = GLB_FAILED_CONFIRMATION
            m_LastErrorMessage = "Conferma su azione corrente fallita per KO remoto with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            Return StatusCode.KO

        End If

    End Function

    ''' <summary>
    '''     Esegue un riallineamento del contatore
    '''     remoto sul servizio Argentea se la risposta
    '''     rispetto all'ultima chiamata mi dà un 903
    ''' </summary>
    ''' <param name="LastResponse">La risposta che si deve prendere in analisi per avere il comportamento rispetto a 903 come staus<see cref="ArgenteaFunctionReturnObject"/></param>
    ''' <returns>
    '''     0 non si è dovuto riallineare procedo sulla funzione
    '''     1 si è dovuto riallineare quindi ho ripetuto nuovamente la funzione 
    '''     2 si è riallineato se si deve ritentare dopo il riallineamento oppure False se ha fallito o se diverso da 903 per segnalazione e inoltro dell'errore
    ''' </returns>
    Private Function CheckIfNecessaryToResetIncrement(ByRef LastResponse As ArgenteaFunctionReturnObject) As StatusCode
        Dim actApiCall As enApiToCall
        Dim funcName As String = "CallResetIncrement"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        ' Status Corrente
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        If LastResponse.CodeResult.Trim() = "903" Then

            ' Questo metodo richiama ricorsivmente Initialization
            ' e per questo usciamo dopo la prima chiamata occorrente.
            If _flagCallOnetimeResetIncrement Then
                _flagCallOnetimeResetIncrement = False
                Return StatusCode.RESETCOUNTER_REQUEST_TO_RECALL
                Return -1    ' Salta incodizionato
            End If
            _flagCallOnetimeResetIncrement = True

            ' 
            ' Riformatto lo status degli errori
            ' conservando per log quello precedente
            ' fino alla fine della sessione.
            '
            m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

            '
            ' Quindi reset per la chiamata corrente
            '
            m_LastStatus = Nothing
            m_LastErrorMessage = Nothing
            m_LastResponseRawArgentea = Nothing

            ' ** Qui essendo su questo specifico errore 903 ritentiamo un
            ' ** riallinemento per ritentare la reinelizzazione.

            ' Azione di Riallineamento del contatore remoto
            actApiCall = enApiToCall.ResetCounter
            metdName = "RiallineamentoDBP"

#If DEBUG_SERVICE = 0 Then

            ' Tento Riallineamento
            retCode = ArgenteaCOMObject.RiallineamentoDBP(RefTo_MessageOut)

#Else

            ''' Per Test
            If True Then
                RefTo_MessageOut = "OK--RIALLINEATO-----0---" ' <-- x test 
                retCode = ArgenteaFunctionsReturnCode.OK
            Else
                RefTo_MessageOut = "KO--NON RIALLINEATO-----0---" ' <-- x test 
                retCode = ArgenteaFunctionsReturnCode.KO
            End If
            ''' to remove:

#End If

            ' ** Response Grezzo in debug
            LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

            ' Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
            m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If m_LastResponseRawArgentea.Successfull Then

                ' Reset contatori interni e progressivo statico a partire nuovamente da 1.
                ''msgUtil.ShowMessage(m_TheModcntr, "RIALLINEO DA " + CStr(m_ProgressiveCall), "LevelITCommonModArgentea_", PosDef.TARMessageTypes.TPINFORMATION)
                m_ProgressiveCall = -1
                _IncrementProgressiveCall()

                ' ** OK --> INIZIALIZZATA e corretamente chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "Reset Counter remote for BP " & m_CurrentBarcodeScan & " successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)

                Return StatusCode.RESETCOUNTER_OK
                Return 1        ' Riallinemaneto effettuato con esito OK (Richiede nuovamente l'initialization)

            Else

                ' ** KO --> Non inizializzata da parte di Argentea per errore remoto in risposta a questo codice.
                LOG_Debug(getLocationString(funcName), "Reset Counter remote for BP " & m_CurrentBarcodeScan & " failed on call method argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' SIGNAL
                m_LastStatus = GLB_FAILED_RESETCOUNTER
                m_LastErrorMessage = "Il Reset per il numero di operazione remoto per i BP ha dato KO con questo messaggio di errore - " & m_LastResponseRawArgentea.ErrorMessage & " - "

                Return StatusCode.RESETCOUNTER_KO
                Return 2        ' Riallinemaneto effettuato con esito KO

            End If

        Else

            ' Altro errore non 903 come risposta remota
            Return StatusCode.RESETCOUNTER_CONTINUE
            Return 0        ' Non era richiesto il riallinemento perchè la risposta è altro tipo

        End If

    End Function

    ''' <summary>
    '''     Richiama il Reset e riporta lo status dopo 
    '''     due 1 tentativo di riallineamento.
    ''' </summary>
    ''' <param name="RequestResetCounter">Il risultato dalla funzione CheckIfNecessaryToResetIncrement</param>
    ''' <param name="funcName">Il nome della funzione che si deve chiamare ricorsivamente dopo l'eventuale reset</param>
    ''' <param name="IdTransactionToReverse">Argomento necessario per la funzione ricorsiva di Void</param>
    ''' <returns>
    '''     True    E' stato necessario fare il Rialliniamento      -- > ( Il Chiamante deve ritornare con False per ripetere da quel punto verso il 2 )
    '''     False   Non è stato necessario fare il Rialliniamento   -- > ( Il Chimante continua il suo corso )
    '''     2       E' stato necessario fare il Rialliniamento      -- > ( Il Chiamante deve riprendere la Risposta della Chiamata ricorsiva a se stessa per il Successfully (True/False) )
    ''' </returns>
    Private Function CheckResponseIfIsNecessaryToResetCounter(RequestResetCounter As StatusCode, funcName As String, Arg1 As String) As StatusCode ' Byte ' Boolean

        Dim _SignalErrorAndExitResetCounter As Boolean = False

        ' In base all'esito.:
        If RequestResetCounter = StatusCode.RESETCOUNTER_REQUEST_TO_RECALL Then

            ' (Salto incodinzionato riport lo status dell'operazione completa ricorsivamente)
            ' Riporta al chimante la m_LastResponseRawArgentea.Successfull
            Return StatusCode.RESETCOUNTER_RETURN_TO_CALL  '2

        ElseIf RequestResetCounter = StatusCode.RESETCOUNTER_CONTINUE Then

            ' Riallineamento non necessario
            _SignalErrorAndExitResetCounter = StatusCode.KO ' False

        ElseIf RequestResetCounter = StatusCode.RESETCOUNTER_OK Then

            ' Riallinemaneto effettuato inizializzo nuovamente (flag interno impedirà il rientro)
            If _flagCallOnetimeResetIncrement Then

                If funcName = "CallDematerialize" Then

                    ' Ricorsivo su Dematerialize
                    CallDematerialize(funcName)

                ElseIf funcName = "CallReverseMaterializated" Then

                    ' Ricorsivo su ReverseMaterializated
                    CallReverseMaterializated(funcName, Arg1)

                ElseIf funcName = "CallInitialization" Then

                    ' Ricorsivo su Initilization
                    CallInitialization(funcName)

                ElseIf funcName = "CallConfirmOperation" Then

                    ' Ricorsivo su Confirm requested
                    CallConfirmOperation(funcName, Arg1)

                Else

                    Throw New ExceptionProxyArgentea("CheckResponseIfIsNecessaryToResetCounter", ExceptionProxyArgentea.LOC_ERROR_NOT_CLASSIFIED, "Chiamata ricorsiva non corretta")

                End If

                ' Se nell'eventualità che si ripeta....
                If m_LastResponseRawArgentea.CodeResult.Trim() <> "903" Then

                    ' (Salto incodinzionato riport lo status dell'operazione completa ricorsivamente)
                    'Return m_LastResponseRawArgentea.Successfull
                    Return StatusCode.RESETCOUNTER_RETURN_TO_CALL ' 2

                Else

                    ' Vuol dire che si è ripetuto per la seconda volta il Progressivo fuori sequenza
                    _SignalErrorAndExitResetCounter = StatusCode.OK ' True ' Lo presentiamo ancora come Fail in riallinemaneto per fargli riprovare all'utente.

                End If

            Else

                ' Riallineamento effettuato con errore KO  (2° tentativo)
                _SignalErrorAndExitResetCounter = StatusCode.OK ' True

            End If

        ElseIf RequestResetCounter = StatusCode.RESETCOUNTER_KO Then

            ' Riallineamento effettuato con errore KO 
            _SignalErrorAndExitResetCounter = StatusCode.OK ' True

        Else

            ' Altro in modo errato nostro....
            _SignalErrorAndExitResetCounter = StatusCode.OK ' True

        End If

        ' Reset del flag condvisio per evitare l'effetto ricorsione
        _flagCallOnetimeResetIncrement = False ' per i successivi

        If _SignalErrorAndExitResetCounter = StatusCode.OK Then ' True Then

            ' 2° e ultimo tentativo altrimenti è fallito punto e basta

            ' ** KO --> NON INIZIALIZZATA e tentativo eventuale riallineamento FALLITO
            LOG_Debug(getLocationString(funcName), "Action for reset fail for this element " & m_CurrentBarcodeScan & " with message error " & m_LastResponseRawArgentea.ErrorMessage)

            ' SIGNAL
            m_LastStatus = GLB_FAILED_RESETCOUNTER
            m_LastErrorMessage = "Il Reset per il numero di operazione remoto per i BP ha dato KO con questo messaggio di errore - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            Return StatusCode.KO ' False

        End If

    End Function

    Private Function ValidationVoucherRequest(barcode As String) As Boolean
        'Logic Comunication Barcode at Argentea Supplier

        ValidationVoucherRequest = False

    End Function

#End Region

#Region "Private per Emulation Pos >>Hardware Comunication mode<<"

    ''' <summary>
    '''     Inizializza la Sessione verso il Dispositivo 
    '''     Pos di Argentea (Monetica) per la Carta Ticket.
    '''     Aspetta da utente un insieme di pagamenti
    ''' </summary>
    ''' <returns>True o False</returns>
    Private Function CallMultiplePaymentsOnPosHardware(_funcName As String) As Boolean
        Dim actApiCall As enApiToCall = enApiToCall.None
        Dim funcName As String = "CallMultiplePaymentsOnPosHardware"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = String.Empty

        ' Se nel contesto della sessione per questo scontrino
        ' è già stato iniziato un nuovo Ticket non lo richiediamo.
        'If FLAG_STATIC_INITIALIZATED Then
        'Return False
        'End If

        ' Status Corrente
        CallMultiplePaymentsOnPosHardware = False
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        '
        '   OUT
        '   L'Id di transazione  recuperata
        '   dopo le chiamate verso Argentea.
        '   (Passato alla dll COM di Argentea e fillato dalla stessa)
        '
        Dim RefTo_Transaction_Identifier As String = String.Empty

        ' Prima operazione di Avvio per il demat
        actApiCall = enApiToCall.MultiplePayments
        metdName = "PaymentBPE"

#If DEBUG_SERVICE = 0 Then

        ' (Idle)
        retCode = ArgenteaCOMObject.PaymentBPE(
                CInt(m_PayableAmount * m_ParseFractMode),
                    RefTo_Transaction_Identifier,
                    RefTo_MessageOut
                )
#Else

        ''' Per Test
        System.Threading.Thread.Sleep(500)
        'System.Windows.Forms.Application.DoEvents()
        RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;5|2|1020|3|414;104;PELLEGRINI;  PAGAMENTO BUONO PASTO " ' <-- x test 
        'RefTo_MessageOut = "KO;    DATI NON RICEVUTI    ;;;;" ' <-- x test 
        '''
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' (With Entrap) Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        ' in questo caso gestiamo l'eccezione per rimanere nel Form corrente
        Try
            m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)
        Catch ex As ExceptionProxyArgentea
            If ex.ErrorComunication Or ex.ErrorActionArgentea Then
                If frmEmulation.AutoCloseOnCompleteOperation Then
                    Throw ex
                Else
                    '' Non Autoclose e forse Nuovo tentativo
                    ''...
                End If
            End If
        Catch ex As Exception   ' Altri tipi di errore usciamo dal form
            Throw ex
        End Try

        ' Marchia in modo statico l'id della 
        ' Trasnazione ripresa dalla collegamento
        ' con il dispositivo hardware corrente.
        m_Transaction_Identifier = RefTo_Transaction_Identifier

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' L'Iniziailizzazione deve essere chiamata una volta sola nel contesto
            ' della sessione in corso, se la sessione sarà conlcusa sarà  chiamata
            ' nuovamente.
            FLAG_STATIC_INITIALIZATED = True

            ' ** OK --> INIZIALIZZATA e corretamente chiamata ad Hardware Argentea
            LOG_Debug(getLocationString(funcName), "Inizialization hardware successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            m_LastStatus = GLB_FAILED_POS_HARDWARE
            m_LastErrorMessage = "Inizializzazione fallita per KO hardware with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            ' ** KO --> Non inizializzata da parte di Argentea per errore hardware in risposta a questo codice.
            LOG_Debug(getLocationString(funcName), "Inizialization hardware failed with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

    ''' <summary>
    '''     Inizializza la Sessione verso il Dispositivo 
    '''     Pos di Argentea (Monetica) per la Carta Ticket.
    '''     Aspetta da utente un insieme di operazioni di storno
    ''' </summary>
    ''' <returns>True o False</returns>
    Private Function CallMultipleVoidOnPosHardware(_funcName As String) As Boolean
        Dim actApiCall As enApiToCall = enApiToCall.None
        Dim funcName As String = "CallMultipleVoidOnPosHardware"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = String.Empty

        ' Se nel contesto della sessione per questo scontrino
        ' è già stato iniziato un nuovo Ticket non lo richiediamo.
        'If FLAG_STATIC_INITIALIZATED Then
        'Return True
        'End If

        ' Status Corrente
        CallMultipleVoidOnPosHardware = False
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        '
        '   OUT
        '   L'Id di transazione  recuperata
        '   dopo le chiamate verso Argentea.
        '   (Passato alla dll COM di Argentea e fillato dalla stessa)
        '
        Dim RefTo_Transaction_Identifier As String = String.Empty

        ' Prima operazione di Avvio per il demat
        actApiCall = enApiToCall.MultipleVoids
        metdName = "VoidBPE"

#If DEBUG_SERVICE = 0 Then

        ' (Idle)
        retCode = ArgenteaCOMObject.VoidBPE(
                CInt(m_VoidableAmount * m_ParseFractMode),
                    RefTo_Transaction_Identifier,
                    RefTo_MessageOut
                )
#Else

        ''' Per Test
        System.Threading.Thread.Sleep(500)
        'System.Windows.Forms.Application.DoEvents()
        ''' 
        RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;4|2|1020|1|720|1|414;104;PELLEGRINI;  STORNO PAGAMENTI BUONO PASTO " ' <-- x test 
        'RefTo_MessageOut = "KO;    DATI NON RICEVUTI    ;;;;" ' <-- x test 
        '''
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
        ''' 
#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' (With Entrap) Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        ' in questo caso gestiamo l'eccezione per rimanere nel Form corrente
        Try
            m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)
        Catch ex As ExceptionProxyArgentea
            If ex.ErrorComunication Or ex.ErrorActionArgentea Then
                If frmEmulation.AutoCloseOnCompleteOperation Then
                    Throw ex
                Else
                    '' Non Autoclose e forse Nuovo tentativo
                    ''...
                End If
            End If
        Catch ex As Exception   ' Altri tipi di errore usciamo dal form
            Throw ex
        End Try

        ' Marchia in modo statico l'id della 
        ' Trasnazione ripresa dalla collegamento
        ' con il dispositivo hardware corrente.
        m_Transaction_Identifier = RefTo_Transaction_Identifier

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' L'Iniziailizzazione deve essere chiamata una volta sola nel contesto
            ' della sessione in corso, se la sessione sarà conlcusa sarà  chiamata
            ' nuovamente.
            FLAG_STATIC_INITIALIZATED = True

            ' ** OK --> INIZIALIZZATA e corretamente chiamata ad Hardware Argentea
            LOG_Debug(getLocationString(funcName), "Inizialization hardware successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            m_LastStatus = GLB_FAILED_POS_HARDWARE
            m_LastErrorMessage = "Inizializzazione fallita per KO hardware with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            ' ** KO --> Non inizializzata da parte di Argentea per errore hardware in risposta a questo codice.
            LOG_Debug(getLocationString(funcName), "Inizialization hardware failed with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

    ''' <summary>
    '''     Nella Sessione richede al dispositivo
    '''     Pos di Argentea (Monetica) le info sulla Carta Ticket.
    '''     Aspetta da utente un azione di inserimento per le info sulla sua carta
    ''' </summary>
    ''' <returns>True o False</returns>
    Private Function CallInfoOnPosHardware(_funcName As String) As Boolean
        Dim actApiCall As enApiToCall = enApiToCall.None
        Dim funcName As String = "CallInfoOnPosHardware"
        Dim metdName As String = "n/d"

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = String.Empty

        ' Status Corrente
        CallInfoOnPosHardware = False
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' 
        ' Riformatto lo status degli errori
        ' conservando per log quello precedente
        ' fino alla fine della sessione.
        '
        m_LogErrors.Add(m_LogErrors.Count + 1, New tLogErr(m_LastStatus, m_LastErrorMessage, m_LastResponseRawArgentea))

        '
        ' Quindi reset per la chiamata corrente
        '
        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        '
        '   OUT
        '   L'Id di transazione  recuperata
        '   dopo le chiamate verso Argentea.
        '   (Passato alla dll COM di Argentea e fillato dalla stessa)
        '
        Dim RefTo_Transaction_Identifier As String = String.Empty

        ' Prima operazione di Avvio per il demat
        actApiCall = enApiToCall.InfoCardUser
        metdName = "BalanceBPE"

#If DEBUG_SERVICE = 0 Then

        ' (Idle)
        retCode = ArgenteaCOMObject.BalanceBPE(
                    RefTo_MessageOut
                )
#Else

        ''' Per Test
        System.Threading.Thread.Sleep(500)
        'System.Windows.Forms.Application.DoEvents()
        ''' 
        RefTo_MessageOut = "OK;OPERAZIONE ACCETTATA;4|2|1020|1|720|1|414;104;PELLEGRINI;  BUONI PASTO " ' <-- x test 
        'RefTo_MessageOut = "KO;    DATI NON RICEVUTI    ;;;;" ' <-- x test 
        '''
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
        ''' 
#End If

        ' ** Response Grezzo in debug
        LOG_Debug(funcName, "API: " & m_CurrentApiNameToCall & " Command: " & m_CommandToCall.ToString() & " Method: " & metdName & " retCode: " & retCode.ToString & ". actApiCall: " & actApiCall.ToString() & " Response Output: " & RefTo_MessageOut)

        ' (With Entrap) Riprendiamo la Risposta da protocollo Argentea (potrebbe sollevare eccezione di Comunication o Parsing)
        ' in questo caso gestiamo l'eccezione per rimanere nel Form corrente
        m_LastResponseRawArgentea = _ParseResponseAndMapToThisResult(funcName, metdName, actApiCall, retCode, RefTo_MessageOut)

        ' Marchia in modo statico l'id della 
        ' Trasnazione ripresa dalla collegamento
        ' con il dispositivo hardware corrente.
        m_Transaction_Identifier = RefTo_Transaction_Identifier

        ' Se Argentea mi dà Successo Procedo altrimenti 
        ' sono un un errore remoto, su eccezione locale
        ' di parsing esco a priori e non passo.
        If m_LastResponseRawArgentea.Successfull Then

            ' ** OK --> EFFETTUATA e corretamente chiamata ad Hardware Argentea
            LOG_Debug(getLocationString(funcName), "Operation hardware successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            m_LastStatus = GLB_FAILED_POS_HARDWARE
            m_LastErrorMessage = "Operazione fallita per KO hardware with - " & m_LastResponseRawArgentea.ErrorMessage & " - "

            ' ** KO --> Non inizializzata da parte di Argentea per errore hardware in risposta a questo codice.
            LOG_Debug(getLocationString(funcName), "Operation hardware failed with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

        End If

    End Function

#End Region

#Region "Functions per la gestione Exception e Error"

    ''' <summary>
    '''     Negli eventi utilizzare questo metodo per impostare
    '''     un eccezione non gestita o errori di cui si vuole che
    '''     il flow in corso sia interrotto regolarmente.
    ''' </summary>
    ''' <param name="funcname">Il nome della funzione che vuole gestire lo stato dell'errore</param>
    ''' <param name="ExternalStatus">Lo Status con cui si deve innescare l'eccezione</param>
    ''' <param name="errorMessageExternal">Il Messaggio di Errore da mostrare</param>
    Friend Sub SetStatusInError(funcname As String, ExternalStatus As String, errorMessageExternal As String, LevelMessage As PosDef.TARMessageTypes, Optional ForceMessageBox As Boolean = False)

        ' Signal (Riporta lo Status della chiamata esterna e il messaggio)
        _SetOperationStatus(funcname,
            ExternalStatus,
            errorMessageExternal,
            LevelMessage, ForceMessageBox
        )

        ' Internamente comunque definisco lo stato generale 
        ' del proxy in esecuzione  come stato di errore per 
        ' interrompere nel flow eventuali prosegqui.
        m_ServiceStatus = enProxyStatus.InError

    End Sub

    ''' <summary>
    '''     Gestisce per segnalarle le eccezioni sui Throw
    '''     arrivati dal proxy locale per comodo all'interpretazione.
    ''' </summary>
    ''' <param name="funcname">Il nome della funzione che sta gestendo il try</param>
    ''' <param name="ex">L'eccezione che è arrivata dalla funzione di throw</param>
    Private Sub SetExceptionsStatus(funcname As String, ex As Exception, Optional ForceShowMessage As Boolean = False)

        ' Impostiamo per restituire la risposta e lo stato
        ' del proxy secondo l'errore in corso. Se questo è
        ' un errore sollevato con eccezione di tipo ExceptionProxyArgentea
        If TypeOf (ex) Is ExceptionProxyArgentea Then

            Dim ProxyError As ExceptionProxyArgentea = CType(ex, ExceptionProxyArgentea)

            If ProxyError.methodName = "[LOCAL_ERROR]" Then

                ' ERRORI INTERNI GESTITI
                m_LastStatus = "LOCAL." & ProxyError.ErrorTarget

            Else

                ' ERRORI GLOBALI CLASSIFICATI
                m_LastStatus = "GENERAL." & ProxyError.ErrorTarget

            End If

            If ProxyError.ErrorTarget = "PARSE_FAILED" Then

                ' Riportiamo la descrizione più estesa
                m_LastErrorMessage = ProxyError.ErrorDescription & " --> " & ProxyError.RefTo_MessageOut

                ' Error su form emulato
                _SetOperationStatusForm(NOT_INFO_POS_ERROR, "POS Error Internal...", FormEmulationArgentea.InfoStatus.Error, 7)

            Else

                ' Riportiamo la descrizione più estesa
                m_LastErrorMessage = ProxyError.ErrorDescription

                ' Error su form emulato
                _SetOperationStatusForm(m_LastStatus, m_LastErrorMessage, FormEmulationArgentea.InfoStatus.Error, 7)
                System.Threading.Thread.Sleep(1000)
            End If

        Else

            ' Altro non previsto in questa funzione             *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            Dim ProxyError As ExceptionProxyArgentea = New ExceptionProxyArgentea(funcname, ExceptionProxyArgentea.LOC_ERROR_NOT_CLASSIFIED, "Errore non classificato e inatteso -- Exception UKNOWED --", ex)

            m_LastStatus = "UKNOWED." & ProxyError.ErrorTarget & "." & ProxyError.retCode
            m_LastErrorMessage = ProxyError.ErrorDescription & "> " & ProxyError.retCode & "<"

            ' Error su form emulato
            _SetOperationStatusForm(m_LastStatus, "POS Err Uknow > " & ProxyError.retCode, FormEmulationArgentea.InfoStatus.Error, 5)

        End If

        ' Se l'eccezione è a cascata di altre...
        If Not ex.InnerException Is Nothing Then
            LOG_Error(funcname, "Errore con exception interna :: " & m_LastStatus & " -- " & m_LastErrorMessage + " -- " & " -- " & ex.Message & " --" & ex.InnerException.Message)
        Else
            LOG_Error(funcname, "Errore gestito :: " & m_LastStatus & " -- " & m_LastErrorMessage & " -- " & ex.Message)
        End If

        ' LOCAL. GENERAL. UKNOWED.
        ' Signal (KO remoto su dematerializzazione + Status remoto per le codifiche da db personalizzate)
        _SetOperationStatus(funcname,
            NOT_INFO_ERROR_INTERNAL,
            m_LastErrorMessage,
            PosDef.TARMessageTypes.TPSTOP, ForceShowMessage  ' <-- Lo status remoto
        )

    End Sub

    ''' <summary>
    '''     Imposta e definisce lo stato corrente dell'operazione
    '''     per restituirlo in notifica al chiamante.
    '''     Visualizza o meno (SilentMode) il messaggio di uscita.
    ''' </summary>
    ''' <remarks>
    '''     Se constType e msgDefault sono passati a Nothing
    '''     Solo per la MsgBox eventuale da mostrare riprende
    '''     l'ultimo stato e l'ultimo messaggio senza reimpostare lo stato.
    ''' </remarks>
    ''' <param name="funcName">Il Nome della funzione da usare per il log</param>
    ''' <param name="constType">Lo status da Impostrae tra le costanti disponbili del modulo</param>
    ''' <param name="msgDefault">Il Messaggio di default per l'eventuale msgBox</param>
    ''' <param name="TypeStatusMsgBox">Il Tipo di msgbox per livello</param>
    ''' <param name="ForceShowMessage">Se mostrare comunque e sempre il Messaggio di Avviso in Cassa</param>
    ''' <param name="InfoExtraMessageStatus">Informazioni extra da accodare al Messaggio di errore</param>
    Private Sub _SetOperationStatus(funcName As String, constType As String, msgDefault As String, TypeStatusMsgBox As PosDef.TARMessageTypes, Optional ForceShowMessage As Boolean = False, Optional InfoExtraMessageStatus As String = "")

        Dim c_StatusMessage As String

        ' Inizializzazione dell'emulatore POS
        If constType = NOT_INFO_POS_INIT Or
           constType = NOT_INFO_POS_CALL Or
           constType = NOT_INFO_POS_ERROR Then

            ' Messagi di stato solo per il Form
            ' in emulazione del POS software.
            _SetOperationStatusForm(constType, msgDefault, TypeStatusMsgBox)
            Return

        End If


        ' Imposta l'ultimo stato  corrente  per 
        ' l'uso successivo alle funzioni di chi
        ' esce.
        If constType Is Nothing And msgDefault Is Nothing Then

            ' Per gli errori sconosciuti che arrivano forziamo la
            ' segnalazione in uscita (per sistemarla in seguito)
            If m_LastStatus = GLB_FAILED_POS_HARDWARE Then

                ' In questo caso etichettiamo messaggi non interprteati come signal
                c_StatusMessage = "SIGNAL_" + m_LastStatus

            Else

                ' In questo caso non imposta l'ultimo stato (usato solo per rinotificare)
                c_StatusMessage = m_LastStatus

            End If

        Else

            ' Per questa tipologia codifichiamo la msgbox al fine
            ' di riprendere la costante dello stato remoto per poter
            ' personalizzare sul db i messaggi
            If constType = NOT_INFO_OPERATION_NOT_VALID_SPECIAL Then

                c_StatusMessage = "REMOTE_" + m_LastStatus
                m_LastStatus = constType
                m_LastErrorMessage = msgDefault

            Else

                ' Imposta l'ultimo stato prima di notificare
                m_LastStatus = constType
                c_StatusMessage = m_LastStatus
                m_LastErrorMessage = msgDefault

            End If

        End If

        ' Msg Utente    --> ** (Ultimo Status e ErrorMessage impostato dall'azione precedente)
        If Not m_SilentMode Or ForceShowMessage Then

            ' Scrive una riga di Log per aiutare l'operatore a individuare il messaggio da tradurre....
            If c_StatusMessage Is Nothing Then
                c_StatusMessage = "(NOT_CODIFICATED)"
                LOG_Error(getLocationString(funcName), m_LastErrorMessage + InfoExtraMessageStatus + " <-:: Voce Non Codificata e non Gestita ::-> " + "LevelITCommonModArgentea_" + c_StatusMessage)
            Else
                LOG_Error(getLocationString(funcName), m_LastErrorMessage + InfoExtraMessageStatus + " <-:: Voce DB x Tradurre ::-> " + "LevelITCommonModArgentea_" + c_StatusMessage)
            End If

            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage + InfoExtraMessageStatus, "LevelITCommonModArgentea_" + c_StatusMessage, TypeStatusMsgBox)

        Else

            ' Scrive una riga di Log per monitorare....
            LOG_Info(getLocationString(funcName), m_ServiceStatus.ToString() + " -> " + m_LastStatus + " -> " + m_LastErrorMessage + " " + InfoExtraMessageStatus)

        End If

        ' Status su Form
        If Not frmEmulation Is Nothing Then
            frmEmulation.SetStatus(PictureMultiStatusControlExpanse.enStatustype.Error)
        End If

    End Sub

    ''' <summary>
    '''     Per impostare il messaggio di stato sul form
    '''     dell'emulatore software in corso.
    ''' </summary>
    ''' <param name="constType"></param>
    ''' <param name="msgDefault"></param>
    ''' <param name="TypeStatusMsgBox"></param>
    Private Sub _SetOperationStatusForm(constType As String, msgDefault As String, TypeStatusMsgBox As FormEmulationArgentea.InfoStatus, Optional CloseTimeLaps As Integer = 0)

        ' Status su Form
        If Not frmEmulation Is Nothing Then
            If TypeStatusMsgBox = FormEmulationArgentea.InfoStatus.Flush Then
                System.Threading.Thread.Sleep(10)
            End If
            frmEmulation.SetMsgStatus(constType, msgDefault, TypeStatusMsgBox, CloseTimeLaps)
            'System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(100)
        End If

    End Sub

#End Region

#Region "Functions Common e Argentea Specifiche"

    ''' <summary>
    '''     Restituisce una stringa codificata ripresa dalla sessione 
    '''     Argentea secondo le sue specifiche di protocollo in input.
    ''' </summary>
    ''' <param name="TypeCodifique">La Stringa da inviare tramite la funzione specifica usata da una tra quelle valide da protocollo</param>
    ''' <param name="IDCrcTransactionSpecifique">Per il Reverse l'id della transazione da stornare è quello associato al momento della transazione stessa.</param>
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
    Public Function GetCodifiqueReceipt(TypeCodifique As TypeCodifiqueProtocol, Optional ByVal IDCrcTransactionSpecifique As String = "") As String
        Dim Result As String

        Result = ""
        m_RUPP = Me.GetPar_RUPP  ' Accertiamo il valore

        Select Case TypeCodifique
            Case TypeCodifiqueProtocol.Inizialization

                Result = "" ' Non Utilizzato

            Case TypeCodifiqueProtocol.Dematerialization

                ' -BPSW2$=" + Me.Get_CodeOperatorID() + "
                Result = "COUP_CODE=" + m_CurrentBarcodeScan + "-PROG=" _
                 + m_ProgressiveCall.ToString() + "-CASSA=" + m_RUPP + "-COD_DEV=" + Me.Get_CodeCashDevice +
                 "-COD_SCON=" + Me.Get_ReceiptNumber + "-RFU_1=-RFU_2=-"

            Case TypeCodifiqueProtocol.Reverse

                Result = "COUP_CODE=" + m_CurrentBarcodeScan + "-PROG=" _
                 + m_ProgressiveCall.ToString() + "-CASSA=" + m_RUPP + "-COD_DEV=" + Me.Get_CodeCashDevice +
                 "-COD_SCON=" + Me.Get_ReceiptNumber + "-TRAN_ID=" + IDCrcTransactionSpecifique + "-RFU_1=-RFU_2=-"

            Case TypeCodifiqueProtocol.Confirm

                Result = "PROG=" + m_ProgressiveCall.ToString() + "-CASSA=" + m_RUPP + "-COUP_CODE=" + m_CurrentBarcodeScan + "-TRAN_ID=" + m_LastCrcTransactionID

        End Select

        Return Result
    End Function


    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = Microsoft.VisualBasic.TypeName(Me) & "." & actMethode & " "
    End Function

    ''' <summary>
    '''     Mostra un Wait Screen rispetto
    '''     alle condizioni e opzioni  per
    '''     questo modulo applicativo
    ''' </summary>
    ''' <param name="Level">In base al livello identificato nel momento della funzione e rispetto alla opzione se visualizzare o meno la Wait Screen</param>
    ''' <param name="Msg1">Il Messaggio di Attesa</param>
    ''' <param name="Msg2">La seconda riga sul Messaggio di Attesa</param>
    Private Sub _ShowWaitScreen(Optional Level As Byte = 0, Optional Msg1 As String = Nothing, Optional Msg2 As String = Nothing)

#If DEBUG_SERVICE = 0 Then

        If m_TheModcntr Is Nothing Or frmEmulation Is Nothing Then
            Return
        End If

        If Level > 0 And Level >= m_OPT_ShowWaitScreenLevel Then
            FormHelper.ShowWaitScreen(m_TheModcntr, False, frmEmulation, Msg1, Msg2)
        End If

#End If

    End Sub

    ''' <summary>
    '''     Ripulisce lo schermo dal WaitScreen
    ''' </summary>
    Private Sub _ClearWaitScreen()
        If m_TheModcntr Is Nothing Or frmEmulation Is Nothing Then
            Return
        End If
        FormHelper.ShowWaitScreen(m_TheModcntr, True, frmEmulation)
    End Sub

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
        'Sub UpdateInitialValues(InitialTotalBPUsed As Integer, InitialTotalPayed As Decimal, InitialTotalVoided As Decimal, InitialTotalInvalid As Decimal)
        Function Contains(item As T) As Boolean
        Function GetEnumerator() As IEnumerator(Of T)
        Function IndexOf(item As T) As Integer

        ''' <summary>
        '''     Aggiunge un elemento all'insieme
        '''     e restituisce il riferimento.
        ''' </summary>
        ''' <param name="BarCode">Il Barcode da associare nell'elenco</param>
        ''' <param name="szValue">Il Valore rispettivo in formato stringa</param>
        ''' <param name="szFaceValue">Il Valore Facciale in formato stringa</param>
        ''' <param name="szEmitter">l'Emettitore circuito che ha emesso l'elemento</param>
        ''' <param name="CodeIssuer">Il Codice dell'emittente</param>
        ''' <param name="NameIssuer">Il Nome dell'emittente</param>
        ''' <param name="IdTransactionCrc">Il suo id di transazione sha effettutato sul server remoto</param>
        ''' <returns><see cref="T"/>L'elemento appena aggiunto</returns>
        Function NewPaid(ByVal BarCode As String, Optional Value As String = "", Optional FaceValue As String = "", Optional Emitter As String = "", Optional CodeIssuer As String = "", Optional NameIssuer As String = "", Optional ByVal IdTransactionCrc As String = "") As PaidEntry

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
        Private _InitialPaymentsInTA As Decimal = 0     ' <-- Il Totale della TA in ingresso 
        '
        Private _InitialTotalBPUsed As Integer = 0      ' <-- All'ingresso il conteggio dei BP già usati nelle sessioni di vendita precedenti
        Private _InitialTotalPayed As Decimal = 0       ' <-- All'ingresso il conteggio in valore già usato nelle sessioni di vendita precedenti
        Private _InitialTotalExcdeed As Decimal = 0     ' <-- All'ingresso se era presente nella TA un eccesso sul Totale (Inteso come resto)

        Private _InitialTotalBPVoided As Integer = 0    ' <-- All'ingresso il conteggio dei BP già usati nelle sessioni di vendita precedenti come storno
        Private _InitialTotalVoided As Decimal = 0      ' <-- All'ingresso il conteggio in valore già usato nelle sessioni di vendita precedenti come stornato

        Private _InitialTotalBPInvalid As Integer = 0   ' <-- All'ingresso il conteggio dei BP già usati nelle sessioni di vendita precedenti come non validi e non contabilizzati
        Private _InitialTotalInvalid As Decimal = 0     ' <-- All'ingresso il conteggio in valore già usato nelle sessioni di vendita precedenti come importi non contabilizzati non validi

        '
        '   Data list del risultato dei 
        '   Barcode scansionati o dal tot<le
        '   che il servizio argenetea ha dato.
        '
        Private m_ListEntries As ResultDataList(Of PaidEntry)

#End Region

#Region ".ctor"

        Public Sub New()

            ' Collection di risultati da riportare al consumer
            m_ListEntries = New ResultDataList(Of PaidEntry)()

            ' L'elemento a scrittura interna
            WriterResultDataList = m_ListEntries

        End Sub

        ''' <summary>
        '''     .ctor
        ''' </summary>
        ''' <param name="InitialPaymentsInTA">Il Totlae sulla TA prima dell'operazione</param>
        ''' <param name="InitialTotalBPUsed">Il Numero dei Buoni Usati Come Pagato</param>
        ''' <param name="InitialTotalPayed">Il Totale inteso come pagato</param>
        ''' <param name="InitialTotalExcdeed">Il Totale sulla TA che rappresenta l'eccesso du un Totale (Resto)</param>
        ''' <param name="InitialTotalBPVoided">Il Totale di elementi usati come storno</param>
        ''' <param name="InitialTotalVoided">Il Totale inteso come stornato</param>
        ''' <param name="InitialTotalBPInvalid">Il Totale di elementi non validi e non contabilizzati</param>
        ''' <param name="InitialTotalInvalid">Il Totale inteso come Non Valido</param>
        Public Sub New(InitialPaymentsInTA As Decimal,
                       InitialTotalBPUsed As Integer, InitialTotalPayed As Decimal, InitialTotalExcdeed As Decimal,
                       InitialTotalBPVoided As Integer, InitialTotalVoided As Decimal,
                       InitialTotalBPInvalid As Integer, InitialTotalInvalid As Decimal)
            Me.New()

            ' Riserbo le iniziali quelli alla chiamata Totali che serviranno al consumer
            _InitialPaymentsInTA = InitialPaymentsInTA

            _InitialTotalBPUsed = InitialTotalBPUsed
            _InitialTotalPayed = InitialTotalPayed
            _InitialTotalExcdeed = InitialTotalExcdeed

            _InitialTotalBPVoided = InitialTotalBPVoided
            _InitialTotalVoided = InitialTotalVoided

            _InitialTotalBPInvalid = InitialTotalBPInvalid
            _InitialTotalInvalid = InitialTotalInvalid

        End Sub

#End Region

#Region "Properties"

        ''' <summary>
        '''     Aggiorna per la Reponse
        '''     il Numero iniziale che
        '''     è relativo a prima di
        '''     aggiornare i valori di stato da un azione demat o void
        ''' </summary>
        ''' <param name="InitialPaymentsInTA">Il Totlae sulla TA prima dell'operazione</param>
        ''' <param name="InitialTotalBPUsed">Il Numero dei Buoni Usati Come Pagato</param>
        ''' <param name="InitialTotalPayed">Il Totale inteso come pagato</param>
        ''' <param name="InitialTotalExcdeed">Il Totale sulla TA che rappresenta l'eccesso du un Totale (Resto)</param>
        ''' <param name="InitialTotalBPVoided">Il Totale di elementi usati come storno</param>
        ''' <param name="InitialTotalVoided">Il Totale inteso come stornato</param>
        ''' <param name="InitialTotalBPInvalid">Il Totale di elementi non validi e non contabilizzati</param>
        ''' <param name="InitialTotalInvalid">Il Totale inteso come Non Valido</param>
        Public Sub UpdateInitialValues(InitialPaymentsInTA As Decimal, InitialTotalBPUsed As Integer, InitialTotalPayed As Decimal, InitialTotalExcdeed As Decimal,
                       InitialTotalBPVoided As Integer, InitialTotalVoided As Decimal,
                       InitialTotalBPInvalid As Integer, InitialTotalInvalid As Decimal)

            If Not InitialPaymentsInTA = Decimal.MinValue Then
                _InitialPaymentsInTA = InitialPaymentsInTA
            End If

            If Not InitialTotalBPUsed = Integer.MinValue Then
                _InitialTotalBPUsed = InitialTotalBPUsed
            End If
            If Not InitialTotalPayed = Decimal.MinValue Then
                _InitialTotalPayed = InitialTotalPayed
            End If
            If Not InitialTotalExcdeed = Decimal.MinValue Then
                _InitialTotalExcdeed = InitialTotalExcdeed
            End If

            If Not InitialTotalBPVoided = Integer.MinValue Then
                _InitialTotalBPVoided = InitialTotalBPVoided
            End If
            If Not InitialTotalVoided = Decimal.MinValue Then
                _InitialTotalVoided = InitialTotalVoided
            End If

            If Not InitialTotalBPInvalid = Integer.MinValue Then
                _InitialTotalBPInvalid = InitialTotalBPInvalid
            End If
            If Not InitialTotalInvalid = Decimal.MinValue Then
                _InitialTotalInvalid = InitialTotalInvalid
            End If

        End Sub

        ' *-*-*-*-

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
        '''     Numero totale di Buoni elaborati nella sessione
        '''     sul POS esterno durante le operazioni di raccolta
        '''     dati compresi quelli non validi (solo nella durata della sessione).
        ''' </summary>
        ''' <returns>Numerico Integer</returns>
        Protected Friend Overridable ReadOnly Property totalTAInCurrentDocument() As Decimal
            Get
                Return _InitialPaymentsInTA - m_TotalPayed_CS
            End Get
        End Property

        ''' <summary>
        '''     Numero totale di Buoni elaborati nella sessione
        '''     sul POS esterno durante le operazioni di raccolta
        '''     dati compresi quelli non validi (solo nella durata della sessione).
        ''' </summary>
        ''' <returns>Numerico Integer</returns>
        Protected Friend Overridable ReadOnly Property totalBPElaboratedInCurrentSession() As Integer
            Get
                Return m_TotalBPElaborated_CS
            End Get
        End Property

        ''' <summary>
        '''     Il totale ottenuto dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sulla vendita corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionPayedWithBP() As Decimal
            Get
                If _InitialTotalPayed = 0 Then
                    Return m_TotalPayed_CS
                Else
                    Return -Math.Abs(m_TotalVoided_CS - _InitialTotalVoided)
                End If
            End Get
        End Property

        ''' <summary>
        '''     Numero totale di Buoni utilizzati dalla sessione
        '''     sul POS esterno per pagare il Totale sulla vendita
        '''     corrente.
        ''' </summary>
        ''' <returns>Numerico Integer</returns>
        Protected Friend Overridable ReadOnly Property SessionBPUsedToPay() As Integer
            Get
                Return Math.Abs(_InitialTotalBPUsed - m_TotalBPUsedToPay_CS)
            End Get
        End Property

        ''' <summary>
        '''     Il totale in eccesso su dei buoni transitati 
        '''     nella sessione del POS sulla vendita corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionExcedeedWithBP() As Decimal
            Get
                Return m_TotalValueExcedeed_CS - _InitialTotalExcdeed
            End Get
        End Property

        ''' <summary>
        '''     Il valore totale stornato dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sullo storno corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionVoidedWithBP() As Decimal
            Get
                Return Math.Abs(_InitialTotalVoided - m_TotalVoided_CS)
            End Get
        End Property

        ''' <summary>
        '''     Il numero totale di taagli stornati dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sullo storno corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionBPUsedToVoid() As Decimal
            Get
                Return Math.Abs(_InitialTotalBPVoided - m_TotalBPUsedToVoid_CS)
            End Get
        End Property

        ''' <summary>
        '''     Il valore totale di bpe che non sono stati contabilizzati
        '''     in quanto nello storno non sono risultati  congrui
        '''     con qeulli che erano della transazione di demat.
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionNotContabilizated() As Decimal
            Get
                Return Math.Abs(_InitialTotalInvalid - m_TotalInvalid_CS)
            End Get
        End Property

        ''' <summary>
        '''     Il valore totale di bpe che non sono stati contabilizzati
        '''     in quanto nello storno non sono risultati  congrui
        '''     con qeulli che erano della transazione di demat.
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionValueExtraVoidNotContabilizated() As Decimal
            Get
                Return m_TotalValueExtraVoidNotContabilizated
            End Get
        End Property


        ''' <summary>
        '''     Il numero totale di bpe che non sono stati contabilizzati
        '''     in quanto nello storno non sono risultati  congrui
        '''     con qeulli che erano della transazione di demat.
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property SessionBPNotValid() As Decimal
            Get
                Return Math.Abs(_InitialTotalBPInvalid - m_TotalBPNotValid_CS)
            End Get
        End Property

        ' *-*-*-

        ''' <summary>
        '''     Il totale ottenuto dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sulla vendita corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalPayedWithBP() As Decimal
            Get
                Return m_TotalPayed_CS - m_TotalVoided_CS
            End Get
        End Property

        ''' <summary>
        '''     Numero totale di Buoni utilizzati dalla sessione
        '''     sul POS esterno per pagare il Totale sulla vendita
        '''     corrente.
        ''' </summary>
        ''' <returns>Numerico Integer</returns>
        Protected Friend Overridable ReadOnly Property totalBPUsedToPay() As Integer
            Get
                Return m_TotalBPUsedToPay_CS
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
        '''     Il valore totale stornato dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sullo storno corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalVoidedWithBP() As Decimal
            Get
                Return m_TotalVoided_CS
            End Get
        End Property

        ''' <summary>
        '''     Il numero totale di taagli stornati dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sullo storno corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalBPUsedToVoid() As Decimal
            Get
                Return m_TotalBPUsedToVoid_CS
            End Get
        End Property

        ''' <summary>
        '''     Il valore totale di bpe che non sono stati contabilizzati
        '''     in quanto nello storno non sono risultati  congrui
        '''     con qeulli che erano della transazione di demat.
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalNotContabilizated() As Decimal
            Get
                Return m_TotalInvalid_CS
            End Get
        End Property

        ''' <summary>
        '''     Il numero totale di bpe che non sono stati contabilizzati
        '''     in quanto nello storno non sono risultati  congrui
        '''     con qeulli che erano della transazione di demat.
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalBPNotValid() As Decimal
            Get
                Return m_TotalBPNotValid_CS
            End Get
        End Property


        ' *-*-*-

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
        ''' <param name="BarcodeToSearch">Un EAN usato come BP per partecipare al pagamento</param>
        ''' <returns></returns>
        Public Function ContainsBarcode(BarcodeToSearch As String, Optional CheckIfFlaggedRemoved As Boolean = False) As Boolean
            For Each itm As PaidEntry In m_ListEntries
                If itm.Barcode = BarcodeToSearch.Trim() Then
                    If CheckIfFlaggedRemoved Then
                        If itm.Voided = False Then
                            Return True
                        End If
                    Else
                        ' Esiste in elenco anche se etichetato come eliminato
                        Return True
                    End If
                End If
            Next
            Return False
        End Function

        ''' <summary>
        '''     Restituisce l'ID di transazione che ha avuto in risposta
        '''     al momento della dematerializzazione
        ''' </summary>
        ''' <param name="m_CurrentBarcodeScan">Il BarCode di cui si vuole l'id della sua transazione di conferma</param>
        ''' <returns></returns>
        Friend Function GetIDTransactionOfThisBarCode(BarcodeToSearch As String) As String
            For Each itm As PaidEntry In m_ListEntries
                If itm.Barcode = BarcodeToSearch.Trim() Then
                    Return itm.IDTransactionCrc
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
            Function CountElementsWithSomeFaceValue(faceValue As String, v As Boolean) As Integer
        End Interface

        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Protected Class ResultDataList(Of T)
            Implements IList(Of T)
            Implements IResultDataList(Of T)
            Implements IWResultDataList(Of T)
            Implements ICloneable

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
            ''' <param name="BarCode">Il Barcode da associare nell'elenco</param>
            ''' <param name="szValue">Il Valore rispettivo in formato stringa</param>
            ''' <param name="szFaceValue">Il Valore Facciale in formato stringa</param>
            ''' <param name="szEmitter">l'Emettitore circuito che ha emesso l'elemento</param>
            ''' <param name="CodeIssuer">Il Codice dell'emittente</param>
            ''' <param name="NameIssuer">Il Nome dell'emittente</param>
            ''' <param name="IdTransactionCrc">Il suo id di transazione sha effettutato sul server remoto</param>
            ''' <returns><see cref="T"/>L'elemento appena aggiunto</returns>
            Public Overridable Function NewPaid(ByVal BarCode As String, Optional Value As String = "", Optional FaceValue As String = "", Optional Emitter As String = "", Optional CodeIssuer As String = "", Optional NameIssuer As String = "", Optional ByVal IdTransactionCrc As String = "") As PaidEntry Implements IWResultDataList(Of T).NewPaid
                ' Metodo valido solo per i tipi PaidEntry
                ' si ptrebbe inseire a default(T) ma non ho tempo cassarola.
                Dim NewElement As Object = New PaidEntry(BarCode, Value, FaceValue, Emitter, IdTransactionCrc)
                CType(NewElement, PaidEntry).CodeIssuer = CodeIssuer
                CType(NewElement, PaidEntry).NameIssuer = NameIssuer
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

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim CC As ResultDataList(Of T) = New ResultDataList(Of T)
                'CC.CopyTo(Me._BackingStore, 0)
                For Each qItm As T In Me._BackingStore
                    CC.Add(qItm)
                Next
                Return CC
            End Function

            Public Function CountElementsWithSomeFaceValue(faceValue As String, Voided As Boolean) As Integer Implements IResultDataList(Of T).CountElementsWithSomeFaceValue
                Dim _CItm As Integer = 0
                Dim DecimalValue As Decimal = Convert.ToDecimal(faceValue)
                For Each qItm As Object In _BackingStore
                    If qItm Is Nothing Then
                        Exit For
                    End If
                    If CType(qItm, PaidEntry).DecimalValue = DecimalValue And
                            CType(qItm, PaidEntry).Voided = Voided And
                            CType(qItm, PaidEntry).Invalid = False Then               '<-- Solo quelli validi
                        _CItm += 1
                    End If

                Next
                Return _CItm
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

''' <summary>
'''     Exception dedicata al Proxy Argentea
''' </summary>
Friend Class ExceptionProxyArgentea : Inherits System.Exception

    ' *^*^*^*^*^*^*^*^*^*^*

    ' ECCEZIONI DI TIPO GLOBAL GENERAL PURPOSE

    ' Su Errore generale di comunciazione che ha
    ' restituito una chiamata ad un metodo della dll di Argentea.
    Public Const GLB_SOCKET_ERROR As String = "SOCKET_ERROR"
    Public Const GLB_MONETICA_ERROR As String = "MONETICA_ERROR"
    Public Const GLB_TIMEOUT_ERROR As String = "TIMEOUT_ERROR"
    Public Const GLB_SENDDATA_FAILED As String = "SENDDATA_FAILED"
    Public Const GLB_OPERATION_USERABORTED As String = "USER_ABORTED"
    Public Const GLB_TICKETCARD_NOTVALID As String = "TICKET_CARD_NOT_VALID"
    Public Const GLB_OPERATION_NOT_SUPP As String = "OPERATION_NOT_SUPPORTED"
    Public Const GLB_OPERATION_USERNOINPUTDATA As String = "USER_NOT_INPUT_DATA"

    ' Su Errori di Parsing durante la decodifica della
    ' Risposta o dell'Errore nella risposta dalla dll di Argentea.
    Public Const GLB_PARSE_EMPTY As String = "PARSE_EMPTY"
    Public Const GLB_PARSE_FAILED As String = "PARSE_FAILED"
    Public Const GLB_ERROR_ONPARSE As String = "ERROR_ONPARSE"


    ' ECCEZIONI DI TIPO LOCAL ALLA CLASSE

    ' Se nel fare il Connect allo STA della classe
    ' incorriamo in un Errore di tipo Double Connect()
    Friend Const LOC_PROXY_ALREADY_RUNNING As String = "ERROR-PROXY-ALREADY-RUN"

    ' Nell'istanziare e castare il Form di appoggio
    ' a scansione dei Barcode solleva questa eccezione.
    Friend Const LOC_ERROR_FORM_INSTANCE As String = "Error-INSTANCE-FORM"

    ' Nell'istanziare e castare il Form di appoggio
    ' per l'handler dell'evento non era un form valido.
    Friend Const LOC_ERROR_FORM_CAST As String = "Error-CAST-FORM"

    ' Nell'eseguire l'evento al chiamante per restituire i dati
    ' c'è stato un errore o un eccezione non gestita nell'evento.
    ' intercettiamo e proponiamo di controllare l'aggiornamento sul consumer.
    Friend Const LOC_ERROR_ON_EVENT_DATA As String = "Error-EVENT-DATA"

    ' Se un parametro di confgiruazione globale non è 
    ' stato configurato.
    Friend Const LOC_PAR_NOT_CONFIGURATED As String = "Error-EXCEPTION-PARNOTCONFIG"

    ' Eccezione nell'istanza di questa classe se  non  dovesse
    ' caricare per qualche motivo i parametri globali dedicati
    ' alla gestione di esecuzione del  servizio  corrente  con
    ' argentea.
    Friend Const LOC_ERROR_INSTANCE_PARAMETERS As String = "Error-ONLOAD-PARAMETERS-ARGENTEA"

    ' Se nel setException per ideintificare l'errore 
    ' arriva una eccezione non di tipo ExceptionArgentea
    Friend Const LOC_ERROR_NOT_CLASSIFIED As String = "Error-NOT-CLASSIFIED"

    ' Eccezione codificata
    Private m_LastResponseRawArgentea As ArgenteaFunctionReturnObject

    ' Se nell'eseguire il parsing sull'errore non c'era comunciazione
    Private m_ErrorComunication As Boolean = False

    ' Se nell'eseguire il parsing sull'errore da codificare
    Private m_ErrorOnParseProtocol As Boolean = False

    ' La codifica dell'errore è bloccante per l'iterazione con l'utente
    Private m_ErrorActionArgentea As Boolean = False

    ' Errore target ripreso dalle costanti predefinite
    Private m_ErrorTarget As String = String.Empty

    ' Errore target descritto per esteso o messaggi di ausilio
    Private m_ErrorDescription As String = String.Empty

    Friend ReadOnly Property ErrorComunication As Boolean
        Get
            Return m_ErrorComunication
        End Get
    End Property

    Friend ReadOnly Property ErrorOnParseProtocol As Boolean
        Get
            Return m_ErrorOnParseProtocol
        End Get
    End Property

    Friend ReadOnly Property ErrorActionArgentea As Boolean
        Get
            Return m_ErrorActionArgentea
        End Get
    End Property


    Friend ReadOnly Property ErrorTarget As String
        Get
            Return m_ErrorTarget
        End Get
    End Property

    Friend ReadOnly Property ErrorDescription As String
        Get
            Return m_ErrorDescription
        End Get
    End Property


    ''' <summary>
    '''     Il nome della funzione che ha sollevato l'eccezione
    ''' </summary>
    ''' <returns>Il nome della funzione che ha sollevato l'eccezione</returns>
    Friend ReadOnly Property funcName As String

    ''' <summary>
    '''     Il nome del metodo che si sta chiamando sulla dll di Argentea
    ''' </summary>
    ''' <returns>Il nome del metodo che si sta chiamando sulla dll di Argentea</returns>
    Friend ReadOnly Property methodName As String

    ''' <summary>
    '''     L'Api chiamata internamente per eseguire il metodo (detta le specifiche del Parsing per la Risposta)
    ''' </summary>
    ''' <returns>l'Api chiamata per la codifica del protocolo durante la fase di parsing internta tra quelle possibili <see cref="ClsProxyArgentea.enApiToCall"/>.</returns>
    Friend ReadOnly Property ApiCalled As ClsProxyArgentea.enApiToCall

    ''' <summary>
    '''     Il retCode restituito al momento della Chiamata che stabilisce se l'errore arriva dalla chiamata stessa (Errore di comunciazione, uknowed interno alla dll)
    ''' </summary>
    ''' <returns>Il retCode restituito al momento della Chiamata al metodo della dll che restituisce o 0 o un id di errore interno (comunicazione o uknowed)</returns>
    Friend ReadOnly Property retCode As Integer

    ''' <summary>
    '''     Il Messaggio passato alla funzione interna della dll di Argentea che è stato utilizzato per Mappare e Nattare su questa eccezione l'errore o il messaggio codificato per l'OK o il KO della risposta remota.
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property RefTo_MessageOut As String

    Friend ReadOnly Property LastResponseRawArgentea As ArgenteaFunctionReturnObject
        Get
            Return m_LastResponseRawArgentea
        End Get
    End Property

    ''' <summary>
    '''     .ctor a <see cref="ExceptionProxyArgentea"/>.
    ''' </summary>
    Private Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    '''     .ctor b <see cref="ExceptionProxyArgentea"/>.
    ''' </summary>
    ''' <param name="message">Messaggio di errore.</param>
    Private Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Per nuove istanze di <see cref="ExceptionProxyArgentea"/> con
    ''' una inner exception da riportare.
    ''' </summary>
    ''' <param name="message">Messaggio di errore.</param>
    ''' <param name="innerException">Eccezione da riportare</param>
    Private Sub New(ByVal message As String, ByVal innerException As System.Exception)
        MyBase.New(message, innerException)
    End Sub


    ''' <summary>
    '''     Tipo di Eccezione sollevata a causa di  un  errore
    '''     aulla chiamata della dll di argentea dove l'errore
    '''     è dovuto a un errore interno della  dll  o  di mancata
    '''     comunicazione, quindi un errore sollevato dalla stessa
    '''     prima che comunicasse con il servizio remoto.
    ''' </summary>
    ''' <param name="func_Name">Il nome della funzione del Proxy di Argentea che ha sollevato questa eccezione</param>
    ''' <param name="Method_Name">Il nome del metodo sulla dll di Argentea da cui si sta ricevendo la response</param>
    ''' <param name="Api_Called">Il Nome della API per la chiamata </param>
    ''' <param name="ret_Code">Il returno code che ha restituito la dll all'uscita</param>
    ''' <param name="Ref_MessageOut">Il messaggio dalla dll di argentea che è stato restituito grezzo</param>
    ''' <param name="ResponseCodeficated">Tupla restituita dalla funzione di Parsing in decodifica alla risposta di Argentea</param>
    Public Sub New(func_Name As String, Method_Name As String, Api_Called As ClsProxyArgentea.enApiToCall,
                    ret_Code As Integer, Ref_MessageOut As String,
                    ByRef ResponseCodeficated As Tuple(Of Boolean, Boolean, Boolean, String, String, ArgenteaFunctionReturnObject),
                    Optional ByVal innerException As System.Exception = Nothing)
        MyBase.New(func_Name & "." & Method_Name & "." & Api_Called.ToString() & "." & CStr(ret_Code), innerException)

        m_ErrorComunication = ResponseCodeficated.Item1                     ' Errore di Comunicazione SI/NO
        m_ErrorOnParseProtocol = ResponseCodeficated.Item2                  ' Errore di Parsing SI/NO
        m_ErrorActionArgentea = ResponseCodeficated.Item3                   ' Errore Azione Utente su Argentea
        m_ErrorTarget = ResponseCodeficated.Item4                           ' Errore Target esrpresso come costante
        m_ErrorDescription = ResponseCodeficated.Item5                      ' Descrizione dell'errore in modo esteso
        m_LastResponseRawArgentea = ResponseCodeficated.Item6               ' Error Response di Argentea
        '
        funcName = func_Name                                                ' Il Nome della funzione che ha sollevato questa eccezione
        methodName = Method_Name                                            ' Il nome del metodo della classe Proxy che si sta eseguendo
        ApiCalled = Api_Called                                              ' Il nome della API della dll in chiamata
        retCode = ret_Code                                                  ' Il RetCode della dll COM in risposta
        RefTo_MessageOut = Ref_MessageOut                                   ' Il Messaggio RAW della risposta sulla dll in COM

        ' ** KO --> Exception su Errori di comunicazione o per risposta remota data da Argentea KO.
        LOG_Error(func_Name, "Exception on .:  " & func_Name & " for Api to Call .: " & Api_Called.ToString() & " to Method Argentea .: " & Method_Name & " in response receive retCode .: " & CStr(ret_Code) & " with raw Message out .: " & Ref_MessageOut)

    End Sub

    ''' <summary>
    '''     Tipo di Eccezione sollevata a causa di  un  errore
    '''     aulla chiamata della dll di argentea dove l'errore
    '''     è stato innescato dal Throw dentro la funzione interna.
    ''' </summary>
    ''' <param name="func_Name">Il nome della funzione del Proxy di Argentea che ha sollevato questa eccezione</param>
    ''' <param name="Error_TargetName">Il nome ripreso dalle costanti LOC del modulo corrente che definiscono quale è la natura dell'errore che si sta solllevando come eccezione per il throw</param>
    ''' <param name="Error_Description">Una descrizione estesa dell'errore e della natura per cui è identificato come Exception per il modulo</param>
    Public Sub New(func_Name As String, Error_TargetName As String, Error_Description As String, Optional ByVal innerException As System.Exception = Nothing)
        MyBase.New(func_Name & "." & Error_TargetName, innerException)

        funcName = func_Name
        methodName = "[LOCAL_ERROR]"                              ''' **** Importante per la classificazione
        ApiCalled = ClsProxyArgentea.enApiToCall.None
        If Not innerException Is Nothing Then
            retCode = 9030  ' Senza Eccezione interna   (già classificato)
        Else
            retCode = 9031  ' Con Eccezione interna     (non classificato)
        End If

        RefTo_MessageOut = String.Empty
        '
        m_ErrorTarget = Error_TargetName
        m_ErrorDescription = Error_Description

        ' ** KO --> Exception su Errori di comunicazione o per risposta remota data da Argentea KO.
        If Not innerException Is Nothing Then
            LOG_Error(func_Name, "Exception .:  " & methodName & " for Internal Throw .: " & m_ErrorTarget & " with message response .: " & Error_Description & " -- " & innerException.Message)
        Else
            LOG_Error(func_Name, "Exception .:  " & methodName & " for Internal Throw .: " & m_ErrorTarget & " with message response .: " & Error_Description)
        End If

        ' Su eccezione locale interna disponiamo
        ' la last response come nuklla.
        m_LastResponseRawArgentea = Nothing

    End Sub

#Region "Function Shared Parsing"

    ''' <summary>
    '''     Sugli OK o KO di Argentea eseguo il Parsing  della  Risposta
    '''     per formulare il success o l'unseccessfull  con il messaggio 
    '''     ripreso dalla codifica del protocollo come  risposta. Sempre
    '''     nella chiamata attraverso il retCode stabiliamo se in errore
    '''     generale di Parsing o Comunicazione o Non previsto nella dll.
    ''' </summary>
    ''' <param name="ApiCalled">La chiamata Dove e per quale si sta chiamando <see cref="enApiToCall"/>.</param>
    ''' <param name="RetCode">La risposta ricevuta dalla chiamata al metodo nella dll di Argentea definita come <see cref="ArgenteaFunctionsReturnCode"/></param>
    ''' <param name="RefTo_MessageOut">L'OUT della chiamata al metodo dove si riceve la risposta dal servizio da decodificare secondo il protocollo inviato in CSV dal servizio, (OK/KO)</param>
    ''' <param name="FuncName">Il Nome della funzione da cui si sta eseguendo il Parser corrente per i log</param>
    ''' <param name="MethodName">Il Nome del Metodo della dll di Argente che si sta eseguendo come Chiamata API usato per la scrittura su log</param>
    ''' <returns>
    '''     Una Tupla con i seguenti Risultati.
    '''     Item1 = Boolean = ErrorComunication --> Definisce se l'errore è generale di comunicazione verso il servizio
    '''     Item2 = Boolean = ErrorParsing      --> Definisce se l'errore è generale di parsing sul protocollo che stiamo cercando di analaizzare
    '''     Item3 = String  = ErrorTarget       --> Riporta la costante predefinita per l'errore parsato
    '''     Item4 = String  = ErrorDescription  --> Riporta una descrizione estesa dell'errore parsato
    '''     Item5 = Object  = ObjectParsed      --> L'oggetto mappato su <see cref="ArgenteaFunctionReturnObject"/> con l'evidenza della comunicazione per l'OK o il KO e tutti i suoi attributi!!
    ''' </returns>
    Friend Shared Function ParseProtocolForMapResponse(
            ApiCalled As ClsProxyArgentea.enApiToCall,
            RetCode As ArgenteaFunctionsReturnCode,
            RefTo_MessageOut As String,
            FuncName As String,
            MethodName As String
        ) As Tuple(Of Boolean, Boolean, Boolean, String, String, ArgenteaFunctionReturnObject)

        Dim _ErrorComunication As Boolean = False
        Dim _ErrorOnParseProtocol As Boolean = False
        Dim _ErrorSystemArgentea As Boolean = False
        Dim _ErrorTarget As String = String.Empty
        Dim _ErrorDescription As String = String.Empty
        Dim ResultResponse As ArgenteaFunctionReturnObject

        Try

            ' Tipo di codifica generalizzata Argentea wrappatra su un ReturnObject
            Dim objTPTAHelperArgentea(0) As ArgenteaFunctionReturnObject
            objTPTAHelperArgentea(0) = New ArgenteaFunctionReturnObject()

            ' ---------------------------------------------------------------------------------------
            '
            '   ERRORI CLASSIFICATI SU RISPOSTA E NATTATI COME ERRORI INTERPERETABILI BLOCCANTI     '
            '
            ' ---------------------------------------------------------------------------------------
            If RefTo_MessageOut = "ERRORE SOCKET" Or                            ' <-- Questo Arriva dalla tentata comunicazione con il Service Remoto
                RefTo_MessageOut.ToUpper().Trim().EndsWith("ERRORE SOCKET") Then       ' <-- Questo arriva dalla tentata comunicazione con il POS Hardware

                ' ** KO --> Codificato Errore Socket 9001
                ResultResponse = New ArgenteaFunctionReturnObject(9001)
                _ErrorTarget = ExceptionProxyArgentea.GLB_SOCKET_ERROR
                _ErrorDescription = "-SOCKET ERROR"

            ElseIf RefTo_MessageOut.ToUpper().Trim().EndsWith("FALLITO Rs232;;;;") Then

                ' ** KO --> Codificato Errore di Configurazione Monetica Ini 9002
                ResultResponse = New ArgenteaFunctionReturnObject(9002)
                _ErrorTarget = ExceptionProxyArgentea.GLB_MONETICA_ERROR
                _ErrorDescription = "-MONETICA ERROR CONFIG"

            ElseIf RefTo_MessageOut.ToUpper().Trim().EndsWith("ERRORE TIMEOUT;") Then              ' KO;Errore timeout;;104;PELLEGRINI;Errore timeout;

                ' ** KO --> Codificato Errore di Timeout richiesta su Transazione 
                ResultResponse = New ArgenteaFunctionReturnObject(9003)
                _ErrorTarget = ExceptionProxyArgentea.GLB_TIMEOUT_ERROR
                _ErrorDescription = "-POS ERROR TIMEOUT"

            ElseIf RefTo_MessageOut.ToUpper().Trim().StartsWith("KO;INVIO DATI FALLITO") Then      ' KO;Invio Dati Fallito ETHERNET;;104;PELLEGRINI;Errore Invio Dati;

                ' ** KO --> Codificato Errore Invio Dati su Transazione POS
                ResultResponse = New ArgenteaFunctionReturnObject(9004)
                _ErrorTarget = ExceptionProxyArgentea.GLB_SENDDATA_FAILED
                _ErrorDescription = "-POS ERROR SEND DATA FAILED"

            ElseIf RefTo_MessageOut.ToUpper().Trim().EndsWith("OPERAZIONE ANNULLATA;") Then        ' KO;Operazione annullata;;000;;Operazione annullata; 

                ' ** KO --> Codificato Errore Operazione annullata da Utente 9005
                ResultResponse = New ArgenteaFunctionReturnObject(9005)
                _ErrorTarget = ExceptionProxyArgentea.GLB_OPERATION_USERABORTED
                _ErrorDescription = "-POS OPERATION ABORTED BY USER"

            ElseIf RefTo_MessageOut.ToUpper().Trim().StartsWith("KO;NESSUN BUONO") Then            ' KO;NESSUN BUONO SELEZIONATO;;104;PELLEGRINI;NESSUN BUONO SELEZIONATO;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  

                ' ** KO --> Codificato Errore Nessun Buono selezionato da Utente 9006
                ResultResponse = New ArgenteaFunctionReturnObject(9006)
                _ErrorTarget = ExceptionProxyArgentea.GLB_OPERATION_USERNOINPUTDATA
                _ErrorDescription = "-POS OPERATION NO INPUT DATA BY USER"

            ElseIf RefTo_MessageOut.ToUpper().Trim().StartsWith("KO;CARTA NON GESTITA") Then       ' KO;CARTA NON GESTITA;;;;NON GESTITA;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  

                ' ** KO --> Codificato Errore Carta non gestita 9007
                ResultResponse = New ArgenteaFunctionReturnObject(9007)
                _ErrorTarget = ExceptionProxyArgentea.GLB_TICKETCARD_NOTVALID
                _ErrorDescription = "-POS OPERATION TICKET CARD NOT VALID"

            ElseIf RefTo_MessageOut.ToUpper().Trim().StartsWith("KO;OPERAZIONE NON SUPPORTATA") Then ' KO;OPERAZIONE NON SUPPORTATA;;;;;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  

                ' ** KO --> Codificato Errore Operazione Non Supportata 9008
                ResultResponse = New ArgenteaFunctionReturnObject(9008)
                _ErrorTarget = ExceptionProxyArgentea.GLB_OPERATION_NOT_SUPP
                _ErrorDescription = "-POS OPERATION NOT IMPLEMENTATED"

            ElseIf RefTo_MessageOut.ToUpper().Trim().StartsWith("KO;    DATI NON RICEVUTI") Then     ' KO;  Dati non ricevuti;;;;;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
                ' ** KO --> Codificato Errore Operazione Non Supportata 9009
                ResultResponse = New ArgenteaFunctionReturnObject(9009)
                _ErrorTarget = ExceptionProxyArgentea.GLB_OPERATION_USERNOINPUTDATA
                _ErrorDescription = "-POS OPERATION NO INPUT DATA RECIEVE"

            ElseIf RefTo_MessageOut Is Nothing Or (RefTo_MessageOut = String.Empty) Then

                ' ** KO --> Codificato Errore Parsing 9010
                ResultResponse = New ArgenteaFunctionReturnObject(9010)
                _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_EMPTY
                _ErrorDescription = "-PARSING ERROR EMPTY"

            Else

                ' Riprendiamo i tipi necessari alla formattazione del Protocollo rispetto alla chiamata che si sta facendo.
                Dim ParsingMode As Tuple(Of InternalArgenteaFunctionTypes, Char, Integer) = ClsProxyArgentea.GetSplitAndFormatModeForParsing(ApiCalled)

                ' Parsiamo la risposta argentea per l'azione
                If (Not CSVHelper.ParseReturnString(RefTo_MessageOut, ParsingMode.Item1, objTPTAHelperArgentea, ParsingMode.Item2, ParsingMode.Item3)) Then

                    ' ** KO --> Codificato Errore Parsing 9010
                    ResultResponse = New ArgenteaFunctionReturnObject(9010)
                    Dim _part As String
                    If RefTo_MessageOut.ToUpper().StartsWith("OK;") Or RefTo_MessageOut.ToUpper().StartsWith("OK-") Then
                        _part = "ACTION IN OK WITH --> "
                    ElseIf RefTo_MessageOut.ToUpper().StartsWith("KO;") Or RefTo_MessageOut.ToUpper().StartsWith("KO-") Then
                        _part = "ACTION IN KO WITH --> "
                    Else
                        _part = "ACTION UKNOWED WITH --> "
                    End If

                    _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_FAILED
                    _ErrorDescription = "-PARSING ERROR FAILED-:: " & _part & "::" & RefTo_MessageOut

                    ' ** KO --> Error Parsing Description
                    LOG_Error(FuncName, "Error Parsing Protocol  .: " & _ErrorDescription)

                Else

                    ' Risposta classifica correttamente (Successfully or not Successfully)

                    ' In OK da Argentea Message è valorizzato       
                    ' in KO da Argentea Error è valorizzato         ( Errori di Notifica non Bloccanti )

                    ' ** INFO --> Parsed Error correttamente alla risposta raw della chiamata ad Argentea
                    ResultResponse = objTPTAHelperArgentea(0)

                End If

            End If

        Catch ex As Exception

            ' ** KO --> Codificato Errore Parsing 9010
            ResultResponse = New ArgenteaFunctionReturnObject(9010)
            _ErrorTarget = ExceptionProxyArgentea.GLB_ERROR_ONPARSE  ' <-- PARSING ERROR EXECPETD LOCAL FUNCTION (se questo errore è solo in questa funzione)
            _ErrorDescription = ex.Message

        End Try

        ' LOG e Return
        If ResultResponse.Status = 9001 Then

            ' --> Su Errore di Comunicazione etichettiamo questa Exception per gestioni succesive da segnalare come errore di comunicazione.
            _ErrorComunication = True

            ' ** KO --> Exception sull'effettuare il Parsing Errori per mancata comunicazione con il sistema remoto Argentea.
            LOG_Error(FuncName, "Comunication Failed with Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out is ERROR SOCKET. CHECK lan to resolve!! ")


        ElseIf ResultResponse.Status >= 9002 And ResultResponse.Status <= 9009 Then

            ' --> Su Errori bloccanti dal sistema Argentea
            _ErrorSystemArgentea = True
            LOG_Error(FuncName, "Failed Operation Argentea for status not valid to complete operation .: " & ApiCalled.ToString() & " with Method Argentea .: " & MethodName & " in response received with retCode .: " & RetCode & " and raw Message out is Empty. Request User action to resolve!! ")

        ElseIf ResultResponse.Status = 9010 Then

            ' --> Su Errore di Parsing etichettiamo questa Exception per gestioni succesive da segnalare come errore di parsing sul protocollo.
            _ErrorOnParseProtocol = True

            ' --> Log dell'errore di Parsing
            If _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_EMPTY Then
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out is Empty. CHECK function to resolve!! ")
            ElseIf _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_FAILED Then
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out .: " & RefTo_MessageOut & " CHECK on errors parsing to resolve!! ")
            ElseIf _ErrorTarget = ExceptionProxyArgentea.GLB_ERROR_ONPARSE Then
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out .: " & RefTo_MessageOut & " CHECK on errors parsing to resolve!! ")
            Else ' Eccezione suq questa funzione in locale (Non dovrebbe mai succedere) GLB_ERROR_ONPARSE
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out is Empty. Message Exception .: " & _ErrorTarget)
            End If

        Else

            ' ** OK --> NAT di status sul Parsing della risposta codificato correttmanete per il Successfull o l'Unsuccessfull
            LOG_Info(FuncName, "Parsed Protocol Argentea for API called .: " & ApiCalled.ToString() & " in Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out .: " & RefTo_MessageOut & " codifquated!! ")

        End If

        ' NAT TO
        Return Tuple.Create(_ErrorComunication, _ErrorOnParseProtocol, _ErrorSystemArgentea, _ErrorTarget, _ErrorDescription, ResultResponse)

    End Function

#End Region

End Class
