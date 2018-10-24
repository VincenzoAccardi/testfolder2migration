Imports System
Imports ARGLIB = PAGAMENTOLib
Imports System.Collections
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos.EFT
Imports TPDotnet.Pos
Imports System.Windows.Forms

' 0=PRODUZIONE 1=TEST
#Const DEBUG_SERVICE = 0

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

    ' *** **** ****

    ' Su errore quando l'opzione di girare il resto in eccesso
    ' assegno la costante di operazione non valida ai fine di pagamento.
    Private Const GLB_OPT_ERROR_VALUE_EXCEDEED As String = "ERROR-OPTION-PAYABLE-WITH-REST"

    ' BarCode già utilizzato in precedenza evitiamo
    ' di richiamare argentea per il controllo
    Private Const GLB_OPT_ERROR_NUMEBP_EXCEDEED As String = "ERROR-OPTION-PAYABLE-NUMBP-EXCEDEED"

    ' BarCode già utilizzato in precedenza evitiamo
    ' di richiamare argentea per il controllo
    Private Const GLB_INFO_CODE_ALREADYINUSE As String = "Error-BARCODE-ALREADYINUSE"

    ' Quando l'importo delle righe è già stato raggiunto (con o 
    ' senza eccesso per eventuale ipotesi di resto) non procediamo.
    Private Const GLB_INFO_IMPORT_ALREADYCOMPLETED As String = "Error-IMPORT-ALREADY_COMPLETED"

    ' BarCode da rimuovere da quelli già scanditi in precedenza 
    ' non presente in elenco
    Private Const GLB_INFO_CODE_NOTPRESENT As String = "Error-BARCODE-NOTPRESENT"

    ' Su errori non bloccanti ma da segnalare all'operatore
    ' come stmapa scontrino non effettuata o altro usiamo questo.
    Private Const GLB_SIGNAL_OPERATOR As String = "Error-SIGNALS_GENERIC"

    ' Su chiamate a Nomi di Funzioni API remote
    ' che non sono gestite da questo proxy
    Private Const GLB_ERROR_API_NOT_VALID As String = "Error-API_NOT_VALID"

    ' In fase di chiamata al servizio remoto di Argentea per la
    ' chiamata API Demtarelializzazione mi restituire NOT VALID
    Private Const GLB_INFO_BP_NOT_VALID As String = "Error-BP_NOT_VALID"

    ' *'*'*'* SULLE CALL 

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
        '''     API per la corretta chiamata al servizio alla fine della Command di demat della dll concludendo sul service remoto la convalida
        ''' </summary>
        Close

    End Enum

    '
    ' Elementi necessari per il parsing
    ' sul protocollo previsto e di frazione
    ' per il risultato.
    '
    'Private m_ParseSplitMode As String = "-"
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
    Private m_CurrentTransactionID As String = String.Empty                     ' <-- La Transazione in corso sulla TA GUID 
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
    'Shared m_TotalVoidedExcedeed_CS As Decimal              ' <-- Il Totale in eccesso/difetto se l'opzione per accettare valori maggiori è abilitata in storno

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

    'Private m_CurrentTerminalID As String = Nothing         ' <-- In Pos Hardware identifica il POS usato in Software l'ID del WebService

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
    ''' <param name="CurrentPaymentsTotal">Il Totale pagato fino adesso prima di effettuare l'eggiornamento dai dati del proxy corrente.</param>
    Protected Friend Sub New(
                             ByRef theModCntr As ModCntr,
                             ByRef taobj As TA,
                             TypeBehavior As enTypeProxy,
                             ByVal CurrentTransactionID As String,
                             ByVal CurrentPaymentsTotal As Decimal
                             )

        Dim funcName As String = "ClsProxyArgentea.New"

        ' Tipo BEHAVIOR
        m_TypeProxy = TypeBehavior

        ' Dati fondamentali
        m_CurrentTransactionID = CurrentTransactionID       ' l'ID della Transazione GUID sulla TA
        m_CurrentPaymentsTotal = CurrentPaymentsTotal       ' l'amount da pagare
        m_LogErrors = New Dictionary(Of Integer, tLogErr)   ' Log degli stati di volta in volta sul corso della sessione

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

            ' Signal (come errore di stampa ma non bloccante)
            m_LastStatus = ExceptionProxyArgentea.LOC_ERROR_INSTANCE_PARAMETERS
            m_LastErrorMessage = "Non è stato possibile caricare i parametri applicativi per eseguire il servizio Argentea"

            ' Msg Utente    ("attenzione non sono riuscito a stampare la ricevuta ma la transazione è valida")
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPSTOP)

            ' Bloccante
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_INSTANCE_PARAMETERS, "Errore nel caricare i parametri applicativi per eseguire il servizio Argentea -- Parametri Argentea non presenti --")

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

            m_TotalVoided_CS = 0
            m_VoidableAmount = 0

            For Each itm As KeyValuePair(Of String, PaidEntry) In m_PrefillVoidable

                '
                ' Questo è un Elemento che va visualizzato
                ' sul form e immesso nel dataResult per 
                ' una corretta gestione in uscita.
                '
                ItemNew = itm.Value

                'Aggiungo al dataResult per il calcolo in
                ' uscita da usare per aggiornare la TA
                WriterResultDataList.Add(ItemNew)

                ' Se è un elemento solo di riporto da stornato
                ' in sessioni precedenti
                If Not ItemNew.Voided Then ' NON etichettato come --> VOIDED -> Stornato

                    ItemNew.Value = (CDec(ItemNew.Value) / m_ParseFractMode).ToString("###,##0.00")
                    ItemNew.FaceValue = (CDec(ItemNew.FaceValue) / m_ParseFractMode).ToString("###,##0.00")

                    m_VoidableAmount += CDec(ItemNew.Value)
                    m_TotalBPUsed_CS += 1

                    formTD.FormatControls()

                    ' Aggiungo l'elemento al controllo Griglia
                    formTD.PaidEntryBindingSource.Add(ItemNew)

                    formTD.Refresh()

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

                    ' SIGNAL
                    m_LastStatus = GLB_ERROR_API_NOT_VALID
                    m_LastErrorMessage = "API called not present on service argentea"

                    If Not m_SilentMode Then

                        'Msg Utente 
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)

                    End If

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
        _flagCallOnetimeResetIncrement = False
        m_ProgressiveCall = 0

        m_LastStatus = Nothing
        m_LastErrorMessage = Nothing
        m_LastResponseRawArgentea = Nothing

        ' Totalizzatori di sessione
        m_TotalPayed_CS = 0
        m_TotalVoided_CS = 0
        m_TotalValueExcedeed_CS = 0
        m_TotalBPUsed_CS = 0

        '
        m_CurrentTransactionID = String.Empty
        m_CurrentPaymentsTotal = 0
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

            ' ** OK --> API CLOSE corretamente chiamata ad Argentea
            LOG_Debug(getLocationString(funcName), "API " & m_CurrentApiNameToCall & " successfuly on response with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

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
    '''         Key = IdCrcTransatcion      Value = Un valore stringa identificante la transazione associata alla ripostasta da argentea a questo barcode quando è stato dematerilizzato
    ''' 
    ''' </param>
    ''' <returns><see cref="ArgenteaFunctionsReturnCode"/></returns>
    Private Function _API_SingleVoid(ParamArray Arguments() As KeyValuePair(Of String, Object)) As ArgenteaFunctionsReturnCode

        Dim sender As FormBuonoChiaro = New FormBuonoChiaro() ' Form fittizio


        ' 0 = "BarCode" 1 = "IdCrcTransatcion" 2 = "Value"
        Dim barcode As String = Arguments(0).Value
        Dim refToCrcTransactionId As String = Arguments(1).Value
        Dim TotValueBP As String = Arguments(2).Value

        _DataResponse = New DataResponse(1, TotValueBP, 0)

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
    Private Function CreateInstanceFormForEmulationOrWaitAction() As Form
        Dim funcName As String = "CreateInstanceFormForEmulationOrWaitAction"
        Dim frmForEmulation As FormBuonoChiaro

        Try

            ' Istanza del form di appoggio  ad uso 
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
                frmForEmulation.Payable = m_VoidableAmount.ToString("###,##0.00")

            End If


        Catch ex As Exception

            ' Sollevo l'eccezione
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_INSTANCE, "Errore nell'istanziare il form nell'emulatore -- Contattare Assistenza --", ex)

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

        ' Sul db per cambiare Le Label della Griglia (Tabella .: MessageText)
        '   szTextID                                    szContextID	szLanguageCode	lRowIndex	szTranslation	szCustomerName	szAdditionalInformation	lTechLayerAccessID	szLastUpdLocal
        '   POSLevelITCommonFormBuonoChiaroHeaderDesc   TPDotnet	it-IT	        0	        Valore	        IT.Common	    NULL	                1	                20170713123500
        '   POSLevelITCommonFormBuonoChiaroHeaderValue  TPDotnet	it-IT	        0	        Codice	        IT.Common	    NULL	                1	                20170713123500
        '
        ' Sul db per cambiare Le Label del Form (Tabella .: ControlText)
        '   szTextID        szContextID	                szLanguageCode	szTranslation	szToolTipTranslation	bUseToolTip	szCustomerName	szLastUpdLocal	lTechLayerAccessID
        '
        '   cmdOK           TPDotnet.Pos.FormBPDemat    it-IT	        Chiudi	        Chiudi	                0	        Base	        20181017121448	1
        '   lblBarcode      TPDotnet.Pos.FormBPDemat    it-IT	        Barcode	        Barcode	                0	        Base	        20181017121448	1
        '   lblPayable      TPDotnet.Pos.FormBPDemat    it-IT	        Pagabile	    Pagabile	            0	        Base	        20181017121448	1
        '   lblrest         TPDotnet.Pos.FormBPDemat    it-IT	        Resto	        Resto	                0	        Base	        20181017121448	1
        '   lblPaid         TPDotnet.Pos.FormBPDemat    it-IT	        Pagato	        Pagato	                0	        Base	        20181017121555	1
        '
        frmTo.cmdOK.Text = getTextFromControlText(m_TheModcntr.contxt, "cmdOK", "FormBPVoid")
        frmTo.lblBarcode.Text = getTextFromControlText(m_TheModcntr.contxt, "lblBarcode", "FormBPVoid")
        frmTo.lblRest.Text = getTextFromControlText(m_TheModcntr.contxt, "lblRest", "FormBPVoid")
        frmTo.lblPayable.Text = getTextFromControlText(m_TheModcntr.contxt, "lblPayable", "FormBPVoid")
        frmTo.lblPaid.Text = getTextFromControlText(m_TheModcntr.contxt, "lblPaid", "FormBPVoid")

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

        ' Sul db per cambiare Le Label della Griglia (Tabella .: MessageText)
        '   szTextID                                    szContextID	szLanguageCode	lRowIndex	szTranslation	szCustomerName	szAdditionalInformation	lTechLayerAccessID	szLastUpdLocal
        '   POSLevelITCommonFormBuonoChiaroHeaderDesc   TPDotnet	it-IT	        0	        Valore	        IT.Common	    NULL	                1	                20170713123500
        '   POSLevelITCommonFormBuonoChiaroHeaderValue  TPDotnet	it-IT	        0	        Codice	        IT.Common	    NULL	                1	                20170713123500
        '
        ' Sul db per cambiare Le Label del Form (Tabella .: ControlText)
        '   szTextID        szContextID	                szLanguageCode	szTranslation	szToolTipTranslation	bUseToolTip	szCustomerName	szLastUpdLocal	lTechLayerAccessID
        '
        '   cmdOK           TPDotnet.Pos.FormBPPayment  it-IT	        Chiudi	        Chiudi	                0	        Base	        20181017121448	1
        '   lblBarcode      TPDotnet.Pos.FormBPPayment  it-IT	        Barcode	        Barcode	                0	        Base	        20181017121448	1
        '   lblPayable      TPDotnet.Pos.FormBPPayment  it-IT	        Pagabile	    Pagabile	            0	        Base	        20181017121448	1
        '   lblPaid         TPDotnet.Pos.FormBPPayment  it-IT	        Pagato	        Pagato	                0	        Base	        20181017121555	1
        '

        frmTo.lblBarcode.Text = getTextFromControlText(m_TheModcntr.contxt, "cmdOK", "FormBPPayment")
        frmTo.lblBarcode.Text = getTextFromControlText(m_TheModcntr.contxt, "lblBarcode", "FormBPPayment")
        frmTo.lblPayable.Text = getTextFromControlText(m_TheModcntr.contxt, "lblPayable", "FormBPPayment")
        frmTo.lblPaid.Text = getTextFromControlText(m_TheModcntr.contxt, "lblPaid", "FormBPPayment")

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
                Case enApiToCall.SingleVoid
                    _ParsingMode = InternalArgenteaFunctionTypes.SingleVoid_BP
                Case enApiToCall.SinglePayment
                    _ParsingMode = InternalArgenteaFunctionTypes.SinglePaid_BP
                Case enApiToCall.MultiplePayments
                    _ParsingMode = InternalArgenteaFunctionTypes.MultiPaid_BP
                Case enApiToCall.MultipleVoids
                    _ParsingMode = InternalArgenteaFunctionTypes.MultiVoid_BP
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

        ' ** KO --> Su GENERAL Errore di comunicazione protocollo o interno della dll (potrebbe risolversi anche in eccezione sul parsing della decodifica dell'errore su com )
        If ret_Code <> ArgenteaFunctionsReturnCode.OK Then
            ' Eccezione di comunicazione socket o com (o eventualmente sul parsing della decodifica dell'errore su com)
            Throw New ExceptionProxyArgentea(func_Name, Method_Name, Api_Called, ret_Code, Ref_MessageOut)
        End If

        Dim Response As Tuple(Of Boolean, Boolean, String, String, ArgenteaFunctionReturnObject) = ClsProxyArgentea.ParseProtocolForMapResponse(
                Api_Called, ret_Code, Ref_MessageOut, func_Name, Method_Name
        )

        ' ** OK/KO --> Su GENERAL Errore di parsing sulla risposta nel protocollo  di risposta
        If Response.Item1 Or Response.Item2 Then
            ' Eccezione di parsing sulla risposta da codificare secondo il rptocollo argentea (riassegniamo con exception PARSING)
            Throw New ExceptionProxyArgentea(func_Name, Method_Name, Api_Called, ret_Code, Ref_MessageOut)
        End If

        ' ** ULTIMO CRC di risposta in risposta da argentea valido
        If Response.Item5.TerminalID <> String.Empty Then
            m_LastCrcTransactionID = Response.Item5.TerminalID
        End If

        ' ** OK --> Parsing dela risposta (OK/KO) effettuato e nattato con successo
        Return Response.Item5                                   ' Response Parsed Object

    End Function

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
        ) As Tuple(Of Boolean, Boolean, String, String, ArgenteaFunctionReturnObject)

        Dim _ErrorComunication As Boolean = False
        Dim _ErrorOnParseProtocol As Boolean = False
        Dim _ErrorTarget As String = String.Empty
        Dim _ErrorDescription As String = String.Empty
        Dim ResultResponse As ArgenteaFunctionReturnObject

        Try

            ' Tipo di codifica generalizzata Argentea wrappatra su un ReturnObject
            Dim objTPTAHelperArgentea(0) As ArgenteaFunctionReturnObject
            objTPTAHelperArgentea(0) = New ArgenteaFunctionReturnObject()

            ' Parsiamo la risposta argentea
            If RefTo_MessageOut = "ERRORE SOCKET" Then

                ' ** KO --> Codificato Errore Socket 9001
                ResultResponse = New ArgenteaFunctionReturnObject(9001)
                _ErrorTarget = ExceptionProxyArgentea.GLB_SOCKET_ERROR
                _ErrorDescription = "-SOCKET ERROR"

            ElseIf RefTo_MessageOut Is Nothing Or (RefTo_MessageOut = String.Empty) Then

                ' ** KO --> Codificato Errore Parsing 9002
                ResultResponse = New ArgenteaFunctionReturnObject(9002)
                _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_EMPTY
                _ErrorDescription = "-PARSING ERROR EMPTY"

            Else

                ' Riprendiamo i tipi necessari alla formattazione del Protocollo rispetto alla chiamata che si sta facendo.
                Dim ParsingMode As Tuple(Of InternalArgenteaFunctionTypes, Char, Integer) = ClsProxyArgentea.GetSplitAndFormatModeForParsing(ApiCalled)

                ' Parsiamo la risposta argentea per l'azione
                If (Not CSVHelper.ParseReturnString(RefTo_MessageOut, ParsingMode.Item1, objTPTAHelperArgentea, ParsingMode.Item2, ParsingMode.Item3)) Then

                    ' ** KO --> Codificato Errore Parsing 9002
                    ResultResponse = New ArgenteaFunctionReturnObject(9002)
                    _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_FAILED
                    _ErrorDescription = "-PARSING ERROR FAILED-:: " & RefTo_MessageOut

                Else

                    ' Definisco lo Status riportato

                    ' ** OK --> Parsed Error correttamente alla risposta raw della chiamata ad Argentea
                    ResultResponse = objTPTAHelperArgentea(0)

                End If

            End If

        Catch ex As Exception

            ' ** KO --> Codificato Errore Parsing 9002
            ResultResponse = New ArgenteaFunctionReturnObject(9002)
            _ErrorTarget = ExceptionProxyArgentea.GLB_ERROR_ONPARSE  ' <-- PARSING ERROR EXECPETD LOCAL FUNCTION (se questo errore è solo in questa funzione)
            _ErrorDescription = ex.Message

        End Try

        ' LOG e Return
        If ResultResponse.Status = 9001 Then

            ' --> Su Errore di Comunicazione etichettiamo questa Exception per gestioni succesive da segnalare come errore di comunicazione.
            _ErrorComunication = True

            ' ** KO --> Exception sull'effettuare il Parsing Errori per mancata comunicazione con il sistema remoto Argentea.
            LOG_Error(FuncName, "Comunication Failed with Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out is ERROR SOCKET. CHECK lan to resolve!! ")

        ElseIf ResultResponse.Status = 9002 Then

            ' --> Su Errore di Parsing etichettiamo questa Exception per gestioni succesive da segnalare come errore di parsing sul protocollo.
            _ErrorOnParseProtocol = True

            If _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_EMPTY Then

                ' ** KO --> Exception sull'effettuare il Parsing Errori remoti in risposta da Argentea come KO.
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out is Empty. CHECK function to resolve!! ")

            ElseIf _ErrorTarget = ExceptionProxyArgentea.GLB_PARSE_FAILED Then

                ' ** KO --> Exception sull'effettuare il Parsing Errori di comunicazione o per risposta remota data da Argentea come KO.
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out .: " & RefTo_MessageOut & " CHECK on errors parsing to resolve!! ")

            Else ' Eccezione suq questa funzione in locale (Non dovrebbe mai succedere) GLB_ERROR_ONPARSE

                ' ** KO --> Exception locale in questa funzione sull'effettuare il Parsing Errori di comunicazione o per risposta remota data da Argentea KO.
                LOG_Error(FuncName, "Parsing Failed On Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out is Empty. Message Exception .: " & _ErrorTarget)

            End If

        Else

            ' ** OK --> Parsing Errori di comunicazione o per risposta remota data da Argentea effettuato correttamente.
            LOG_Info(FuncName, "Parsing Protocol Argentea for API to call .: " & ApiCalled.ToString() & " to Method Argentea .: " & MethodName & " with response received retCode .: " & RetCode & " and raw Message out .: " & RefTo_MessageOut & " codifquated!! ")

        End If

        Return Tuple.Create(_ErrorComunication, _ErrorOnParseProtocol, _ErrorTarget, _ErrorDescription, ResultResponse)

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

                    ' SIGNAL
                    m_LastStatus = GLB_SIGNAL_OPERATOR
                    m_LastErrorMessage = getPosTxtNew(m_TheModcntr.contxt, "POSLevelITCommonPrinterFailed", 0)

                    ' Msg Utente    (attenzione non sono riuscito a stampare la ricevuta ma la transazione è valida)
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
        Dim excepted As Exception = Nothing

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

            ' Capture
            excepted = ex

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

                ' Capture
                excepted = ex

            Finally

                If Not excepted Is Nothing Then

                    '
                    ' Alla chiusura se uno degli eventi ha solevato 
                    ' un eccezione nel consumer lo catturiamo espondendo
                    ' il problema di mancato aggiornamento al chiamante.
                    '
                    SetExceptionsStatus(funcName, excepted)

                    ' Con ritorno a Status InError
                    m_ServiceStatus = enProxyStatus.InError

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "--" + excepted.InnerException.ToString())

                    ' Msg Utente    --> ** (Ultimo Status e ErrorMessage impostato dall'azione precedente)
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

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
        Dim excepted As Exception = Nothing

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

            ' Capture
            excepted = ex

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

                ' Capture
                excepted = ex

            Finally

                If Not excepted Is Nothing Then

                    '
                    ' Alla chiusura se uno degli eventi ha solevato 
                    ' un eccezione nel consumer lo catturiamo espondendo
                    ' il problema di mancato aggiornamento al chiamante.
                    '
                    SetExceptionsStatus(funcName, excepted)

                    ' Con Ritorno a Status InError
                    m_ServiceStatus = enProxyStatus.InError

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "--" + excepted.InnerException.ToString())

                    ' Msg Utente --> ** (Ultimo Status e ErrorMessage impostato dall'azione precedente)
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

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
        '    ->  KO su ADD BPC (CallDematerialize)      :::         KO-903 - PROGRESSIVO FUORI SEQUENZA-----0--- 
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

                    ' SIGNAL
                    m_LastStatus = GLB_OPT_ERROR_NUMEBP_EXCEDEED
                    m_LastErrorMessage = "Il numero di titoli di pagamento per questa vendita è stato superato!!"

                    ' Msg Utente
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)

                    Return

                End If

                ' Controllo se nell'elenco è già stato considerato il BarCode
                If _DataResponse.ContainsBarcode(m_CurrentBarcodeScan) Then

                    ' SIGNAL
                    m_LastStatus = GLB_INFO_CODE_ALREADYINUSE
                    m_LastErrorMessage = "Il barcode per questo titolo di pagamento è già stato usato per questa vendita!!"

                    ' Msg Utente
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

                    ' SIGNAL
                    m_LastStatus = GLB_INFO_IMPORT_ALREADYCOMPLETED
                    m_LastErrorMessage = "L'importo da pagare è già stato completato per questa vendita completare con inoltro!!"

                    ' Msg Utente
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

                                ' Log locale
                                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + "-- Transaction Dematerialize Argentea ::KO:: Excedeed")

                                ' SIGNAL
                                m_LastStatus = GLB_OPT_ERROR_VALUE_EXCEDEED
                                m_LastErrorMessage = "Il Valore del Titolo di Pagamento eccede il valore rispetto al totale (non è possibile dare resto)"

                                ' Msg Utente
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
                        m_TotalPayed_CS += m_CurrentValueOfBP
                        m_TotalBPUsed_CS += 1                         ' <-- Conteggio numero di bpc usati in local per ogni ingresso sulla vendita

                        Try
                            ' Riprendo il sender che p il Form
                            ' dove voglio aggiungere alla lista
                            ' l'n elemento appena validato.
                            formBC = TryCast(sender, FormBuonoChiaro)
                            If formBC Is Nothing Then

                                ' Sollevo l'eccezione
                                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

                            End If

                            ' Aggiungo l'elemento al controllo Griglia
                            formBC.PaidEntryBindingSource.Add(ItemNew) ' New PaidEntry(m_CurrentBarcodeScan, paidValue.ToString("###,##0.00"), faceValue.ToString("###,##0.00"), ""))

                            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
                            formBC.Paid = m_TotalPayed_CS.ToString("###,##0.00")

                        Catch ex As Exception

                            ' Sollevo l'eccezione
                            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --", ex)

                        End Try

                    Else

                        ' Errata Dematerializzione o Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                        ' Log locale
                        'LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Remote")

                        ' SIGNAL
                        'm_LastStatus = GLB_INFO_BP_NOT_VALID
                        'm_LastErrorMessage = "Non è stato possibile dematerializzare il BP presso il servizio Argentea!! " & m_LastResponseRawArgentea.ErrorMessage

                        ' Msg Utente  ( Ultimo Status e ErrorMessage impotato da Demat o Confirm )
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Remote")

                    ' Msg Utente  ( Ultimo Status e ErrorMessage impotato da tentativo di Demat senza richiesta di confirm)
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

            Else

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

        Catch ex As Exception

            FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

            ' Catturo per riporto l'eccezione interna o glbale di sistema
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

                    ' SIGNAL
                    m_LastStatus = GLB_INFO_CODE_NOTPRESENT
                    m_LastErrorMessage = "Il BP non è presente tra le scelte possibili!!"

                    ' Msg utente
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)

                    Return

                End If

                If Inizializated Then

                    FormHelper.ShowWaitScreen(m_TheModcntr, False, sender)

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
                            If formBC Is Nothing Then

                                ' Sollevo l'eccezione
                                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

                            End If

                            ' Sul Form rimuovo dalla griglia l'elemento
                            formBC.PaidEntryBindingSource.RemoveCurrent()

                            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
                            formBC.Paid = m_TotalPayed_CS.ToString("###,##0.00")

                        Catch ex As Exception

                            ' Sollevo l'eccezione
                            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --", ex)

                        End Try

                    Else

                        ' Errata Reverse per Dematerializzione o Reverse Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                        ' Log locale
                        LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Reverse Demat Argentea ::KO:: Local")

                        ' Msg Utente  ( Ultimo Status e ErrorMessage impotato da Void o Confirm )
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Reverse Demat Argentea ::KO:: Not Intializated")

                    ' Msg Utente  ( Ultimo Status e ErrorMessage impotato da Void senza Confirm )
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

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
    ''' <param name="UndoBPCForExcedeed">Per irutilizzare la funzione su opzioni di Eccesso su totale previo Resto valido non valido</param>
    Protected Overridable Sub BarcodeRemoveVoidHandler(ByRef sender As Object, ByVal barcode As String)
        Dim funcName As String = "BarcodeRemoveVoidHandler"

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
    '''     Completa l'aggiornamento dei dati sul form
    '''     se la void è andata a buon fine.
    '''     ( Usato per appoggio a BarcodeRemoveHandler flaggato per uscire prima)
    ''' </summary>
    ''' <param name="sender">Il form di storno</param>
    Private Sub _updateVoidedForm(ByRef sender As Object)
        Dim funcName As String = "_updateVoidedForm"
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
                itm.Voided = True  ' Etitchettato come --> VOIDED -> Stornato
                _revoke = True
                Exit For
            End If
        Next

        If Not _revoke Then

            ' SIGNAL
            m_LastStatus = GLB_INFO_CODE_NOTPRESENT
            m_LastErrorMessage = "Il BP non è presente come pagato!!"

            ' Msg Utente
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPINFORMATION)

            Return

        End If

        ' Se è stato rimosso correttametne procediamo

        m_TotalBPUsed_CS -= 1                         ' <-- Conteggio numero di bpc usati in local per ogni rimosso

        ' Per il Form in azione corrente mi
        ' aggiorno il Totale da Pagare rispetto a
        ' quelli già in elenco
        m_VoidableAmount -= m_CurrentValueOfBP
        m_TotalVoided_CS += m_CurrentValueOfBP

        ' Riportiamo aggiornato il form

        Try
            ' Riprendo il sender che p il Form
            ' dove voglio aggiungere alla lista
            ' l'n elemento appena validato.
            formBC = TryCast(sender, FormBuonoChiaro)
            If formBC Is Nothing Then

                ' Sollevo l'eccezione
                Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --")

            End If

            ' Sul Form rimuovo dalla griglia l'elemento
            For Each itm As PaidEntry In formBC.PaidEntryBindingSource
                If itm.Barcode = m_CurrentBarcodeScan Then
                    formBC.PaidEntryBindingSource.Remove(itm)
                    Exit For
                End If
            Next

            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
            formBC.Paid = m_TotalVoided_CS.ToString("###,##0.00")
            formBC.Payable = m_VoidableAmount.ToString("###,##0.00")

        Catch ex As Exception

            ' Sollevo l'eccezione
            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_FORM_CAST, "Errore nell'istanziare il form come Form compatibile per l'evento -- Contattare Assistenza --", ex)

        End Try


    End Sub

#End Region

#Region "Functions Private per Hardware Mode"

    ''' <summary>
    '''     Inizializza la Sessione verso Argentea
    '''     e parte la numerazione interna delle chiamate
    '''     da 1
    ''' </summary>
    Private Function CallHardwareWaitMode(_funcName As String) As Boolean
        Dim actApiCall As enApiToCall = enApiToCall.None
        Dim funcName As String = "CallHardwareWaitMode"
        Dim metdName As String = "n/d"

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
        ' (idle)
        ' Pagamento (Payment) in una sesione su un POS terminale Hardware
        ' Storno (Void) in una sessione su un POS terminale Hardware
        '
        '   amount = L'importo per avviare il POS a farsi pagare in BP l'importo dettato
        '

#If DEBUG_SERVICE = 0 Then

        ' (Idle)
        If m_CommandToCall = enCommandToCall.Payment Then
            
            metdName = "PaymentBPE"
            actApiCall = enApiToCall.MultiplePayments
            retCode = ArgenteaCOMObject.PaymentBPE(
                    CInt(m_PayableAmount * m_ParseFractMode),
                        RefTo_Transaction_Identifier,
                        RefTo_MessageOut
                    )

        ElseIf m_CommandToCall = enCommandToCall.Void Then

            metdName = "VoidBPE"
            actApiCall = enApiToCall.MultipleVoids
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
            metdName = "PaymentBPE"
            actApiCall = enApiToCall.MultiplePayments
            RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;2|5|1020|1|414;104;PELLEGRINI;  PAGAMENTO BUONO PASTO " ' <-- x test 
            '''
        ElseIf m_CommandToCall = enCommandToCall.Void Then
            ''' 
            metdName = "VoidBPE"
            actApiCall = enApiToCall.MultipleVoids
            RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;2|4|1020|1|414;104;PELLEGRINI;  PAGAMENTO BUONO PASTO " ' <-- x test 
            ''''
        End If
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
            Next

            ' ** OK --> ATTESA COMPLETATA e corretamente chiamata vs Hardware Terminal POS
            LOG_Debug(getLocationString(funcName), "BP comunication with terminal pos successfuly on call first with message " & m_LastResponseRawArgentea.SuccessMessage)
            Return True

        Else

            ' ** KO --> Non inizializzata da parte di Argentea per errore remoto in risposta a questo codice.
            LOG_Debug(getLocationString(funcName), "BPE comunication remote failed on first call to terminal argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
            Return False

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

                m_TotalPayed_CS = m_VoidableAmount   ' <-- Che lo riporterà a _DataResponse.totalPayedWithBP  come differenza quello pagato e quello stornato
                RaiseEvent Event_ProxyCollectDataVoidedAtEnd(Me, _DataResponse)

            End If

        Catch ex As Exception

            ' Intercettiamo l'errore per il  contesto  probabilmente
            ' erchè il consumer non l'ha fatto per suo conto, quindi
            ' rimane che per noi il consumer con i dati non è aggioranto.

            Throw New ExceptionProxyArgentea(funcName, ExceptionProxyArgentea.LOC_ERROR_ON_EVENT_DATA, "Errore nell'evento durante il collect dei dati al consumer del Proxy -- Consumer in errore --", ex)

        End Try

    End Sub

#End Region

#Region "Functions Private per Emulation Pos Software Service mode"

    ''' :: Flag per il tentativo di Reset del Counter remoto presso Argentea
    Private _flagCallOnetimeResetIncrement As Boolean = False
    ''' :: Log durante delle chiamate di ogni errore o stato precedente
    Private m_LogErrors As Dictionary(Of Integer, tLogErr)

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

#Else

        ''' Per Test
        If Not _flagCallOnetimeResetIncrement Then
            ' 1° Tentativo
            RefTo_MessageOut = "KO-903-PROGRESSIVO FUORI SEQUENZA-----0---"            ' <-- x test su questo signal
            RefTo_MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 
        Else
            ' 2° tenttivo
            RefTo_MessageOut = "KO-903-ALTRO ERRORE-----0---"            ' <-- x test su questo signal
            RefTo_MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 

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
        RefTo_MessageOut = "OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--"    ' <-- x test 
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
                RefTo_MessageOut = "KO-0 - BUONO GIa' STORNATO -68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
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
                m_LastErrorMessage = "Reset Counter remote for BP KO remote failed with message - " & m_LastResponseRawArgentea.ErrorMessage & " - "

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

            ' (Salto incodinzionato riport lo status dell'operazione completata ricorsivamente)
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

                    ' (Salto incodinzionato riport lo status dell'operazione completata ricorsivamente)
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
            m_LastErrorMessage = "Tentivo di riallineamento fallito"

            Return StatusCode.KO ' False

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
                If m_CommandToCall = enCommandToCall.Payment Then
                    CType(frmEmulation, FormBuonoChiaro).Paid = m_PaidAmount
                    CType(frmEmulation, FormBuonoChiaro).Payable = m_PayableAmount
                Else
                    CType(frmEmulation, FormBuonoChiaro).Paid = m_VoidAmount
                    CType(frmEmulation, FormBuonoChiaro).Payable = m_VoidableAmount
                End If
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

            ' Msg Utente ( Status e ErrorMessage definiti da un azione dentro l'hanlder di un evento gestito esternamene )
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

            Else

                ' Riportiamo la descrizione più estesa
                m_LastErrorMessage = ProxyError.ErrorDescription

            End If

        Else

            ' Altro non previsto in questa funzione             *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            Dim ProxyError As ExceptionProxyArgentea = New ExceptionProxyArgentea(funcname, ExceptionProxyArgentea.LOC_ERROR_NOT_CLASSIFIED, "Errore non classificato e inatteso -- Exception UKNOWED --", ex)

            m_LastStatus = "UKNOWED." & ProxyError.ErrorTarget & "." & ProxyError.retCode
            m_LastErrorMessage = ProxyError.ErrorDescription & "> " & ProxyError.retCode & "<"

        End If

        ' Se l'eccezione è a cascata di altre...
        If Not ex.InnerException Is Nothing Then
            LOG_Debug(funcname, "Errore con exception interna :: " & m_LastStatus & " -- " & m_LastErrorMessage + " -- " & " -- " & ex.Message & " --" & ex.InnerException.Message)
        Else
            LOG_Debug(funcname, "Errore gestito :: " & m_LastStatus & " -- " & m_LastErrorMessage & " -- " & ex.Message)
        End If

        ' Msg Utente  ( Ultimo Status e ErrorMessage impotato dal Tipo di Exception gestita )
        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
        '

    End Sub


#End Region

#Region "Functions Common and Argentea Specifique"

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
        '''     Il totale stornato dall'nsieme dei buoni transitati 
        '''     nella sessione del POS sullo storno corrente 
        ''' </summary>
        ''' <returns></returns>
        Protected Friend Overridable ReadOnly Property totalVoidedWithBP() As Decimal
            Get
                Return m_TotalVoided_CS
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
        ''' <param name="BarcodeToSearch">Un EAN usato come BP per partecipare al pagamento</param>
        ''' <returns></returns>
        Public Function ContainsBarcode(BarcodeToSearch As String) As Boolean
            For Each itm As PaidEntry In m_ListEntries
                If itm.Barcode = BarcodeToSearch.Trim() Then
                    Return True
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
'''     Exceptio dedicata al Proxy Argentea
''' </summary>
Friend Class ExceptionProxyArgentea : Inherits System.Exception

    ' *^*^*^*^*^*^*^*^*^*^*

    ' ECCEZIONI DI TIPO GLOBAL GENERAL PURPOSE

    ' Su Errore generale di comunciazione che ha
    ' restituito una chiamata ad un metodo della dll di Argentea.
    Public Const GLB_SOCKET_ERROR As String = "SOCKET_ERROR"

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
    ''' <param name="ret_Code">Il returno code che ha restituito la dll all'uscita</param>
    ''' <param name="Ref_MessageOut">Il messaggio dalla dll di argentea che è stato restituito</param>
    Public Sub New(func_Name As String, Method_Name As String, Api_Called As ClsProxyArgentea.enApiToCall, ret_Code As Integer, Ref_MessageOut As String, Optional ByVal innerException As System.Exception = Nothing)
        MyBase.New(func_Name & "." & Method_Name & "." & Api_Called.ToString() & "." & CStr(ret_Code), innerException)

        funcName = func_Name
        methodName = Method_Name
        ApiCalled = Api_Called
        retCode = ret_Code
        RefTo_MessageOut = Ref_MessageOut
        '
        m_ErrorTarget = String.Empty
        m_ErrorDescription = String.Empty

        ' ** KO --> Exception su Errori di comunicazione o per risposta remota data da Argentea KO.
        LOG_Error(func_Name, "Exception on .:  " & func_Name & " for Api to Call .: " & Api_Called.ToString() & " to Method Argentea .: " & Method_Name & " in response receive retCode .: " & CStr(ret_Code) & " with raw Message out .: " & Ref_MessageOut)

        ' Su risposta da COM  in  negativo
        ' in ogni formatto il returnString
        ' ma con la variante che già mi filla
        ' l'attributro ErrorMessage
        ' .::: ApiCalled, retCode, RefTo_MessageOut sono già valorizzati
        m_LastResponseRawArgentea = _ParseErrorAndMapToThisException()

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
            retCode = 9010  ' Senza Eccezione interna   (già classificato)
        Else
            retCode = 9011  ' Con Eccezione interna     (non classificato)
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

    ''' <summary>
    '''     Esegue il parsing del protocollo su una risposta di Argentea
    '''     per formulare il success o l'unseccessfull con il  messaggio 
    '''     ripreso dalla codifica del protocollo come risposta.
    ''' </summary>
    ''' <returns>Restituisce un tipo <see cref="ArgenteaFunctionReturnObject"/> mappato con gli attributi della risposta decodificata dal protocollo!</returns>
    Private Function _ParseErrorAndMapToThisException() As ArgenteaFunctionReturnObject

        ' MAP e NAT della codifica da protocollo della Risposta Argentea
        Dim Response As Tuple(Of Boolean, Boolean, String, String, ArgenteaFunctionReturnObject) = ClsProxyArgentea.ParseProtocolForMapResponse(
                ApiCalled, retCode, RefTo_MessageOut, funcName, methodName
        )

        m_ErrorComunication = Response.Item1                    ' Errore di Comunicazione SI/NO
        m_ErrorOnParseProtocol = Response.Item2                 ' Errore di Parsing SI/NO
        m_ErrorTarget = Response.Item3                          ' Errore Target esrpresso come costante
        m_ErrorDescription = Response.Item4                     ' Descrizione dell'errore in modo esteso
        Return Response.Item5                                   ' Error Object

    End Function

End Class
