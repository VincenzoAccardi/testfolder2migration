Imports System
Imports ARGLIB = PAGAMENTOLib
Imports System.Collections
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos.EFT
Imports TPDotnet.Pos
Imports System.Windows.Forms

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

    ' COSTANTI PARAMETRI UTILIZZATE in Operator su Parameter
    Private Const OPT_BPParameterRuppArgentea As String = "BP_ParameterRuppArgentea"        ' <-- Parametro Stringa RUPP per protocollo Argentea

    ' COSTANTI PARAMETRI UTILIZZATE in Operator su IT.Parameter
    Private Const OPT_BPAcceptExcedeedValues As String = "BP_AcceptExcedeedValues"          ' <-- Accetta o meno il Resto sui BP Y o N
    Private Const OPT_BPNumMaxPayablesOnVoid As String = "BP_NumMaxPayablesOnVoid"          ' <-- Numero massimo di Buoni Pasto utilizzati per la vendita in corso 0 o ^n

    ''' <summary>
    '''     Collection usata per le Transazioni
    '''     Argentea andati a buon fine dopo la
    '''     chiamata alla funzione Argentea o da 
    '''     POS hardware
    ''' </summary>
    'Private _listBpCompletated As New Collections.Generic.Dictionary(Of String, BPType)(System.StringComparer.InvariantCultureIgnoreCase)
    Private _DataResponse As DataResponse

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

    Private m_CurrentSplitMode As String = "-"


#End Region

#Region "Property specifiche riprese dalla Transazione o dai parametri di Configurazione del Controller globale"
    Private m_RUPP As String = Nothing


    ''' <summary>
    '''     Restituisce per la codifica il riferimento al Codice RUPP ripreso dalla Configurazione Globale presa
    '''     nel Backsotre di Cassa dal Client che identitica il POS Hardware e Service ID di Account verso Argentea.
    ''' </summary>
    ''' <returns>String</returns>
    Private ReadOnly Property GetPar_RUPP() As String
        Get
            ' tech. primo accesso ( * per i parametri uso questo trick )
            If m_RUPP <> Nothing And m_RUPP <> "" Then m_RUPP = m_TheModcntr.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPParameterRuppArgentea)
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
        '''     o il pos locale per un pagamento
        ''' </summary>
        Payment

        ''' <summary>
        '''     Avvia il servizio remoto
        '''     o il pos locale per uno storno
        ''' </summary>
        Void


    End Enum

    ''' <summary>
    '''     Usato nelle chiamate Hundler per
    '''     formattare secondo le specifiche
    '''     del protocollo Argentea il CSV in
    '''     risposta sul MsgOut
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
    '''     Tipo di elemento da restituire 
    '''     nel set di risultati del servizio.
    ''' </summary>
    Public Structure BPType

        Public szKey As String
        Public szVal As String

        Public Sub New(sKey As String, sVal As String)
            szKey = sKey
            szVal = sVal
        End Sub

    End Structure


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

    '
    ' Per il servizio che usa un Form interno
    ' per l'inserimento manuale dei Buoni Pasto
    ' mi appoggio su un Form della cassa già presente.
    '
    Protected frmScanCodes As FormBuonoChiaro = Nothing     ' <-- Il Form di appoggio per servire il POS software sulla cassa corrente

    '
    ' Passati dal Chiamante per essere
    ' letti agiornati in uscita.
    '
    Private m_Paid As Decimal                               ' <-- Il Pagato fino ad adesso all'entrata
    Private m_PayableAmount As Decimal                      ' <-- Il pagabile con le azioni del servizio
    Private m_Void As Decimal                               ' <-- Lo Storno attuale fino ad adesso all'entrata
    Private m_VoidableAmount As Decimal                     ' <-- Lo Stornabile o lo stornato con le azioni del servizio

    '
    ' Aggiornati per il Risultato
    '
    Shared m_TotalBPUsed_CS As Integer                      ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento
    Shared m_TotalPayed_CS As Decimal                       ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente
    Shared m_TotalValueExcedeed_CS As Decimal               ' <-- Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata

    '
    ' Variabili private
    '

    Private m_LastStatus As String                          ' <-- Ultimo Status di Costante per errore in STDOUT
    Private m_LastErrorMessage As String                    ' <-- Ultimo Messaggio di errore STDOUT


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
    Protected m_taobj As TA
    Protected m_TheModcntr As ModCntr

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
    Public Event Event_ProxyCollectDataReturnedAtEnd(ByRef sender As Object, ByRef resultData As DataResponse)

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
    Protected Friend Sub New(ByRef theModCntr As ModCntr, ByRef taobj As TA, TypeBehavior As enTypeProxy,
                             ByVal CurrentTransactionID As String, ByVal CurrentPaymentsTotal As Decimal
                             )

        m_TypeProxy = TypeBehavior
        m_CurrentTransactionID = CurrentTransactionID
        m_CurrentPaymentsTotal = CurrentPaymentsTotal

        ' BEHAVIOR
        If TypeBehavior = enTypeProxy.Pos Then

            '
            ' 
            '


        Else ' default service

            ' Istanza del form di appggio ad uso 
            ' operatore per l'inserimento per ogni
            ' BP che deve partecipare al pagamento.
            frmScanCodes = theModCntr.GetCustomizedForm(GetType(FormBuonoChiaro), STRETCH_TO_SMALL_WINDOW)

            '
            ' Riporto come property al form  da 
            ' visualizzare per una sua gestione
            ' interna il Controller e la Transazione
            '
            frmScanCodes.theModCntr = theModCntr
            frmScanCodes.taobj = taobj

            '
            ' Preparo ad accettare l'handler degli eventi gestiti
            ' durante la scansione di ogni singolo buono per 
            ' l'accettazione o lo storno oltre che l'inizializzazione.
            '
            AddHandler frmScanCodes.BarcodeRead, AddressOf BarcodeReadHandler
            AddHandler frmScanCodes.BarcodeRemove, AddressOf BarcodeRemoveHandler

            '
            ' Evento chiave all'ok del form o alla chiusura del pos
            ' per il collect dei dati in risposta al chiamante.
            '
            AddHandler frmScanCodes.FormClosed, AddressOf CloseOperationHandler

        End If

        '
        ' Gli oggetti di base
        '
        m_taobj = taobj
        m_TheModcntr = theModCntr

    End Sub

#End Region

#Region "Properties Pubbliche"

    ''' <summary>
    '''     All'entrata definisce il Pagato al momento
    '''     All'usscita è aggiornato con il pagato dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property Paid() As Decimal
        Get
            Return m_Paid
        End Get
        Set(ByVal value As Decimal)
            m_Paid = value
            _updatePosForm()
        End Set
    End Property

    ''' <summary>
    '''     All'entrata definisce il Pagabile massimo
    '''     All'uscita è aggiornato con il pagato dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property Payable() As Decimal
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
    Public Property Void() As Decimal
        Get
            Return m_Void
        End Get
        Set(ByVal value As Decimal)
            m_Void = value
            _updatePosForm()
        End Set
    End Property

    ''' <summary>
    '''     All'entrata definisce lo Stornabile massimo
    '''     All'uscita è aggiornato con lo stornato dopo lo STDIN
    ''' </summary>
    ''' <returns>Valore espresso in decimal</returns>
    Public Property Voidable() As Decimal
        Get
            Return m_VoidableAmount
        End Get
        Set(ByVal value As Decimal)
            m_VoidableAmount = value
            _updatePosForm()
        End Set
    End Property


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
            If Not frmScanCodes Is Nothing Then
                frmScanCodes.Paid = m_Paid
                frmScanCodes.Payable = m_PayableAmount
            End If
        End If

    End Sub

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

    'Friend Property ListBpCompletated As Dictionary(Of String, BPType)
    'Get
    'Return _listBpCompletated
    'End Get
    'Set(value As Dictionary(Of String, BPType))
    '       _listBpCompletated = value
    'End Set
    'End Property

#End Region

#Region "Il progressive call Privato (Relativo a tutte le chiamate in sequenza richieste dal protocollo)"

    Protected m_ProgressiveCall As Integer = 1

    ''' <summary>
    '''     Argomento specializzato per le chiamate a 
    '''     Argentea come Progressivo di call tra servizi.
    ''' </summary>
    ''' <returns>Integer</returns>
    Private ReadOnly Property ProgressiveCall() As Integer
        Get
            Return m_ProgressiveCall
        End Get
    End Property

    ''' <summary>
    '''     Increment Progressive internal for Argentea
    '''     to call vs Remote service BPC
    ''' </summary>
    ''' <returns></returns>
    Private Function IncrementProgressiveCall() As Integer
        m_ProgressiveCall += 1
        Return IncrementProgressiveCall
    End Function

#End Region

#Region "Actions functions Public"

    ''' <summary>
    '''     Esegue il Connect
    '''     al POS per farlo
    '''     attendere sulle operazioni
    '''     o al SERVIZIO remoto
    '''     visualizzando il form per le scansioni
    ''' </summary>
    Friend Sub Connect()

        Dim funcName As String = "ProxyArgentea.Connect"

        ' Reset conteggio
        m_TotalPayed_CS = 0
        m_TotalValueExcedeed_CS = 0
        m_TotalBPUsed_CS = 0

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

                ' Avvio il form con la gestione
                ' del POS software tramite service.
                StartPosSoftware()

            Else

                ' Avvio il pos locale con la gestione
                ' del POS hardware tramite terminale.
                StartPosHardware()

            End If

            ' Flag locale che stato attivo
            m_bWaitActive = False

        Else
            ' Sollevo l'eccezione
            Throw New Exception(GLB_PROXY_ALREDAY_RUNNING)
        End If

    End Sub

    ''' <summary>
    '''     Utilità per parcheggiare in caso di
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
            frmScanCodes.Hide()
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
                frmScanCodes.Show()
                System.Windows.Forms.Application.DoEvents()

            Else

                ' Mostra la finestra di Wati avviata al connect
                FormHelper.ShowWaitScreen(m_TheModcntr, False, Nothing)

            End If

        End If

    End Sub

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
    '''     Azzera completamente lo stato del
    '''     Proxy per istanziare una nuova chiamata.
    ''' </summary>
    Friend Sub Close()

        '
        ' Chiudo eventuali finestre per il wait
        ' operatore. 
        '
        FormHelper.ShowWaitScreen(m_TheModcntr, True, Nothing)

        ' Deferenziazione (free mem)
        ArgenteaCOMObject = Nothing
        frmScanCodes = Nothing
        _DataResponse = Nothing

    End Sub

    ''' <summary>
    '''     Funzione di Utility su chiamate
    '''     esterne combiante.
    ''' </summary>
    Friend Sub Reset()

        ' Reset dello status dei contatori
        _DataResponse = Nothing
        m_TotalPayed_CS = 0
        m_TotalValueExcedeed_CS = 0
        m_TotalBPUsed_CS = 0
        m_CurrentTransactionID = String.Empty
        m_CurrentBarcodeScan = String.Empty
        m_bWaitActive = False
        m_ServiceStatus = enProxyStatus.Uninitializated
        m_PayableAmount = 0
        m_Paid = 0
        m_VoidableAmount = 0
        m_Void = 0

    End Sub

#End Region

#Region "Parser Functions Privates"

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

        ' Parser Type (Valorizza anche m_CurrentSplitMode)
        Dim ParsingMode As InternalArgenteaFunctionTypes = GetSplitAndFormatModeForParsing()
        m_CurrentSplitMode = ";"

        ' Parsiamo la risposta argentea per l'azione BP
        If (Not CSVHelper.ParseReturnString(CSV, ParsingMode, objTPTAHelperArgentea, m_CurrentSplitMode)) Then

            LOG_Debug(getLocationString(funcName), "BPC Parsing Protcol Argentea Fail to Parse 'Message Response' for this " & funcName & " response in MessageOut")

            ' Su Errore di Parsing solleviamo immediatamente l'eccezione per uscire dalla
            ' gestione della comunicazione Argentea.
            Throw New Exception(GLB_ERROR_PARSING)

        Else

            ' RIPORTO SUL FLOW quelli concerni allo Stato di OK Success o KO Error

            ' Risposta Codicficata da Risposta Raw Argentea
            ResponseArgentea = objTPTAHelperArgentea(0)

            ' Log in risposta e decodifica di rpotocoolo effettutato di argentea.
            LOG_Debug(getLocationString(funcName), "Parsed Protcol Argentea 'Message Response' for this " & funcName & " Status: " + ResponseArgentea.Successfull.ToString() + " Response: " + ResponseArgentea.SuccessMessage + ResponseArgentea.ErrorMessage)


            ' E per questa specifica Azione fortunatamente
            ' abbiamo il codice di Stato
            'pParams.Status = ResponseArgentea.CodeResult

            ' Riprendo queste  notazioni  rispettivamente
            ' per il TerminalID che ha eseguito  la trans
            ' il Valore del Buono Pasto dato da  Argentea
            ' e soprattutto se richiede un ulteriore call
            ' verso argentea di conferma alla trans.
            'pParams.TerminalID = ResponseArgentea.TerminalID
            'pParams.Value = CDec(ResponseArgentea.Amount) / 100
            'pParams.CommittRequired = CBool(ResponseArgentea.RequireCommit)

            Return ResponseArgentea

        End If

    End Function

    Private Function GetSplitAndFormatModeForParsing() As InternalArgenteaFunctionTypes

        ' Behavior
        Dim ParsingMode As InternalArgenteaFunctionTypes
        Dim m_CurrentSplitMode As String = "-"
        If m_TypeProxy = enTypeProxy.Pos Then

            If m_CommandToCall = enCommandToCall.Payment Then

                ParsingMode = InternalArgenteaFunctionTypes.BPEPayment
                m_CurrentSplitMode = ";"

            ElseIf m_CommandToCall = enCommandToCall.Void Then

                ParsingMode = InternalArgenteaFunctionTypes.BPEVoid
                m_CurrentSplitMode = ";"

            End If

        Else ' service

            If m_CommandToCall = enCommandToCall.Payment Then

                ParsingMode = InternalArgenteaFunctionTypes.BPCPayment
                m_CurrentSplitMode = ";"

            ElseIf m_CommandToCall = enCommandToCall.Void Then

                ParsingMode = InternalArgenteaFunctionTypes.BPCVoid
                m_CurrentSplitMode = ";"

            End If

        End If

        Return ParsingMode

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

        ' Parser Type (Valorizza anche m_CurrentSplitMode)
        Dim ParsingMode As InternalArgenteaFunctionTypes = GetSplitAndFormatModeForParsing()
        m_CurrentSplitMode = ";"

        If (Not CSVHelper.ParseReturnString(CSV, ParsingMode, objTPTAHelperArgentea, m_CurrentSplitMode)) Then

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


#Region "** TERMINALE LOCALE POS HARDWARE -> Con i sui metodi nella dll COM in gestione al dispositivo"

    ''' <summary>
    '''     Avvia e mette in attesa il termianale
    '''     hardware collegato alla cassa corente.
    ''' </summary>
    Private Sub StartPosHardware()
        Dim funcName As String = "StartPosHardware"

        ' Entrap sull'idle
        Try

            ' In questa modalità avvio solo un form
            ' modale a pieno schermo per  attendere 
            ' le operazioni dal Pos Locale collegato.
            FormHelper.ShowWaitScreen(m_TheModcntr, False, Nothing, "Attesa su Terminale Locale", "BP Wait")

            ' Dispongo le proprietà del Form fittizio di Cassa
            ' ripreso nel Controller globale per la
            ' preparazione a non prendere lo  status
            ' attivo durante il collegamento al terminale dove 
            ' si sta operando con il controllo utente sulla
            ' dematerializzazione della carta elettronica BP.
            m_TheModcntr.DialogActiv = True
            m_TheModcntr.DialogFormName = "POS HARDWARE"
            m_TheModcntr.SetFuncKeys((False))

            ' Status
            m_ServiceStatus = enProxyStatus.InRunning

            ' Idle
            Dim _CallHardware As Boolean = CallHardware("StartPosHardware")

            'Do While frmScanCodes.bDialogActive = True
            System.Threading.Thread.Sleep(100)
            System.Windows.Forms.Application.DoEvents()
            'Loop
            ' Emulo l'event Handler come in modalità service
            CloseOperationHandler(Nothing, Nothing)

            ' Dichiaro come concluso correttamente tutto
            If m_ServiceStatus = enProxyStatus.InRunning Then

                ' Se era rimasto in Running e non InError
                ' tutto è filato liscio e torno con OK
                m_ServiceStatus = enProxyStatus.OK

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
                '  Rimetto sul controler principale 
                ' lo stato attivo dopo che il terminale
                ' hardware ha completato la sua attesa.
                '
                m_TheModcntr.DialogActiv = False
                m_TheModcntr.DialogFormName = ""
                m_TheModcntr.SetFuncKeys((True))
                m_TheModcntr.EndForm()
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

                ' Log e segnale non aggiornato in uscita e chiusura 
                LOG_Debug(getLocationString(funcName), m_LastErrorMessage + "--" + ex.InnerException.ToString())
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            Finally

                ' Chiudo per memory leak con argentea
                ArgenteaCOMObject = Nothing

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
    Private Sub StartPosSoftware()
        Dim funcName As String = "StartPosSoftware"

        ' Entrap sull'idle
        Try

            '
            ' Mostro il form per la gestione e comunicazione
            ' con il Servizio remoto di convalida su azioni di
            ' Dematerializzazione e Storno.
            '
            frmScanCodes.Show() ' non modal VB Dialog

            ' Dispongo le proprietà del Form  Cassa
            ' ripreso nel Controller globale per la
            ' preparazione a non prendere lo  status
            ' attivo durante la scansione dove si sta 
            ' operando con il controllo locale del form 
            ' che ha la textbox per prendere i codici EAN
            m_TheModcntr.DialogActiv = True
            m_TheModcntr.DialogFormName = frmScanCodes.Text
            m_TheModcntr.SetFuncKeys((False))

            '
            ' Idle sul Form Locale
            ' Finestra di dialogo avviata e rimango in idle 
            ' finchè l'operatore non finisce le azioni necessarie.
            '
            frmScanCodes.bDialogActive = True

            ' Status
            m_ServiceStatus = enProxyStatus.InRunning

            ' Idle
            Do While frmScanCodes.bDialogActive = True
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()
            Loop

            ' Dichiaro come concluso correttamente tutto
            If m_ServiceStatus = enProxyStatus.InRunning Then

                ' Se era rimasto in Running e non InError
                ' tutto è filato liscio e torno con OK
                m_ServiceStatus = enProxyStatus.OK

            End If

        Catch ex As Exception

            ' Scrive una riga di Log per l'errore in corso
            ' e lo gestisce in seguito sotto nel finally...
            LOG_ErrorInTry(getLocationString(funcName), ex)

        Finally

            Try

                '
                '  Provo a chiudere il Form del POS
                ' software se siamo in questa modalità.
                '
                If Not frmScanCodes Is Nothing Then
                    m_TheModcntr.DialogActiv = False
                    m_TheModcntr.DialogFormName = ""
                    m_TheModcntr.SetFuncKeys((True))
                    m_TheModcntr.EndForm()
                    frmScanCodes.Close()
                    frmScanCodes = Nothing
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

                ' Log e segnale non aggiornato in uscita e chiusura 
                LOG_Debug(getLocationString(funcName), m_LastErrorMessage + "--" + ex.InnerException.ToString())
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            Finally

                ' Chiudo per memory leak con argentea
                ArgenteaCOMObject = Nothing

                ' Effettuo un Dispose forzato per 
                ' la chiusura del form su eccezioni.
                If Not frmScanCodes Is Nothing Then
                    frmScanCodes.Dispose()
                    frmScanCodes = Nothing
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
                ' Opzione Max BP pagabili per vendita.: 
                '       Se il Numero di Buoni pagabili per una vendita
                '       è superiore al numero di buoni passato procediamo
                '       con la sgnazlazione.
                Dim OptPayablesBPStr As String = m_TheModcntr.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPNumMaxPayablesOnVoid).Trim()
                Dim OptPayablesBP As Integer = CInt(Microsoft.VisualBasic.IIf(OptPayablesBPStr = "" Or OptPayablesBPStr = Nothing, "0", OptPayablesBPStr))

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
                    ' Opzione Operatività.: 
                    '       Se il Totale in ingresso è minore rispetto 
                    '       al valore di facciata del Buono Pasto una volta
                    '       ottenuto dalla materializzazione, opto per troncare su totale.
                    Dim OptAcceptExceeded As Boolean = Microsoft.VisualBasic.IIf(m_TheModcntr.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPAcceptExcedeedValues).Trim().ToUpper() = "Y", True, False)

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

                                ' Segnalo Operatore di Cassa
                                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
                                LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Excedeed")

                                ' Immediatamente annullo verso il sistema argnetea l'operazione
                                ' Per rimuoverlo tramite il metodo stesso per l'annullo
                                m_FlagUndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
                                Me.BarcodeRemoveHandler(sender, m_CurrentBarcodeScan)
                                m_FlagUndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

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

                            ' Errata Conferma Dematerializzazione
                            LOG_Debug(getLocationString(funcName), "Transaction Dematerialize on Argentea ::KO:: ON CONFIRM")

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
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
                        LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Remote")
                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
                    LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

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
    '''     per le azioni di annullo di materializzazione del buono pasto cartaceo.
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
                    m_LastErrorMessage = "Il BarCode non è presente tra le scelte possibili!!"

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

                            ' Errata Conferma Annullament Buono Dematerializzato
                            LOG_Debug(getLocationString(funcName), "Transaction Reverse Dematerialize Argentea ::KO:: ON CONFIRM")

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
                        WriterResultDataList.Remove(New PaidEntry(m_CurrentBarcodeScan, m_CurrentTerminalID))
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
                        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
                        LOG_Debug(getLocationString(funcName), "Transaction Reverse Dematerialize Argentea ::KO::")
                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(m_TheModcntr, True, sender)
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
                    LOG_Debug(getLocationString(funcName), "Transaction Reverse Dematerialize Argentea ::KO:: NOT INIZIALIZATED")

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
    '''     Handdler dell'evento chiave sia quando è in modalità service che pos
    '''     per convogliare i dati in ingresso e restituirli al chiamante,.
    ''' </summary>
    ''' <param name="sender">Il form del Pos software o il COM del service pos locale hardware</param>
    ''' <param name="e"></param>
    Private Sub CloseOperationHandler(sender As Object, e As FormClosedEventArgs)

        Try

            '
            '   Evento chiave di chiusura
            '
            RaiseEvent Event_ProxyCollectDataTotalsAtEnd(Me, _DataResponse)

        Catch ex As Exception

            ' Intercettiamo l'errore per il contesto probabilmente
            ' erchè il consumer non l'ha fatto per suo conto, quindi
            ' rimane che per noi il consumer con i dati non è aggioranto.
            Throw New Exception(GLB_ERROR_ON_EVENT_DATA, ex)

        End Try

    End Sub

#End Region

#Region "Functions Private per Hardware Mode"

    ''' <summary>
    '''     Inizializza la Sessione verso Argentea
    '''     e parte la numerazione interna delle chiamate
    '''     da 1
    ''' </summary>
    Private Function CallHardware(funcName As String) As Boolean

        ' OUT su chiamate
        Dim RefTo_MessageOut As String = Nothing

        CallHardware = False

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
        ' Idle sulla chiamata diretta al POS 
        ' chiamato dalla funzione API di 
        ' argentea per avviare una  sessione
        ' sul POS locale di pagamento.
        '
        '   amount =                 ''' L'importo per avviare il POS a farsi pagare in BP l'importo dettato
        '

#If DEBUG_SERVICE = 0 Then
        If m_CommandToCall = enCommandToCall.Payment Then

            retCode = ArgenteaCOMObject.PaymentBPE(
                    CInt(m_PayableAmount * 100),
                     RefTo_Transaction_Identifier,
                     RefTo_MessageOut
                 )

        ElseIf m_CommandToCall = enCommandToCall.Void Then

            retCode = ArgenteaCOMObject.VoidBPE(
                    CInt(m_VoidableAmount * 100),
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
            RefTo_MessageOut = "OK;TRANSAZIONE ACCETTATA;2|5|1020|1|414;104;PELLEGRINI;  PAGAMENTO BUONO PASTO " ' <-- x test 
        End If
        retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:
#End If


        ' Riprendiamo la Risposta cos' come è stata
        ' data per il log di debug grezza
        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". BP: Hardware Output: " & RefTo_MessageOut)

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            Dim ParseErr As ArgenteaFunctionReturnObject = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Activation check for BPE with returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)
            Return False

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Dim RespSrv As ArgenteaFunctionReturnObject = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If RespSrv.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                'IncrementProgressiveCall()

                '
                ' A differenza del Software  Creo  voci
                ' di TA quanti sono stati inoltrati nel
                ' dispositivo.
                '
                m_TotalBPUsed_CS = RespSrv.NumBPEvalutated          ' <-- Il Numero dei buoni utilizzati in questa sessione di pagamento
                m_TotalPayed_CS = RespSrv.Amount                    ' <-- L'Accumulutaroe Globale al Proxy corrente nella sessione corrente
                m_TotalValueExcedeed_CS = 0                         ' <-- ?? TODO:: Il Totale in eccesso se l'opzione per accettare valori maggiori è abilitata

                ' Riprendo l'elenco riportato dall'hardware
                ' per ogni taglio e colloco ricopiandolo il 
                ' pezzo interessato
                For Each itm As Object In RespSrv.ListBPsEvaluated

                    ' Questo dall'hardware non c'è l'abbiamo
                    ' e portiamo un code contatore
                    Dim paidValue As Decimal = itm.Value
                    Dim faceValue As Decimal = itm.Value

                    m_CurrentBarcodeScan = itm.Key
                    m_CurrentTerminalID = RespSrv.TerminalID

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
                LOG_Debug(getLocationString(funcName), "BP comunication with terminal pos successfuly on call first with message " & RespSrv.SuccessMessage)
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

#Region "Functions Private per Service mode"

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
        RefTo_MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 
        retCode = ArgenteaFunctionsReturnCode.OK
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
            Dim ParseErr As ArgenteaFunctionReturnObject = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Activation check for BPC with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)
            Return False

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Dim RespSrv As ArgenteaFunctionReturnObject = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If RespSrv.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                IncrementProgressiveCall()

                ' ** INIZIALIZZATA e corretamente chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP inizialization " & m_CurrentBarcodeScan & " successfuly on call first with message " & RespSrv.SuccessMessage)
                Return True

            Else

                ' Non inizializzata da parte di Argentea per
                ' errore remoto in risposta a questo codice.
                LOG_Debug(getLocationString(funcName), "BPC inizialization " & m_CurrentBarcodeScan & " remote failed on first call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)
                Return False

            End If

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
            Dim ParseErr As ArgenteaFunctionReturnObject = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Dematerialization for BP with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)

            ' Esco dal  flow immediatamente
            m_CurrentValueOfBP = ParseErr.GetAmountValue(1)
            m_CurrentTerminalID = ParseErr.TerminalID
            Return CallDematerialize

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Dim RespSrv As ArgenteaFunctionReturnObject = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If RespSrv.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                IncrementProgressiveCall()

                ' Se la risposta argenta richiede un ulteriore 
                ' conferma allora procedo ad uscire per il flow.
                If RespSrv.RequireCommit Then

                    ' ** DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP dematirializated with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & RespSrv.SuccessMessage)

                    ' RICHIESTO CONFERMA
                    m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                    m_CurrentTerminalID = RespSrv.TerminalID
                    CallDematerialize = StatusCode.CONFIRMREQUEST
                    Return CallDematerialize
                Else

                    ' ** DEMATERIALIZZATO corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP dematirializated " & m_CurrentBarcodeScan & " successfuly on call with message " & RespSrv.SuccessMessage)

                    ' COMPLETATO
                    m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                    m_CurrentTerminalID = RespSrv.TerminalID
                    CallDematerialize = StatusCode.OK
                    Return CallDematerialize
                End If

            Else

                ' Non dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BP dematerializated " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' NON EFFETTUATO
                m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                m_CurrentTerminalID = RespSrv.TerminalID
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
            Dim ParseErr As ArgenteaFunctionReturnObject = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Reverse Dematerialization for BPC with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)

            ' Esco dal  flow immediatamente
            m_CurrentValueOfBP = ParseErr.GetAmountValue(1)
            CallReverseMaterializated = StatusCode.KO
            Return CallReverseMaterializated
        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Dim RespSrv As ArgenteaFunctionReturnObject = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If RespSrv.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                IncrementProgressiveCall()

                ' Se la risposta argenta richiede un ulteriore 
                ' conferma allora procedo ad uscire per il flow.
                If RespSrv.RequireCommit Then

                    ' ** REVERSE SU DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP reverse dematirializated with wait confirm " & m_CurrentBarcodeScan & " successfuly on call with message " & RespSrv.SuccessMessage)

                    ' RICHIESTO CONFERMA
                    m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                    m_CurrentTerminalID = RespSrv.TerminalID
                    CallReverseMaterializated = StatusCode.CONFIRMREQUEST
                    Return CallReverseMaterializated
                Else

                    ' ** REVERSE SU DEMATERIALIZZATO corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BP reverse dematirializated " & m_CurrentBarcodeScan & " successfuly on call with message " & RespSrv.SuccessMessage)

                    ' COMPLETATO
                    m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                    m_CurrentTerminalID = RespSrv.TerminalID
                    CallReverseMaterializated = StatusCode.OK
                    Return CallReverseMaterializated
                End If

            Else

                ' Non reverse su dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BP reverse dematerializated " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' NON EFFETTUATO
                m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                m_CurrentTerminalID = RespSrv.TerminalID
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
            Dim ParseErr As ArgenteaFunctionReturnObject = Me.ParseErrorAndMapToParams(funcName, retCode, RefTo_MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Confirm " & sConfirmOperation & " for BP with  " & m_CurrentBarcodeScan & " returns error: " & m_LastErrorMessage & ". The message raw output is: " & RefTo_MessageOut)

            ' Esco dal  flow immediatamente
            m_CurrentValueOfBP = ParseErr.GetAmountValue(1)
            CallConfirmOperation = StatusCode.KO
            Return CallConfirmOperation
        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Dim RespSrv As ArgenteaFunctionReturnObject = Me.ParseResponseProtocolArgentea(funcName, RefTo_MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If RespSrv.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                IncrementProgressiveCall()

                ' ** CONFIRM su REVERSE o DEMATERIALIZZATO effettuata corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BP confirm " & sConfirmOperation & " for " & m_CurrentBarcodeScan & " successfuly on call with message " & RespSrv.SuccessMessage)

                ' COMPLETATO
                m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                m_CurrentTerminalID = RespSrv.TerminalID
                CallConfirmOperation = StatusCode.OK
                Return CallConfirmOperation
            Else

                ' Non confirm su reverse o dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BP confirm " & sConfirmOperation & " for " & m_CurrentBarcodeScan & " remote failed on call to argentea with message code " & m_LastStatus & " relative to " & m_LastErrorMessage)

                ' NON EFFETTUATO
                m_CurrentValueOfBP = RespSrv.GetAmountValue(1)
                m_CurrentTerminalID = RespSrv.TerminalID
                CallConfirmOperation = StatusCode.KO
                Return CallConfirmOperation
            End If

        End If

    End Function

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
        LOG_Debug(getLocationString(funcname), m_LastErrorMessage + "--" + ex.InnerException.ToString())
        msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
        '

    End Sub


    Private Function ValidationVoucherRequest(barcode As String) As Boolean
        'Logic Comunication Barcode at Argentea Supplier

        ValidationVoucherRequest = False


    End Function


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