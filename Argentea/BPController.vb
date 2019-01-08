Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports TPDotnet.IT.Common.Pos.EFT
Imports TPDotnet.IT.Common.Pos

Public Class BPControllerBase
    Implements IBPDematerialize

    ' VARIANTE PER ELETTRONICI O CARTACEI

    ' Identifica nei metadata la Key per tipo
    Private TYPE_SPECIFIQUE As String = "N/D"                                       ' <-- Costante per identificare il sottotipo nell'applicazione (Cartacei/Elettronici)

    ' Modalità del servizio Proxy di Argentea
    Private MODE_PROXY_USE As ClsProxyArgentea.enTypeProxy = Nothing                ' <-- Il Proxy servizio avviato in modalità Emulazione Software (Servizio Remoto WebService o Pos Hardware Locale)

#Region "CONST di INFO e ERRORE Private"

    ' Messaggeria per codifica segnalazioni ID di errore remoti 
    Private msgUtil As New TPDotnet.IT.Common.Pos.Common

    ' Su Errore di configurazione o di sistema quando la 
    ' procedura non trova il file xslt di trasformazione 
    ' usato per riprendere i totali dei prodotti che  si
    ' possono pagare tramite BP
    Private Const OPR_ERROR_FILE_FILTER As String = "Error-FILE_FILTER_PAYABLES_BP"

    ' Su errore di configurazione per i tipi di media che
    ' non risultavo validi per questa tipologia in corso.
    Private Const OPR_ERROR_MEDIA_TYPE_NOT_VALID As String = "Error-MEDIA_TYPE_NOT_VALID_BP"

    ' Per la corretta istanza di questo modello applicativo
    ' sono sempre necessari nei parametri caricati dinamicamente
    ' op er istanza di questa classe sia il controller ModCntr
    ' da cui viene avviato e la transazione in corso.
    Private Const OPR_ERROR_START_PROXY_INTERNAL As String = "ERRPR_START_PROXY_INTERNAL"

    ' ** Notifiche

    ' Su Errore Pagabile rispetto alla vendita corrente, non sono 
    ' stati trovati prodotti nella vendita corrente per questa tipologia.
    Private Const NOT_ERROR_CALL_DEMAT_INCONGRUENT As String = "Error-CALL_DEMAT_INCONGRUENT"

    ' Su Errore Pagabile rispetto alla vendita corrente, non sono 
    ' stati trovati prodotti nella vendita corrente per questa tipologia.
    Private Const NOT_ERROR_CALL_VOID_INCONGRUENT As String = "Error-CALL_VOID_INCONGRUENT"

    ' Su Errore Pagabile rispetto all'ammontare del pagamento
    ' assegno errore di non valido per uscire dal pagamento BP
    Private Const NOT_ERROR_EXCEDED_PAYABLE As String = "Error-PAYABLE_EXCEDEED"

    ' Su Errore Pagabile rispetto all'ammontare già totalmente pagato
    ' assegno errore di non valido per uscire dal pagamento BP
    Private Const NOT_ERROR_ALREADY_PAYED As String = "Error-TOTAL_ALREADY_PAYED"

    ' Throw Interne 

    ' Nel Flow della nel momento In cui si trova a creare un
    ' MEDIA di resto se c'è un errore di valutazione interna enon viene creato.
    Private Const INT_ERROR_NOT_CREATE_EXCEDEED As String = "Error-EXCEPTION-CREATE-EXCEDEED"

    ' ** Eccezioni

    ' Nel Flow della nel momento in cui si trova a creare un
    ' MEDIA di resto se c'è un eccezione interna enon viene creato.
    Private Const GLB_ERROR_ON_CREATE_EXCEDEED As String = "Error-EXCEPTION-ON-CREATE-EXCEDEED"

    ' Nell'istanziare la classe servizio verso Argentea (proxy mode)
    ' il controller di gestione corrente ha sollevato un eccezione.
    Private Const GLB_ERROR_INSTANCE_SERVICE As String = "Error-EXCEPTION-PROXY-INSTANCE"

    ' Nel riprendere i dati provenienti dal proxy di Argentea (proxy mode)                  ' ** Speciale nell'evento
    ' il controller di gestione corrente ha sollevato un eccezione.
    Private Const GLB_ERROR_COLLECT_DATA_SERVICE As String = "Error-EXCEPTION-COLLECT-SRVC-DATA"

    ' Nel Flow della funzione Entry il Throw non previsto (call assistenza).
    Private Const GLB_ERROR_NOT_UNEXPECTED As String = "Error-EXCEPTION-UNEXPECTED"

#End Region

#Region "Membri Privati"

    '
    ' Variabili private per il totalizzatore
    '
    Private m_PayableAmout As Decimal = 0               ' Nella Transazione corrente il totale in valore dei BP già usati nella vendita
    Private m_VoidAmount As Decimal = 0                 ' Nella chiamata di storno imposta il valore ripreso dalla TA Media selezionata in entrata
    Private _InitialBPPayed As Integer = 0              ' Nella Transazione corrente il conteggio dei BP già usati nella vendita

    '
    ' Variabili private
    '
    Private m_SilentMode As Boolean = False             ' <-- Se mostrare all'utente i messaggi di errore e di avviso
    Private m_LastStatus As String                      ' <-- Ultimo Status di Costante per errore in STDOUT
    Private m_LastErrorMessage As String                ' <-- Ultimo Messaggio di errore STDOUT

    '
    ' Parametri personalizzati per questo controller
    '
    Private pParams As BPParameters                     ' <-- Parametri Interni per Controller di tipo BP

    ' 
    ' Interni per gestione
    '
    Protected m_TheModcntr As ModCntr                   ' <-- Il Controller di riferimento dell'applicazione
    Protected m_taobj As TA                             ' <-- La TA in corso di sessione della transazione corrente
    Protected m_CurrMedia As TaMediaRec                 ' <-- La TA di pagamento per il quale ci stiamo muovendo

    Private m_ParseFractMode As Integer = 100           ' <-- Costante per Valuta Cents in €

#End Region

#Region ".ctor"

    ''' <summary>
    '''     .ctor
    ''' </summary>
    ''' <param name="theModCntr">controller -> Il Controller per riferimento dal chiamante</param>
    ''' <param name="taobj">transaction -> La TA per riferimento dal chiamante</param>
    'Public Sub New(ByRef theModCntr As ModCntr, ByRef taobj As TA)
    Protected Sub New(TypeSpecifique As String, ModeProxy As ClsProxyArgentea.enTypeProxy)

        LOG_Info("CONTROLLER", "CONTROLLER SPECIFIQUE " + TypeSpecifique + " TYPE " + ModeProxy.ToString())

        TYPE_SPECIFIQUE = TypeSpecifique
        MODE_PROXY_USE = ModeProxy

    End Sub

#End Region

#Region "Properties"

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


    ''' <summary>
    '''     Se visualizare o meno messagi di avviso
    '''     o di errore all'operatore tramite msgbox.
    ''' </summary>
    ''' <returns>True/False</returns>
    Public Property SilentMode() As Boolean Implements IBPDematerialize.SilentMode
        Get
            Return m_SilentMode
        End Get
        Set(Value As Boolean)
            m_SilentMode = Value
        End Set
    End Property

#End Region

#Region "IAction principe implementata"

    ''' <summary>
    '''     Gestiamo un wrap verso un Form da visualizzare
    '''     con l'handler degli eventi chiave per la gestione.
    ''' </summary>
    ''' <param name="Parameters">Dictionary di Parametri dinamici</param>
    ''' <returns>Stato a completamente o errore sull'azione principe corrente!! <see cref="IBPReturnCode"/></returns>
    Public Function Dematerialize(ByRef Parameters As Dictionary(Of String, Object)) As IBPReturnCode Implements IBPDematerialize.Dematerialize
        Dim funcName As String = "Dematerialize"

        '
        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPParameters.LoadCommonFunctionParameter
        '
        Dematerialize = _initializeParameters(funcName, Parameters, True)
        If Dematerialize = IBPReturnCode.KO Then
            Exit Function
        End If

        Try

            ' Riprendo dalla TA solo i prodotti relativi a quelli 
            ' che possono essere pagati con i Buoni Pasto
            Dim sTotalTransaction As String = Nothing
            If Not Common.ApplyFilterStyleSheet(m_TheModcntr, m_taobj, "BPCType.xslt", sTotalTransaction) Then

                ' Signal (File xslt per il raggurpamento di titoli validi per questa vendita non presente nel sistema)
                _SetOperationStatus(funcName,
                    OPR_ERROR_FILE_FILTER,
                    "Errore di configurazione del prodotto BP - File di trasformazione per la vendita con Buoni Pasto non valido o non presente  - (Chiamare assistenza).",
                    PosDef.TARMessageTypes.TPSTOP, True
                )

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            End If

            ' Se invece siamo con un totale a 0 da pagare perchè tutti i prodotti
            ' sono esclusi dalla vendita possibile tramite i Buoni Pasto.
            Dim TotalTransaction As Decimal = CDec(sTotalTransaction) / m_ParseFractMode

            If TotalTransaction = 0 Then

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                ' Signal (Prodotti non presenti in questa vendita pagabili con Buoni Pasto)
                _SetOperationStatus(funcName,
                    NOT_ERROR_CALL_DEMAT_INCONGRUENT,
                    "Non sono presenti prodotti in questa vendita che è possibile pagare con i Buoni Pasto",
                    PosDef.TARMessageTypes.TPERROR, True
                )

                Return IBPReturnCode.KO

            Else

                ' Questo è quello che si può pagare in totale.
                m_PayableAmout = TotalTransaction

                ' Questo è quello che ho già pagato
                Dim m_PayedAmount As Decimal = m_CurrMedia.dTaPaidTotal

                ' Quindi controlliamo se il totale intanto è maggiore di 0 quindi pagabile.
                If m_CurrMedia.dTaPaidTotal > 0 Then

                    ' Ricavo l'Amount rispetto al Totale
                    m_PayableAmout = Math.Min(m_taobj.GetTotal - (m_taobj.GetTotal - m_PayedAmount), m_PayableAmout)

                End If

            End If

            ' Quindi controlliamo se il totale intanto è maggiore di 0 quindi pagabile.
            If m_CurrMedia.dTaPaidTotal > 0 Then

                ' E se l'ammontare del pagmaneto non eccede il totale dovuto
                If m_PayableAmout <= 0 Then

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                    ' Signal (Pagamento già completo rispetto al totale da pagare)
                    _SetOperationStatus(funcName,
                        NOT_ERROR_EXCEDED_PAYABLE,
                        "Il pagamento eccede sul totale della vendita in corso!",
                        PosDef.TARMessageTypes.TPERROR, True
                    )

                    Dematerialize = IBPReturnCode.KO

                    Exit Function

                End If

                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **
                '
                ' Richiamiamo il Metodo per l'Azione per visualizzarlo
                ' passando l'intera transazione il Controller corrente
                ' Il Metodo di pagamento passato come argomento
                '
                If HandlePaymentBPCall(m_PayableAmout) Then
                    Dematerialize = IBPReturnCode.OK
                Else
                    Dematerialize = IBPReturnCode.KO
                End If
                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

            Else

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                ' Signal (Pagamento già completo rispetto al totale da pagare)
                _SetOperationStatus(funcName,
                    NOT_ERROR_ALREADY_PAYED,
                    "Il pagamento ricopre già l'intero totale della vendita in corso!",
                    PosDef.TARMessageTypes.TPERROR, True
                )

                Dematerialize = IBPReturnCode.KO

                Exit Function

            End If

        Catch ex As Exception

            Dematerialize = IBPReturnCode.KO

            ' Signal (PEccezione interna non gestita in questa funzione)
            _SetOperationStatus(funcName,
                GLB_ERROR_NOT_UNEXPECTED,
                "Errore interno non previsto 1 -- exception on call void in controller bp with argentea -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

        Finally
            LOG_FuncExit(getLocationString(funcName), Dematerialize.ToString())
        End Try

    End Function

    ''' <summary>
    '''     Azione da TA del controller Padre su Cassa
    '''     per eseguire lo storno su chiamate del tasto
    '''     Annullo sulla cassa per la voce specifica.
    ''' </summary>
    ''' <param name="Parameters">Dictionary di Parametri dinamici</param>
    ''' <returns>Stato a completamente o errore sull'azione principe corrente!! <see cref="IBPReturnCode"/></returns>
    Public Function Void(ByRef Parameters As Dictionary(Of String, Object)) As IBPReturnCode Implements IBPDematerialize.Void
        Dim funcName As String = "Void"

        '
        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPParameters.LoadCommonFunctionParameter
        '
        Void = _initializeParameters(funcName, Parameters, True)
        If Void = IBPReturnCode.KO Then
            Exit Function
        End If

        Try

            ' Riprendo l'importo da stornare per l'azione
            ' sulla riga corrente della TA (Con i controlli del caso)
            Dim sTotalVoid As String = Nothing
            '   (m_CurrMedia.dTaPaid >= 0) Or
            If m_CurrMedia Is Nothing _
                OrElse (m_CurrMedia.PAYMENTinMedia Is Nothing) _
                OrElse (m_CurrMedia.PAYMENTinMedia.szExternalID <> TYPE_SPECIFIQUE) Then

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Void Argentea ::KO:: Local")

                ' Signal (Storno non congruente per funzione mancante nella configurazione di sistema)
                _SetOperationStatus(funcName,
                    OPR_ERROR_MEDIA_TYPE_NOT_VALID,
                    "Errore di configurazione del prodotto BP - Funzione di storno non impostata - (Chiamare assistenza).",
                    PosDef.TARMessageTypes.TPERROR
                )

                Return IBPReturnCode.KO
            Else

                ' Importo per lo stornabile rimanente
                sTotalVoid = m_CurrMedia.dTaPaid

            End If

            ' Se invece siamo con uno storno a >= 0 da fare perchè tutti 
            ' l'erore è nella procedura di configuraione.
            Dim TotalToVoid As Decimal = CDec(sTotalVoid) '/ m_ParseFractMode

            If TotalToVoid >= 0 Then

                If False Then

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Void Argentea ::KO:: Local")

                    ' Signal (Storno non congruente per funzione mancante nella configurazione di sistema)
                    _SetOperationStatus(funcName,
                        NOT_ERROR_CALL_VOID_INCONGRUENT,
                        "Funzione di storno non più disponibile rispetto al totale di buoni già usati e stornati.",
                        PosDef.TARMessageTypes.TPERROR, True
                    )

                    Return IBPReturnCode.KO_CONTINUE_STANDARD

                    '' todo...
                    m_taobj.RemoveWithRefs(m_CurrMedia.theHdr.lTaCreateNmbr)
                    m_taobj.TARefresh(False)

                    If Me.DeleteOperation() Then
                        ' Ritorno comunque come oprazione Completata
                        Return IBPReturnCode.OK
                    Else
                        Return IBPReturnCode.KO_SKIP_STANDARD
                    End If

                Else

                    ''

                End If

            End If

            ' Questo è quello che si può pagare in totale.
            m_VoidAmount = +(TotalToVoid)

            ' Partiamo Che non è risucita
            Void = IBPReturnCode.KO

            '
            '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **
            '
            ' Richiamiamo il Metodo per l'Azione per visualizzarlo
            ' passando l'intera transazione il Controller corrente
            ' Il Metodo di pagamento passato come argomento
            '
            ' Se ci stiamo facendo un Annullo su un elemento che non è
            ' un Raggruppamento ed ha questa proprietà a  False allora
            ' possiamo direttamente richiamare l'API di Void Singola per
            ' il titolo selezionato in corso.b
            ' Altrimenti essendo un Ragguppamento ripresentiamo l'intero
            ' elenco dei titoli per dare iterazione all'utente di annullare
            ' scandedo i barcode o clicando sull'icona a lato di ogni singolo
            ' elemento presente in questo raggurppamento.
            '
            If CBool(m_CurrMedia.GetPropertybyName("bGroupedPaymentsOrder")) = False Then

                ' VOID singolo per questo Titolo Singolo

                Dim CurrentBarcode As String = m_CurrMedia.GetPropertybyName("szbp_grp_itm")
                Dim CurrentCrcTransactionId As String = m_CurrMedia.GetPropertybyName("szbp_grp_itm_IDTransaction") ' szbp_grp_itm_IDTransaction
                Dim CurrentTotValueBp As Decimal = CDec(m_CurrMedia.GetPropertybyName("dbp_grp_itm_Value"))

                If HandleSingleVoidBPCall(CurrentBarcode, CurrentCrcTransactionId, CurrentTotValueBp) Then

                    Void = IBPReturnCode.OK

                End If

            Else

                ' VOID con eleneco dei titoli raggrupapti

                If HandleVoidBPCall(m_VoidAmount) Then
                    Void = IBPReturnCode.OK
                Else
                    Void = IBPReturnCode.KO
                End If

            End If
            '
            '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

        Catch ex As Exception

            Void = IBPReturnCode.KO

            ' Signal (PEccezione interna non gestita in questa funzione)
            _SetOperationStatus(funcName,
                GLB_ERROR_NOT_UNEXPECTED,
                "Errore interno non previsto 2 -- exception on call void in controller bp with argentea -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPERROR, True
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

        Finally
            LOG_FuncExit(getLocationString(funcName), Void.ToString())
        End Try

    End Function


    Private Function DeleteOperation() As Boolean

        Try

            Dim MyTaMediaMemberDetailRec As TaFailedMediaRec = m_taobj.GetTALine(m_taobj.GetRefPositionFromCreationNmbr(TPDotnet.Pos.TARecTypes.iTA_MEDIA_MEMBER_DETAIL,
                                                                                m_CurrMedia.theHdr.lTaCreateNmbr,
                                                                                0))

            Dim sSelectedLine As Short = m_taobj.GetPositionFromCreationNmbr(m_CurrMedia.theHdr.lTaCreateNmbr)
            m_taobj.ChangeSign(sSelectedLine)

            'remove the previously added MediaStat-object
            m_taobj.Remove(m_taobj.GetPositionFromCreationNmbr(m_CurrMedia.theHdr.lTaCreateNmbr))

            m_taobj.TARefresh(False)
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function


    ''' <summary>
    '''     Azione di Chiusura
    '''     (per uso e consumo della clsEndTAHandling)
    '''     fa sì che si chiudano tutte le transazioni in corso per i BP
    '''     ti tipo TicketRestaurant denominati BPC (Buoni Pasti Cartacei)
    '''     chiamando l'API dedicata sul servizio di Argentea Close
    ''' </summary>
    ''' <param name="Parameters">
    '''     Il Set di parametri dinamici ad uso e consumo
    '''     del Controller che implementa il metodo passati
    '''     in modo dinamico previsti sul DB di BackStore
    ''' </param>
    ''' <param name="SilentMode">
    '''     Se mostrare o meno messaggi di errore o di stato
    ''' </param>
    ''' <returns>True se l'azione API ha dato esito OK altrimenti False</returns>
    Public Function Close(ByRef Parameters As Dictionary(Of String, Object), SilentMode As Boolean) As Boolean Implements IBPDematerialize.Close
        Dim funcName As String = "Close"

        '
        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPParameters.LoadCommonFunctionParameter
        '
        Dim Result As IBPReturnCode = _initializeParameters(funcName, Parameters, False)
        If Result = IBPReturnCode.KO Then
            Return False
        End If

        Try

            '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **
            '
            ' Richiamiamo il Metodo per l'Azione per visualizzarlo
            ' passando l'intera transazione il Controller corrente
            ' Il Metodo di chiusura come fine operazioni
            '
            Return HandleCloseBPCall(SilentMode)
            '
            '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

        Catch ex As Exception

            ' Signal (PEccezione interna non gestita in questa funzione)
            _SetOperationStatus(funcName,
                GLB_ERROR_NOT_UNEXPECTED,
                "Errore interno non previsto 3 -- exception on call void in controller bp with argentea -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPERROR
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            Return False

        Finally
            LOG_FuncExit(getLocationString(funcName), Close.ToString())
        End Try


    End Function

    ''' <summary>
    '''     Funzione helper per il caricamento ed assegnamento
    '''     dei parametri necessari ai metodi di interfaccia.
    ''' </summary>
    ''' <param name="parameters">I parametri dinamici da caricare ed associare tramite reflection</param>
    ''' <param name="v">Se mostrare o meno il messaggio di errore all'operatore</param>
    ''' <returns>OK o KO<see cref="IBPReturnCode"/></returns>
    Private Function _initializeParameters(ByRef funcName As String, ByRef parameters As Dictionary(Of String, Object), ForceShowMessageOperatorError As Boolean) As IBPReturnCode

        '
        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPParameters.LoadCommonFunctionParameter
        '
        pParams = New BPParameters()

        Try
            LOG_Debug(getLocationString(funcName), "We are entered In Argentea IBPDematerialize Function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            ' Recuperiamo e valorizziamo i parametri per reflection mode byref
            pParams.LoadCommonFunctionParameter(parameters)

            '
            ' Importanti per la gestione sul controller
            ' corrente  e  che non possono mancare mai.
            '
            m_TheModcntr = pParams.Controller
            m_taobj = pParams.Transaction
            m_CurrMedia = pParams.MediaRecord   ' <-- non necessariamente

            If m_TheModcntr Is Nothing Or m_taobj Is Nothing Then

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                ' Signal (Errore nell'istanza del proxy da eseguire)
                _SetOperationStatus(funcName,
                    OPR_ERROR_START_PROXY_INTERNAL,
                    "Errore di configurazione del prodotto BP - Funzione chiamata con parametri errati - (Chiamare assistenza).",
                    PosDef.TARMessageTypes.TPERROR, ForceShowMessageOperatorError
                )

                Return IBPReturnCode.KO

            End If

        Catch ex As Exception

            ' Signal (PEccezione interna non gestita in questa funzione)
            _SetOperationStatus(funcName,
                GLB_ERROR_NOT_UNEXPECTED,
                "Errore interno non previsto 4 -- exception on initialize parameter in controller bp with argentea -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPERROR, ForceShowMessageOperatorError
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            Return IBPReturnCode.KO

        Finally
            LOG_FuncExit(getLocationString(funcName), _initializeParameters.ToString())
        End Try


    End Function


#End Region

#Region "Handler Internal Form Action Overridable in classi Child"

    ''' <summary>
    '''     Handle per azionare il pagamento tramite Proxy
    '''     Con Form POS simulato per questa cassa.
    ''' </summary>
    ''' <returns>Boolean True (Conclusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandlePaymentBPCall(ByVal dPayableAmount As Decimal) As Boolean
        Dim funcName As String = "HandlePaymentBPCall"
        Dim proxyPos As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandlePaymentBPCall = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Recupero per la TA corrente il numero
            ' totale dei Buoni usati fino ad adesso
            ' per ottemperare all'opzione al numero
            ' massimo di buoni utilizzabili per  lo
            ' scontrino.
            '
            Dim TotalBPInUse As Integer = GetTotalBpUsedInCurrentTA()

            ' msgUtil.ShowMessage(m_TheModcntr, "Numero BP usati .:" + CStr(TotalBPInUse), "ll")

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al service di Argentea.
            '
            Dim NmRnd As Integer = New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer)).Next()
            'Dim NmRnd As Integer = 12345678

            If proxyPos Is Nothing Then
                proxyPos = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                NmRnd.ToString(),
                m_CurrMedia.dTaPaidTotal,               '   <-- Il Pagato fino adesso insieme agli altri media (Diventa il Pagabile nella sessione)
                TotalBPInUse                           '   <-- Il Numero di BP usati in questo scontrino fino ad adesso
            )
                'm_taobj.GetPayedValue()                 '   <-- Il Totale complessivo della TA
                'pParams.TransactionID,                  '   <-- L'id della transazione in corso

                '
                ' Preparo ad accettare l'handler degli eventi gestiti
                ' prima e dopo la comunicazione con il POS locale.
                '
                AddHandler proxyPos.Event_ProxyCollectDataTotalsAtEnd, AddressOf ProxyCollectDataTotalsAtEnd_Handler

            End If

            ' PAYMENT Service Mode
            proxyPos.Command = ClsProxyArgentea.enCommandToCall.Payment

            ' SILENT MODE o meno
            proxyPos.SilentMode = Me.SilentMode

            '
            ' Preparo l'oggetto a quello che si deve a  spettare
            ' come totale da pagare e quello pagabile
            '
            proxyPos.AmountPaid = 0
            proxyPos.AmountPayable = dPayableAmount

            ' Definisco questa  variabile  Privata 
            ' per il conteggio dei Buoni eventuali
            ' già presenti nella TA che sono stati
            ' usati in precedenza dall'operatore.
            _InitialBPPayed = pParams.GetAlreadyOnTAScanned()

        Catch ex As Exception

            'LOG_Debug(getLocationString(funcName), "Instance proxyPOS Argentea ::KO:: Local")

            ' Signal (Errore non previsto)
            _SetOperationStatus(funcName,
                GLB_ERROR_INSTANCE_SERVICE,
                "Errore interno non previsto A -- exception on initialize class Argentea controller -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            Return False

        Finally
            ''
        End Try


        ' >>>> ***************************************** <<<<<<
        '
        ' RUN -> Avvio il FORM Locale ed attendo!! with try entrapment
        '
        If Not proxyPos.IsLive Then
            proxyPos.Connect()
        Else
            proxyPos.Unpark()
        End If
        '
        ' >>>> ***************************************** <<<<<<

        ' Del resto concludo
        Dim StatusResult As ClsProxyArgentea.enProxyStatus = proxyPos.ProxyStatus
        proxyPos.Close()

        ' E restituisco 
        If StatusResult = ClsProxyArgentea.enProxyStatus.OK Then

            ' Tutto è filato come doveva e  le 
            ' operazioni sono conformi per cio
            ' che era previsto.
            Return True

        Else ' KO Error o altro stato che l'ha fatto finire

            ' Può essere che c'è stato un  non
            ' OK in qualche procedura che  non
            ' valida la transazione.
            Return False

        End If

    End Function

    ''' <summary>
    '''     Handle per azionare lo storno tramite Form
    '''     POS simulato in questa cassa.
    ''' </summary>
    ''' <returns>Boolean True (Conslusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandleVoidBPCall(ByVal dVoidableAmount As Decimal) As Boolean
        Dim funcName As String = "HandleVoidBPCall"
        Dim proxyPos As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandleVoidBPCall = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Recupero per la TA corrente il numero
            ' totale dei Buoni usati fino ad adesso
            ' per ottemperare all'opzione al numero
            ' massimo di buoni utilizzabili per  lo
            ' scontrino.
            '
            Dim TotalBPInUse As Integer = GetTotalBpUsedInCurrentTA()

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al proxy di Argentea.
            '
            Dim NmRnd As Integer = New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer)).Next()
            'Dim NmRnd As Integer = 12345678
            If proxyPos Is Nothing Then

                proxyPos = New ClsProxyArgentea(
                    m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                    m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                    MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                    NmRnd.ToString(),                       '   <-- L'Id della sessione per questo Gruppo BP
                    m_CurrMedia.dTaPaidTotal,               '   <-- Il Pagabile con questo Media
                    TotalBPInUse,                           '   <-- Il Numero Totale di BP nell'intera TA Corrente
                    m_taobj.GetPayedValue()                 '   <-- Il Totale sulla TA corrente da pagare fino ad adesso
                )
                'pParams.TransactionID,                  '   <-- L'id della transazione in corso

                '
                ' Preparo ad accettare l'handler degli eventi gestiti
                ' prima e dopo la comunicazione con il POS locale.
                '
                AddHandler proxyPos.Event_ProxyCollectDataVoidedAtEnd, AddressOf ProxyCollectDataVoidedAtEnd_Handler

            End If

            ' VOID Proxy Service Mode
            proxyPos.Command = ClsProxyArgentea.enCommandToCall.Void

            ' SILENT Mode o meno
            proxyPos.SilentMode = SilentMode

            '
            ' Dato che è uno storno vediamo se sul POS
            ' di servizio vogliamo ripresentare  i  BP
            ' se questi sono in un  raggruppamento  da 
            ' usare su un FORM prefillato con l'elenco.
            '
            If Not m_CurrMedia Is Nothing Then
                proxyPos.PrefillVoidable = LoadAndFillListOfGroupItems(proxyPos.FractParsing)
            End If

            '
            ' Preparo l'oggetto a quello che si deve a  spettare
            ' come totale da stornare e quello stornato
            '
            proxyPos.AmountVoidable = -dVoidableAmount

            ' Definisco questa  variabile  Privata 
            ' per il conteggio dei Buoni eventuali
            ' già presenti nella TA che sono stati
            ' usati in precedenza dall'operatore.
            _InitialBPPayed = pParams.GetAlreadyOnTAScanned()

        Catch ex As Exception

            LOG_Debug(getLocationString(funcName), "Instance proxyPos Argentea ::KO:: Local")

            ' Signal (Errore non previsto)
            _SetOperationStatus(funcName,
                GLB_ERROR_INSTANCE_SERVICE,
                "Errore interno non previsto B -- exception on initialize class Argentea controller -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            Return False

        Finally
            ''
        End Try

        ' >>>> ***************************************** <<<<<<
        '
        ' RUN -> Avvio il FORM Locale ed attendo!! with try entrapment
        '
        If Not proxyPos.IsLive Then
            proxyPos.Connect()
        Else
            proxyPos.Unpark()
        End If
        '
        ' >>>> ***************************************** <<<<<<

        ' Del resto concludo
        Dim StatusResult As ClsProxyArgentea.enProxyStatus = proxyPos.ProxyStatus
        proxyPos.Close()

        ' E restituisco 
        If StatusResult = ClsProxyArgentea.enProxyStatus.OK Then

            ' Tutto è filato come doveva e  le 
            ' operazioni sono conformi per cio
            ' che era previsto.
            Return True

        ElseIf StatusResult = ClsProxyArgentea.enProxyStatus.InError Then

            ' Non elimino la TA e la lascio inalterata 
            'm_taobj.RemoveWithRefs(m_CurrMedia.theHdr.lTaCreateNmbr)
            'm_taobj.TARefresh(False)

        Else

            ' Può essere che c'è stato un  non
            ' OK in qualche procedura che  non
            ' valida la transazione.
            Return False

        End If

    End Function


    ''' <summary>
    '''     Handle per eseguire il comando di VoidSingle
    '''     sul servizio Argentea Remoto come chiamata a 
    '''     storno su un titolo in base al suo id di transazione.
    ''' </summary>
    ''' <returns>Boolean True (Conslusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandleSingleVoidBPCall(barcode As String, idCrcTransaction As String, TotValueOfBP As Decimal) As Boolean
        Dim funcName As String = "HandleSingleVoidBPCall"
        Dim proxyPos As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandleSingleVoidBPCall = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al proxy di Argentea.
            '
            Dim NmRnd As Integer = New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer)).Next()

            'Dim NmRnd As Integer = 12345678
            If proxyPos Is Nothing Then
                proxyPos = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                NmRnd.ToString(),
                m_CurrMedia.dTaPaidTotal,,              '   <-- Il Pagabile di questo media
                m_taobj.GetPayedValue()                 '   <-- Il Totale della TA ancora da pagare
            )

            End If

            ' SILENT Mode o meno
            proxyPos.SilentMode = SilentMode

        Catch ex As Exception

            LOG_Debug(getLocationString(funcName), "Instance proxyPos Argentea ::KO:: Local")

            ' Signal (Errore non previsto)
            _SetOperationStatus(funcName,
                GLB_ERROR_INSTANCE_SERVICE,
                "Errore interno non previsto C -- exception on initialize class Argentea controller -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            Return False

        Finally
            ''
        End Try

        ' >>>> ***************************************** <<<<<<
        '
        ' CALL -> Esecuzione dell'API sul sistema service Remoto!! with response entrapment
        '
        Dim _ArgBarCode As KeyValuePair(Of String, Object) = New KeyValuePair(Of String, Object)("BarCode", barcode)
        Dim _IdCrcTransaction As KeyValuePair(Of String, Object) = New KeyValuePair(Of String, Object)("IdCrcTransaction", idCrcTransaction)
        Dim _TotValueItem As KeyValuePair(Of String, Object) = New KeyValuePair(Of String, Object)("TotFaceValue", TotValueOfBP)

        Dim StatusResult As ClsProxyArgentea.enProxyStatus = proxyPos.CallAPI("SINGLEVOID", _ArgBarCode, _IdCrcTransaction, _TotValueItem)
        '
        ' >>>> ***************************************** <<<<<<

        ' E restituisco 
        If StatusResult = ClsProxyArgentea.enProxyStatus.OK Then

            ' Tutto è filato come doveva e  le 
            ' operazioni sono conformi per cio
            ' che era previsto.
            Return True

        Else

            ' Può essere che c'è stato un  non
            ' OK in qualche procedura che  non
            ' valida la transazione.
            Return False

        End If

    End Function

    ''' <summary>
    '''     Handle per eseguire il comando di Close
    '''     sul servizio Argentea Remoto come chiamata a 
    '''     chiusura delle operazioni effettutate fino ad adesso.
    ''' </summary>
    ''' <param name="SilentMode">Se avviare il Proxy con messaggi e segnalazioni utente o in modo silenzione senza messaggi</param>
    ''' <returns>Boolean True (Conslusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandleCloseBPCall(SilentMode As Boolean) As Boolean
        Dim funcName As String = "HandleCloseBPCall"
        Dim proxyPos As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandleCloseBPCall = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al proxy di Argentea.
            '
            Dim NmRnd As Integer = New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer)).Next()

            'Dim NmRnd As Integer = 12345678
            If proxyPos Is Nothing Then
                proxyPos = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                NmRnd.ToString(),
                m_CurrMedia.dTaPaidTotal,,              '   <-- Il Pagabile di questo media
                m_taobj.GetPayedValue()                 '   <-- Il Totale complessivo della TA
            )

            End If

            ' SILENT Mode o meno
            proxyPos.SilentMode = SilentMode

        Catch ex As Exception

            LOG_Debug(getLocationString(funcName), "Instance proxyPos Argentea ::KO:: Local")

            ' Signal (Errore non previsto)
            _SetOperationStatus(funcName,
                GLB_ERROR_INSTANCE_SERVICE,
                "Errore interno non previsto D -- exception on initialize class Argentea controller -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            Return False

        Finally
            ''
        End Try

        ' >>>> ***************************************** <<<<<<
        '
        ' CALL -> Esecuzione dell'API sul sistema service Remoto!! with response entrapment
        '
        Dim StatusResult As ClsProxyArgentea.enProxyStatus = proxyPos.CallAPI("Close")
        '
        ' >>>> ***************************************** <<<<<<

        ' E restituisco 
        If StatusResult = ClsProxyArgentea.enProxyStatus.OK Then

            ' Tutto è filato come doveva e  le 
            ' operazioni sono conformi per cio
            ' che era previsto.
            Return True

        Else

            ' Può essere che c'è stato un  non
            ' OK in qualche procedura che  non
            ' valida la transazione.
            Return False

        End If

    End Function


    ''' <summary>
    '''     Preparaun Dictionari di valori da prefillare sul
    '''     Form POS Software o Hardware in relazione a  uno
    '''     Storno in corso da parte di un raggruppamento di
    '''     Pagato.
    ''' </summary>
    ''' <param name="FractToValues">Il valore per la frazione da centesimi in euro usata dal protocollo</param>
    ''' <returns></returns>
    Private Function LoadAndFillListOfGroupItems(FractToValues As Integer) As Dictionary(Of String, PaidEntry)
        Dim _NumCurrT As Integer = 1
        Dim KeyCBP As String
        Dim KeyVBP As String
        Dim _ARVoided As PaidEntry

        LoadAndFillListOfGroupItems = Nothing

        '
        If m_CurrMedia.GetPropertybyName("ibp_GROUPED") = 1 Then

            '
            '   Di tipo 1 Sono i BP cartacei solitamente
            '   listati e non raggruppati nei Metatag
            '

            Dim itmBarCode As String
            Dim itmValueDc As Decimal
            Dim itmFaceVal As Decimal
            Dim itmIDCrcTr As String
            Dim itmOtherMd As String
            Dim itmEmitter As String
            Dim itmCodIssr As String
            Dim itmNamIssr As String
            Dim itmToPass As PaidEntry

            LoadAndFillListOfGroupItems = New Dictionary(Of String, PaidEntry)

            Dim _TotBpUsed As Integer = m_CurrMedia.GetPropertybyName("ibp_TOT_BP_USED")

            If Not m_CurrMedia.GetPropertybyName("ibp_TOT_BP_VOIDED") = "" Then
                _TotBpUsed -= m_CurrMedia.GetPropertybyName("ibp_TOT_BP_VOIDED")
            End If

            For x As Integer = 0 To _TotBpUsed - 1

                '
                '   Elenco come metatag ogni singolo bp
                '   con relativo barcode usato nell'insieme
                '   di quelli elaborati per pagare.
                '
                KeyCBP = "bp_itm_" + CStr(x + 1)
                itmBarCode = m_CurrMedia.GetPropertybyName("sz" & KeyCBP)
                itmValueDc = m_CurrMedia.GetPropertybyName("d" & KeyCBP + "_Value") '* FractToValues
                itmFaceVal = m_CurrMedia.GetPropertybyName("d" & KeyCBP + "_FaceValue") '* FractToValues
                itmIDCrcTr = m_CurrMedia.GetPropertybyName("sz" & KeyCBP + "_IDTransaction")
                itmOtherMd = m_CurrMedia.GetPropertybyName("sz" & KeyCBP + "_OtherInfo")
                itmEmitter = itmOtherMd.Split("-")(0)
                itmCodIssr = itmOtherMd.Split("-")(1)
                itmNamIssr = itmOtherMd.Split("-")(2)

                itmToPass = New PaidEntry(
                                itmBarCode,
                                CStr(itmValueDc),
                                CStr(itmFaceVal),
                                itmEmitter,
                                itmIDCrcTr
                )
                itmToPass.CodeIssuer = itmCodIssr
                itmToPass.NameIssuer = itmNamIssr

                LoadAndFillListOfGroupItems.Add(itmBarCode, itmToPass)

                _NumCurrT += 1

            Next

            '
            '   Per riporto nelle sessioni successive riporto 
            '   anche quelli che in uno storno precedente  ha
            '   già presente nei Metatag (Per ricomporlo in storico all'uscita)
            '
            If Not m_CurrMedia.GetPropertybyName("ibp_TOT_BP_VOIDED") = "" Then

                For x As Integer = 0 To m_CurrMedia.GetPropertybyName("ibp_TOT_BP_VOIDED") - 1

                    KeyVBP = "bp_itm_voided_" + CStr(x + 1)
                    itmBarCode = m_CurrMedia.GetPropertybyName("sz" & KeyVBP)
                    itmValueDc = m_CurrMedia.GetPropertybyName("d" & KeyVBP + "_Value") '/ FractToValues
                    itmFaceVal = m_CurrMedia.GetPropertybyName("d" & KeyVBP + "_FaceValue") '/ FractToValues
                    itmIDCrcTr = m_CurrMedia.GetPropertybyName("sz" & KeyVBP + "_IDTransaction")
                    itmOtherMd = m_CurrMedia.GetPropertybyName("sz" & KeyVBP + "_OtherInfo")
                    itmEmitter = itmOtherMd.Split("-")(0)
                    itmCodIssr = itmOtherMd.Split("-")(1)
                    itmNamIssr = itmOtherMd.Split("-")(2)

                    itmToPass = New PaidEntry(
                                itmBarCode,
                                CStr(itmValueDc),
                                CStr(itmFaceVal),
                                itmEmitter,
                                itmIDCrcTr
                    )
                    itmToPass.CodeIssuer = itmCodIssr
                    itmToPass.NameIssuer = itmNamIssr
                    itmToPass.Voided = True ' --> VOIDED -> Stornato

                    LoadAndFillListOfGroupItems.Add(itmBarCode, itmToPass)
                    _NumCurrT += 1

                Next

            End If


        ElseIf m_CurrMedia.GetPropertybyName("ibp_GROUPED") = 2 Then

            '
            '   Di tipo 2 Sono i BP elettronici solitamente
            '   raggruppati per taglio e numero di pezzi nei Metatag
            '

            LoadAndFillListOfGroupItems = New Dictionary(Of String, PaidEntry)

            Dim KeyQTA As String
            Dim KeyTOT As String
            Dim KeyVALUE As String

            Dim TypeEdgesDemat As String() = m_CurrMedia.GetPropertybyName("szbp_EDGES_DEMAT").Split("|")

            For x As Integer = 1 To UBound(TypeEdgesDemat) '- 1

                ' Elemento Cntabilizzato
                KeyQTA = "lBP_QUANTITY_C_" + x.ToString
                KeyTOT = "dBP_AMOUNT_C_" + x.ToString
                KeyVALUE = "dBP_VALUE_C_" + x.ToString

                Dim ValQTA As Integer = m_CurrMedia.GetPropertybyName(KeyQTA)                       ' Pezzi
                Dim ValTOT As Decimal = m_CurrMedia.GetPropertybyName(KeyTOT) / FractToValues       ' Totale
                Dim ValVALUE As Decimal = m_CurrMedia.GetPropertybyName(KeyVALUE) / FractToValues   ' Taglio Chiave

                For y As Integer = 1 To ValQTA
                    Dim KeyItm As String = CStr(ValVALUE).Replace(".", "_").Replace(",", "_") + "_" + CStr(y)
                    LoadAndFillListOfGroupItems.Add("bp_demat_" + KeyItm, New PaidEntry("bp_demat_" + KeyItm, ValVALUE))
                Next

            Next

            Dim TypeEdgesVoid As String() = m_CurrMedia.GetPropertybyName("szbp_EDGES_VOID").Split("|")

            For x As Integer = 1 To UBound(TypeEdgesVoid) '- 1

                'Elemento Stornato
                KeyQTA = "lBP_QUANTITY_V_" + x.ToString
                KeyTOT = "dBP_AMOUNT_V_" + x.ToString
                KeyVALUE = "dBP_VALUE_V_" + x.ToString

                Dim ValQTA As Integer = m_CurrMedia.GetPropertybyName(KeyQTA)                       ' Pezzi
                Dim ValTOT As Decimal = m_CurrMedia.GetPropertybyName(KeyTOT) / FractToValues       ' Totale
                Dim ValVALUE As Decimal = m_CurrMedia.GetPropertybyName(KeyVALUE) / FractToValues   ' Taglio Chiave

                For y As Integer = 1 To ValQTA
                    Dim KeyItm As String = CStr(ValVALUE).Replace(".", "_").Replace(",", "_") + "_" + CStr(y)
                    Dim NewPaid As PaidEntry = New PaidEntry("bp_void_" + KeyItm, ValVALUE)
                    NewPaid.Voided = True
                    LoadAndFillListOfGroupItems.Add("bp_void_" + KeyItm, NewPaid)
                Next

            Next

            Dim TypeEdgesInvalid As String() = m_CurrMedia.GetPropertybyName("szbp_EDGES_INVALID").Split("|")

            For x As Integer = 1 To UBound(TypeEdgesInvalid) '- 1

                'Elemento Stornato
                KeyQTA = "lBP_QUANTITY_E_" + x.ToString
                KeyTOT = "dBP_AMOUNT_E_" + x.ToString
                KeyVALUE = "dBP_VALUE_E_" + x.ToString

                Dim ValQTA As Integer = m_CurrMedia.GetPropertybyName(KeyQTA)                       ' Pezzi
                Dim ValTOT As Decimal = m_CurrMedia.GetPropertybyName(KeyTOT) / FractToValues       ' Totale
                Dim ValVALUE As Decimal = m_CurrMedia.GetPropertybyName(KeyVALUE) / FractToValues   ' Taglio Chiave

                For y As Integer = 1 To ValQTA
                    Dim KeyItm As String = CStr(ValVALUE).Replace(".", "_").Replace(",", "_") + "_" + CStr(y)
                    Dim NewPaid As PaidEntry = New PaidEntry("bp_invalid_" + KeyItm, ValVALUE)
                    NewPaid.Invalid = True
                    NewPaid.InfoExtra = "Element non valid and not contabilizated"
                    LoadAndFillListOfGroupItems.Add("bp_invalid_" + KeyItm, NewPaid)
                Next

            Next

        End If
    End Function

    ''' <summary>
    '''     Restituisce attraversando la TA corrente
    '''     il numero dei buoni usati nel contesto.
    ''' </summary>
    ''' <returns>Il Numero dei Buoni usati nel contesto dello scontrino corrente</returns>
    Private Function GetTotalBpUsedInCurrentTA() As Integer
        Dim funcName As String = "GetTotalBpUsedInCurrentTA"
        Dim _Count As Integer = 0

        Dim KeyTotBp As String = "ibp_TOT_BP_USED"
        Dim KeyTotVd As String = "ibp_TOT_BP_VOIDED"

        Try
            Dim TARecord As TaBaseRec
            For Each TARecord In m_taobj.taCollection
                If TARecord.sid = PosDef.TARecTypes.iTA_MEDIA Then

                    Dim CurrMedia As TaMediaRec = CType(TARecord, TaMediaRec)

                    ' Se ha questa Proprietà Chiave  è  un  Tipo  Gestito
                    ' come Titolo di Buono Pasto sullo scontrino corrente.
                    '
                    If Not CurrMedia.GetPropertybyName("bGroupedPaymentsOrder") = "" Then

                        If CBool(CurrMedia.GetPropertybyName("bGroupedPaymentsOrder")) = False Then
                            _Count += 1
                        Else

                            If Not CurrMedia.GetPropertybyName(KeyTotBp) = "" Then
                                _Count += CurrMedia.GetPropertybyName(KeyTotBp)
                            End If
                            If Not CurrMedia.GetPropertybyName(KeyTotVd) = "" Then
                                _Count -= CurrMedia.GetPropertybyName(KeyTotVd)
                            End If

                            ' Per i pagamenti elettronici etichettati con questo
                            ' attributo al momento del stream su xml il conteggio
                            ' dei stornati è da intendere a conteggio su tutto quindi
                            ' da fare x 2 in modo che scarti il numero totale di quelli 
                            ' che inizialmente sono stati usati
                            'If CurrMedia.GetPropertybyName("szTypePaymentOrder") = ClsProxyArgentea.enTypeBP.TicketsCard.ToString() Then
                            'If Not CurrMedia.GetPropertybyName(KeyTotVd) = "" Then
                            '_Count -= CurrMedia.GetPropertybyName(KeyTotVd)
                            'End If
                            'End If

                        End If

                    End If

                End If

            Next TARecord

            Return _Count

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function " + funcName + " returns " & _Count.ToString)
        End Try

    End Function

#End Region

#Region "Eventi che arrivano dal Proxy di servizio"

    ''' <summary>
    '''     Il totale di Pagamenti effettuati con BP e e C
    '''     una volta che il proxy ha finito con le sue
    '''     operazioni di scansione o chiamata a pos locale.
    ''' </summary>
    ''' <param name="sender">Il Proxy</param>
    ''' <param name="resultData">I Codidi EAN e importo utilizzati per l'intero importo pagato</param>
    Private Sub ProxyCollectDataTotalsAtEnd_Handler(ByRef sender As Object, ByRef resultData As ClsProxyArgentea.DataResponse)

        Dim NewTaMediaRec As TaMediaRec
        Dim PeExcedeedRec As TaMediaRec = Nothing
        Dim RemoveFirstOriginalMedia As Boolean = False
        Dim funcName As String = "Event_CollectDataTotals_BP"

        If resultData.totalBPElaboratedInCurrentSession = 0 And resultData.totalBPUsedToPay = 0 And
             resultData.totalVoidedWithBP = 0 Then

            '
            ' Dato che stiamo parlando di uno storno BPE
            ' controlliamo che se ci sono stati degli storni
            ' Rimuovo la TA originale per ricrearla
            '

            ' Se tutti sono stati eliminati
            ' dato che la TA è stata eliminata esco
            m_taobj.RemoveWithRefs(m_CurrMedia.theHdr.lTaCreateNmbr)
            m_taobj.TARefresh(False)
            Return

        End If

        '
        ' GESTIONE DEL MEDIA SULLA
        ' TRANSAZIONE IN CORSO.
        '
        Try

            '
            ' Opzione Operatività.: 
            '       Se vogliamo che sia accorpato l'elenco dei Media in 
            '       elenco sulla transazione o meno.
            Dim OptAccorpateMediaForBP As Boolean = CType(sender, ClsProxyArgentea).ArgenteaParameters.BP_AccorpateOnTA

            ' Nel caso di BPE sarà sempre raggruppato
            If resultData.typeBPElaborated = ClsProxyArgentea.enTypeBP.TicketsCard Then
                OptAccorpateMediaForBP = True
            End If
            'OptAccorpateMediaForBP = False

            If OptAccorpateMediaForBP Then

                '
                ' Aggiungo una sola TA con un set di 
                ' BP riepilogati per Taglio che sono
                ' stati elaborati nel processo sul proxy
                '
                '       NOTA.: Viene creato Un solo MediaRecord con un riepilogo dei BP utilizzati.
                '

                ' Creo un Id di riferimento per l'ELemento con il Raggruppamento tramite MetaData
                Dim NmRnd As Integer = New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer)).Next()

                ' Le'elemtno Gruppo Singolo Pagante Ripreso da >>>( resultData.totalPayedWithBP )<<<<
                Dim ItmPe As PaidEntry = New PaidEntry("Group_BP_" & resultData.typeBPElaborated.ToString(), resultData.totalPayedWithBP)

                ' Riporto come Id di transazione per questo titolo (raggruppato) l'id univoco
                ItmPe.IDTransactionCrc = NmRnd.ToString()
                ItmPe.InfoExtra = CStr(resultData.totalBPUsedToPay) + "/" + CStr(resultData.totalBPElaboratedInCurrentSession)      ' <- Mi riporto il numero dei titoli che sono raggruppati

                ' Aggiungo sulla Transazione corrente
                ' la TA relativa al Media di pagamento.
                NewTaMediaRec = AddNewTaMedia(ItmPe, resultData.typeBPElaborated.ToString(), True)

                '
                ' Aggiungo il Riepilogo alla TA appena creata
                ' e controllo all'uscita se mi ha prodotto un
                ' resto da accodare come Media al pagamento
                '
                If resultData.typeBPElaborated = ClsProxyArgentea.enTypeBP.TicketsCard Then

                    ' Per i Buoni Pasti Elettronici riepilogo
                    ' per Taglio e valuta il numero dei conumati
                    PeExcedeedRec = SaveAndPutItemsOnMediaSingleBPE(NewTaMediaRec, resultData)

                ElseIf resultData.typeBPElaborated = ClsProxyArgentea.enTypeBP.TicketsRestaurant Then

                    ' Peri Buoni Pasti Cartacei riepilogo 
                    ' per ogni singolo Buono con riportato
                    ' il suo relativo BarCode utilizzato.
                    PeExcedeedRec = SaveAndPutItemsOnMediaSingleBPC(NewTaMediaRec, resultData)

                End If

                '
                ' Aggiungo alla transazione l'elenco
                ' del media record di tipo BP appena 
                ' creato con i dati riportati.
                '

                m_taobj.Add(NewTaMediaRec)


            Else

                '
                ' Aggiungo tanti TA per ogni voce  di
                ' BP elaborato nel processo sul proxy
                '
                '       NOTA.: Vengono creati tanti MediaRecord uno per ogni Buono utilizzato.
                '
                ' E controllo all'uscita se mi ha prodotto un
                ' resto da accodare come Media al pagamento
                '

                ' Creo un Id di riferimento per ogni Titolo che va a finire nella TA per riferimento alla sessione che li ha creati
                Dim NmRnd As Integer = New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer)).Next()
                Dim idGroupSessionReferement As String = "Group_Session_" & resultData.typeBPElaborated.ToString() & "_" & NmRnd.ToString()

                ' Ed aggiungo tanti elementi nella TA quanti sono i BP nella sessione che sono stati usati
                PeExcedeedRec = SaveAndPutItemsOnMediaGroup(m_taobj, resultData, idGroupSessionReferement, resultData.totalBPUsedToPay)

            End If

            '
            ' Se la voce di Resto per eccesso è stata
            ' creata l'aggiungo nella coda della transazione
            ' corrente.
            '
            If Not PeExcedeedRec Is Nothing Then
                m_taobj.Add(PeExcedeedRec)
            End If

            '
            ' Quindi se almeno un elemento è stato usato
            ' come pagamento in coda rimuovo la MEDIA che
            ' in origine è stata creata nella sessione 
            ' precedente.
            '
            If CType(sender, ClsProxyArgentea).Command = ClsProxyArgentea.enCommandToCall.Payment Then

                ' In Caso di Pagametno in Buoni rimuoviamo quella con cui siamo entrati in origine
                m_taobj.RemoveWithRefs(m_CurrMedia.theHdr.lTaCreateNmbr)

            ElseIf CType(sender, ClsProxyArgentea).Command = ClsProxyArgentea.enCommandToCall.Void Then

                ' In Caso di Storno di Buoni Pagati rimuoviamo lavoriamo direttamente su questa
                LOG_Debug(funcName, "Called renew Void for bp " & CStr(m_CurrMedia.theHdr.lTaCreateNmbr))

            End If

            ' Refresh del record
            m_taobj.TARefresh(False)

        Catch ex As Exception

            ' Etichettiamo l'errore per la gestione
            m_LastStatus = GLB_ERROR_COLLECT_DATA_SERVICE
            m_LastErrorMessage = "Eccezione interna non gestita -- Exception on data collector to set data result from Proxy -- (Chiamare assistenza)"

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            ' Forzo il Proxy allo stato di errore in corso
            CType(sender, ClsProxyArgentea).SetStatusInError(
                funcName,
                m_LastStatus,
                m_LastErrorMessage,
                TARMessageTypes.TPERROR
            )

        Finally

            LOG_FuncExit(getLocationString(funcName), "Media Original : " + (Not RemoveFirstOriginalMedia).ToString)

        End Try

    End Sub

    ''' <summary>
    '''     Il totale di Pagamenti stornati con BP e e C
    '''     una volta che il proxy ha finito con le sue
    '''     operazioni di scansione o chiamata a pos locale.
    ''' </summary>
    ''' <param name="sender">Il Proxy</param>
    ''' <param name="resultData">I Codidi EAN e importo utilizzati per l'intero importo stornato</param>
    Private Sub ProxyCollectDataVoidedAtEnd_Handler(ByRef sender As Object, ByRef resultData As ClsProxyArgentea.DataResponse)

        If resultData Is Nothing Then

            ' Nel Proxy qualcosa è andato in errore interno
            Return

        End If

        '
        ' Quindi Richiamo lo stesso Handler di evento di aggiunta
        ' dato che mi deve ricreare con tutto il metadata il record
        '
        ProxyCollectDataTotalsAtEnd_Handler(sender, resultData)

    End Sub

#End Region

#Region "Functions Common"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

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

        ' Imposta l'ultimo stato  corrente  per 
        ' l'uso successivo alle funzioni di chi
        ' esce.
        If constType Is Nothing And msgDefault Is Nothing Then

            ' In questo caso non imposta l'ultimo stato (usato solo per rinotificare)
            c_StatusMessage = m_LastStatus

        Else

            ' Per gli errori di eccezione non previste o previste 
            ' mostriamo i messaggi non codificabili nel db  della
            ' gestione per lingue e personalizzazioni
            If constType = GLB_ERROR_COLLECT_DATA_SERVICE Or
                constType = GLB_ERROR_INSTANCE_SERVICE Or
                constType = GLB_ERROR_NOT_UNEXPECTED Or
                constType = INT_ERROR_NOT_CREATE_EXCEDEED Or
                constType = GLB_ERROR_ON_CREATE_EXCEDEED Then

                ' In questo caso non c'è da tradurre mostriamo grezzo il messaggio come arriva
                m_LastStatus = constType
                c_StatusMessage = Nothing
                m_LastErrorMessage = msgDefault

            Else

                ' Per questa tipologia codifichiamo la msgbox al  fine
                ' di riprendere la costante dello stato operazione non
                ' valida per personalizzare sul db i messaggi

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
                LOG_Error(getLocationString(funcName), m_LastErrorMessage + InfoExtraMessageStatus + " <-:: Exception not Coded ::-> " + "LevelITCommonModArgentea_" + c_StatusMessage)
            Else
                LOG_Error(getLocationString(funcName), m_LastErrorMessage + InfoExtraMessageStatus + " <-:: Voce DB x Tradurre ::-> " + "LevelITCommonModArgentea_" + c_StatusMessage)
            End If

            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage + InfoExtraMessageStatus, "LevelITCommonModArgentea_" + m_LastStatus, TypeStatusMsgBox)

        Else

            ' Scrive una riga di Log per monitorare....
            LOG_Info(getLocationString(funcName), m_LastStatus + " -> " + m_LastErrorMessage + " " + InfoExtraMessageStatus)

        End If

    End Sub

#End Region

#Region "Functions Private"

    ''' <summary>
    '''     Crea un tipo MediaRec compatibile
    '''     per essere inserito nella Transazione
    '''     in corso come riga di pagamento.
    ''' </summary>
    ''' <param name="PeTo">Il new entry di tipo Pagamento In totale se raggruppato o lo Stesso Importo del titolo se singolo<see cref="PaidEntry"/></param>
    ''' <param name="GroupElement">True = Un Gruppo di titoli che formano un Totale di Riga inteso come pagato - False = Un Singolo elemento dove il suo Importo corrispone al totale di Riga come pagato</param>
    ''' <returns>Una Media compatibile con la transazione <see cref="TaMediaRec"/></returns>
    Private Function AddNewTaMedia(PeTo As PaidEntry, TypePaymentOrder As String, GroupElement As Boolean) As TaMediaRec

        ' Preparo la MediaRecord iniziale con cui
        ' sono entrato per farne un clone di rimpiazzo.
        Dim NewTaMediaRec As TaMediaRec = m_taobj.CreateTaObject(Of TaMediaRec)(PosDef.TARecTypes.iTA_MEDIA)

        ' Ne clono e aggiungo le propietà
        ' relative a quelle del Buono Pasto.
        With NewTaMediaRec

            ' Riprendo il MediaRecord iniziale 
            ' all'ingresso della funzione.
            .Clone(m_CurrMedia, m_CurrMedia.theHdr.lTaCreateNmbr)

            ' Porto a non linkato l'intestazione
            ' del nodo che non dipende da altri.
            .theHdr.lTaRefToCreateNmbr = 0
            .theHdr.lTaCreateNmbr = 0

            ' CODICE TIPOLOGIA e RAGGRUPPAMENTO(SI/NO) DEL TITOLO/I

            ' Codice del Titolo di Pagamento BPC (Buoni/o Pasto Cartacei/o) BPE (Buoni/o Basto Elettronici/o) (Usato dalla procedura corrente)
            .AddField("szCodePaymentOrder", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("szCodePaymentOrder", TYPE_SPECIFIQUE)

            ' Tipo di Titolo di Pagamento TicketsRestaurant (Buoni/o Pasto Cartacei/o) TicketsCard (Buoni/o Basto Elettronici/o) (Usato dal Proxy di Argentea)
            .AddField("szTypePaymentOrder", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("szTypePaymentOrder", TypePaymentOrder)

            ' Titolo di pagamento Raggruppato o Singolo elemento
            .AddField("bGroupedPaymentsOrder", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("bGroupedPaymentsOrder", CStr(GroupElement))

            ' Se il Titolo di pagamento è un Raggruppamento nelle info extra ho il numero di titoli che sono stati usati
            .AddField("szGroupedNumPayments", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("szGroupedNumPayments", PeTo.InfoExtra)

            ' Il Valore Facciale del Titolo o il Totale dei Valori Facciali dei Titoli raggruppati
            .AddField("szFaceValue", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("szFaceValue", PeTo.FaceValue)

            'CP#1337781:1:  default there's not rest, the mv was truncated
            .dTaQty = 1                                                         ' Che sia un solo titolo o un raggruppamento la quantità come pagato è sempre di 1
            .dTaPaid = Convert.ToDecimal(PeTo.Value)
            .dTaPaidTotal = Convert.ToDecimal(PeTo.Value)
            .dPaidForeignCurr = Convert.ToDecimal(PeTo.Value)
            .dPaidForeignCurrTotal = Convert.ToDecimal(PeTo.Value)

        End With

        Return NewTaMediaRec

    End Function

    ''' <summary>
    '''     Crea per una TA una serie di voci
    '''     MetaData con info di riepilogo su
    '''     un set di BP processati tramite il
    '''     proxy di Argentea suddivisi per Taglio Cutoff
    ''' </summary>
    ''' <param name="RootTaMediaRec">La TA di tipo Pagamento dove posizionare i MetaData riepilogativi</param>
    ''' <param name="resultData">Il set di risultati ottenuti dopo il processo di elaborazione sul proxy</param>
    ''' <returns>Se nel riepilogare i metadata ci accorgiamo che è stato superato per eccesso l'importo in pagamento restituiamo in un nuovo Media il resto eventuale da gestire</returns>
    Private Function SaveAndPutItemsOnMediaSingleBPE(ByRef RootTaMediaRec As TaMediaRec, ByRef resultData As ClsProxyArgentea.DataResponse) As TaMediaRec

        ' Durante il raggruppamento se
        ' un set di buoni ha  superato
        ' per eccesso il pagamento richiesto
        ' restituiamo in forma di TA nuova
        ' la voce della differenza.
        Dim PeExcedeedRec As TaMediaRec = Nothing
        Dim szValues As Dictionary(Of String, String) = New Dictionary(Of String, String)

        ' Totalizzatori per Taglio
        Dim _edges_DEMAT As String = ""
        Dim _edges_VOID As String = ""
        Dim _edges_INVALID As String = ""

        Dim KeyQTA As String
        Dim KeyTOT As String
        Dim KeyVALUE As String

        Dim _Num_Payed As Integer = 0
        Dim _Num_Voided As Integer = 0
        Dim _Num_Invalid As Integer = 0

        Dim _Tot_Payed As Decimal = 0
        Dim _Tot_Voided As Decimal = 0
        Dim _Tot_Invalid As Decimal = 0

        Dim _NumCurrT As Integer = 0
        Dim _ValCurrT As Decimal = 0
        Dim szValue As String = String.Empty
        Dim lIndex As Integer = 0

        szValues.Clear()

        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            '
            '   Accorpo nei metatag dell'unico  media
            '   in corso per ogni taglio il numero di
            '   tagli usati e l'importo totale. Cutoff
            '

            If Not pe.Voided And Not pe.Invalid Then

                ' Elemento
                Dim szItm As String = pe.DecimalValue.ToString().Replace(",", "_")
                ' Per gruppo
                If (Not szValues.ContainsKey(szItm)) Then
                    szValues.Add(szItm, szItm)
                    lIndex += 1
                Else
                    GoTo end_of_for
                End If

                ' Elemento Cntabilizzato
                KeyQTA = "lBP_QUANTITY_C_" + lIndex.ToString
                KeyTOT = "dBP_AMOUNT_C_" + lIndex.ToString
                KeyVALUE = "dBP_VALUE_C_" + lIndex.ToString

                For Each pet As PaidEntry In resultData.PaidEntryBindingSource

                    If Not pet.Voided And Not pet.Invalid And pet.DecimalValue = pe.DecimalValue Then

                        If Not RootTaMediaRec.ExistField(KeyQTA) Then

                            RootTaMediaRec.AddField(KeyQTA, DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                            RootTaMediaRec.AddField(KeyTOT, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            RootTaMediaRec.AddField(KeyVALUE, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            _NumCurrT = 0
                            _ValCurrT = 0
                            _edges_DEMAT += szItm + "|"   ' Ogni nuovo Taglio pe.FaceValue.Replace(",", "_").Trim() 

                        Else
                            _NumCurrT = RootTaMediaRec.GetPropertybyName(KeyQTA)
                            _ValCurrT = RootTaMediaRec.GetPropertybyName(KeyTOT) / m_ParseFractMode
                        End If
                        '
                        _Num_Payed += 1
                        _Tot_Payed += pet.DecimalValue
                        RootTaMediaRec.setPropertybyName(KeyQTA, _NumCurrT + 1)
                        RootTaMediaRec.setPropertybyName(KeyTOT, _ValCurrT + pet.DecimalValue)
                        RootTaMediaRec.setPropertybyName(KeyVALUE, pet.DecimalValue)

                    End If

                Next

            End If

end_of_for:

        Next

        _NumCurrT = 0
        _ValCurrT = 0
        szValue = String.Empty
        lIndex = 0

        szValues.Clear()

        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            '
            '   Accorpo nei metatag dell'unico  media
            '   in corso per ogni taglio il numero di
            '   tagli usati e l'importo totale per gli storni. Cutoff
            '

            If pe.Voided And Not pe.Invalid Then

                ' Elemento
                Dim szItm As String = pe.DecimalValue.ToString().Replace(",", "_")
                ' Per gruppo
                If (Not szValues.ContainsKey(szItm)) Then
                    szValues.Add(szItm, szItm)
                    lIndex += 1
                Else
                    GoTo end_of_for2
                End If

                'Elemento Stornato
                KeyQTA = "lBP_QUANTITY_V_" + lIndex.ToString
                KeyTOT = "dBP_AMOUNT_V_" + lIndex.ToString
                KeyVALUE = "dBP_VALUE_V_" + lIndex.ToString


                For Each pet As PaidEntry In resultData.PaidEntryBindingSource

                    If pet.Voided And Not pet.Invalid And pet.DecimalValue = pe.DecimalValue Then

                        If Not RootTaMediaRec.ExistField(KeyQTA) Then

                            RootTaMediaRec.AddField(KeyQTA, DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                            RootTaMediaRec.AddField(KeyTOT, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            RootTaMediaRec.AddField(KeyVALUE, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            _NumCurrT = 0
                            _ValCurrT = 0
                            _edges_VOID += szItm + "|"   ' Ogni nuovo Taglio pe.FaceValue.Replace(",", "_").Trim()
                        Else
                            _NumCurrT = RootTaMediaRec.GetPropertybyName(KeyQTA)
                            _ValCurrT = RootTaMediaRec.GetPropertybyName(KeyTOT) / m_ParseFractMode
                        End If
                        '
                        _Num_Voided += 1
                        _Tot_Voided += pet.DecimalValue
                        RootTaMediaRec.setPropertybyName(KeyQTA, _NumCurrT + 1)
                        RootTaMediaRec.setPropertybyName(KeyTOT, _ValCurrT + pet.DecimalValue)
                        RootTaMediaRec.setPropertybyName(KeyVALUE, pet.DecimalValue)

                    End If

                Next

            End If

end_of_for2:

        Next

        _NumCurrT = 0
        _ValCurrT = 0
        lIndex = 0

        szValues.Clear()

        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            '
            '   Accorpo nei metatag dell'unico  media
            '   in corso per ogni taglio il numero di
            '   tagli usati e l'importo totale non contabilizzato per sotrni non congrui. Cutoff
            '

            If pe.Invalid Then

                ' Elemento
                Dim szItm As String = pe.DecimalValue.ToString().Replace(",", "_")
                ' Per gruppo
                If (Not szValues.ContainsKey(szItm)) Then
                    szValues.Add(szItm, szItm)
                    lIndex += 1
                Else
                    GoTo end_of_for3
                End If

                'Elemento Stornato
                KeyQTA = "lBP_QUANTITY_E_" + lIndex.ToString
                KeyTOT = "dBP_AMOUNT_E_" + lIndex.ToString
                KeyVALUE = "dBP_VALUE_E_" + lIndex.ToString

                For Each pet As PaidEntry In resultData.PaidEntryBindingSource

                    If pet.Invalid And pet.DecimalValue = pe.DecimalValue Then

                        If Not RootTaMediaRec.ExistField(KeyQTA) Then

                            RootTaMediaRec.AddField(KeyQTA, DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                            RootTaMediaRec.AddField(KeyTOT, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            RootTaMediaRec.AddField(KeyVALUE, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            _NumCurrT = 0
                            _ValCurrT = 0
                            _edges_INVALID += szItm + "|"   ' Ogni nuovo Taglio '  pe.FaceValue.Replace(",", "_").Trim() 
                        Else
                            _NumCurrT = RootTaMediaRec.GetPropertybyName(KeyQTA)
                            _ValCurrT = RootTaMediaRec.GetPropertybyName(KeyTOT) / m_ParseFractMode
                        End If
                        '
                        _Num_Invalid += 1
                        _Tot_Invalid += pet.DecimalValue
                        RootTaMediaRec.setPropertybyName(KeyQTA, _NumCurrT + 1)
                        RootTaMediaRec.setPropertybyName(KeyTOT, _ValCurrT + pet.DecimalValue)
                        RootTaMediaRec.setPropertybyName(KeyVALUE, pet.DecimalValue)

                    End If

                Next

            End If
end_of_for3:
        Next

        ' Come meta riporto il restpo delle info per il raggruppamento
        RootTaMediaRec.AddField("ibp_GROUPED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_GROUPED", 2)

        Debug.OutputDebugString(
                CStr(_Num_Payed) + " / " + CStr(_Tot_Payed) + vbCrLf +
                CStr(_Num_Voided) + " / " + CStr(_Tot_Voided) + vbCrLf +
                CStr(_Num_Invalid) + " / " + CStr(_Tot_Invalid) + vbCrLf
        )

        ' BPE Dematerializzati

        RootTaMediaRec.AddField("szbp_EDGES_DEMAT", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
        RootTaMediaRec.setPropertybyName("szbp_EDGES_DEMAT", _edges_DEMAT)

        RootTaMediaRec.AddField("dbp_TOT_PAYED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_PAYED", _Tot_Payed - _Tot_Voided) ' resultData.totalPayedWithBP)          ' _Tot_Payed

        RootTaMediaRec.AddField("ibp_TOT_BP_USED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_USED", _Num_Payed) 'resultData.totalBPUsedToPay)        ' _Num_Payed 

        RootTaMediaRec.AddField("dbp_TOT_EXCEDEED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_EXCEDEED", resultData.totalExcedeedWithBP)

        ' BPE Stornati da demat precedente

        RootTaMediaRec.AddField("szbp_EDGES_VOID", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
        RootTaMediaRec.setPropertybyName("szbp_EDGES_VOID", _edges_VOID)

        RootTaMediaRec.AddField("dbp_TOT_VOIDED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_VOIDED", _Tot_Voided) 'resultData.totalVoidedWithBP)        ' _Tot_Voided

        RootTaMediaRec.AddField("ibp_TOT_BP_VOIDED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_VOIDED", _Num_Voided) 'resultData.totalBPUsedToVoid)     ' _Num_Voided

        ' BPE Errati non contabilizzati in void precedenti

        RootTaMediaRec.AddField("szbp_EDGES_INVALID", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
        RootTaMediaRec.setPropertybyName("szbp_EDGES_INVALID", _edges_INVALID)

        RootTaMediaRec.AddField("dbp_TOT_INVALID", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_INVALID", resultData.totalNotContabilizated)  ' _Tot_Invalid

        RootTaMediaRec.AddField("ibp_TOT_BP_INVALID", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_INVALID", resultData.totalBPNotValid)      ' _Num_Invalid

        ' Il Numero totale degli elementi in elenco MD
        RootTaMediaRec.AddField("ibp_TOT_BP_MD", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_MD", resultData.PaidEntryBindingSource.Count)

        '
        ' Se nella sessione c'è stato un eccesso
        ' rispetto al valore del buono quindi se
        ' opzionalmente c'è o meno da dare il resto
        ' (pe.Value <> pe.FaceValue) <-- Questo mi dice esatamente quale BP è in difetto rispetto al pagabile.
        '
        If resultData.totalExcedeedWithBP <> 0 Then 'AndAlso RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember Then

            Dim pe As New PaidEntry("Excdeed_BPE_GROUP", resultData.totalExcedeedWithBP)
            pe.InfoExtra = "Item BPD excdeed total for session"

            '
            ' Su un eccesso rispetto al  BP  corrente
            ' lo aggiungo alla transazione e con esso
            ' l'opzione che ha segnato di crearlo.
            ' E' sempre e solo una voce solitamente alla fine dei BP scansionati
            '
            PeExcedeedRec = FillMediaForExcedeed(RootTaMediaRec.PAYMENTinMedia.lMediaMember, pe)
            PeExcedeedRec.dTaPaid = -pe.DecimalValue
            PeExcedeedRec.dTaPaidTotal = -pe.DecimalValue
            PeExcedeedRec.theHdr.bIsVoided = True
            PeExcedeedRec.PAYMENTinMedia.szDesc = "Buono di Resto "

        End If

        Return PeExcedeedRec

    End Function

    ''' <summary>
    '''     Crea per una TA una serie di voci
    '''     MetaData con info di riepilogo su
    '''     un set di BP processati tramite il
    '''     proxy di Argentea senza suddivisione
    '''     ma l'intero elenco per ogni Barcode usato
    ''' </summary>
    ''' <param name="RootTaMediaRec">La TA di tipo Pagamento dove posizionare i MetaData riepilogativi</param>
    ''' <param name="resultData">Il set di risultati ottenuti dopo il processo di elaborazione sul proxy</param>
    ''' <returns>Se nel riepilogare i metadata ci accorgiamo che è stato superato per eccesso l'importo in pagamento restituiamo in un nuovo Media il resto eventuale da gestire</returns>
    Private Function SaveAndPutItemsOnMediaSingleBPC(ByRef RootTaMediaRec As TaMediaRec, ByRef resultData As ClsProxyArgentea.DataResponse) As TaMediaRec

        ' Durante il raggruppamento se
        ' un set di buoni ha  superato
        ' per eccesso il pagamento richiesto
        ' restituiamo in forma di TA nuova
        ' la voce della differenza.
        Dim PeExcedeedRec As TaMediaRec = Nothing
        Dim _NumCurrT As Integer = 0
        Dim _NumVoidedT As Integer = 0
        Dim _NumNotValidT As Integer = 0
        Dim _TotValPayed As Decimal = 0
        Dim _TotValVoided As Decimal = 0
        Dim _TotValNotValid As Decimal = 0
        Dim KeyCBP As String
        Dim KeyVBP As String
        Dim KeyIBP As String
        Dim KeyBP As String
        Dim OtherInfo As String

        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            '
            '   Elenco come metatag ogni singolo bp
            '   con relativo barcode usato nell'insieme
            '   di quelli elaborati per pagare.
            '
            KeyCBP = "bp_itm_" + CStr(_NumCurrT + 1)
            KeyVBP = "bp_itm_voided_" + CStr(_NumVoidedT + 1)
            KeyIBP = "bp_itm_invalid_" + CStr(_NumNotValidT + 1)

            ' Se è etichettato come Emitter = 'VOIDED'
            ' Vuol dire che siamo in una operazione di
            ' storno e quindi lo etichettiamo in  modo
            ' diverso per il presentation
            If pe.Voided And Not pe.Invalid Then ' --> VOIDED -> Stornato
                _NumVoidedT += 1
                KeyBP = KeyVBP
            ElseIf pe.Invalid Then               ' --> Invalid
                _NumNotValidT += 1
                KeyBP = KeyIBP
            Else                                 ' --> Valid USED
                _NumCurrT += 1
                KeyBP = KeyCBP
            End If

            ' Fillo le Meta Property

            ' itm BarCode
            RootTaMediaRec.AddField("sz" & KeyBP, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            RootTaMediaRec.setPropertybyName("sz" & KeyBP, pe.Barcode)

            ' itm Value
            RootTaMediaRec.AddField("d" & KeyBP + "_Value", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            RootTaMediaRec.setPropertybyName("d" & KeyBP + "_Value", pe.DecimalValue)

            ' itm FaceValue
            RootTaMediaRec.AddField("d" & KeyBP + "_FaceValue", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            RootTaMediaRec.setPropertybyName("d" & KeyBP + "_FaceValue", pe.DecimalFaceValue)

            ' itm IDCrcTransaction associata
            RootTaMediaRec.AddField("sz" & KeyBP + "_IDTransaction", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            RootTaMediaRec.setPropertybyName("sz" & KeyBP + "_IDTransaction", pe.IDTransactionCrc)

            ' itm Emitter e CodeIssuer
            OtherInfo = pe.Emitter.Replace("-", "_") & "-" & pe.CodeIssuer.Replace("-", "_") & "-" & pe.NameIssuer.Replace("-", "_")
            RootTaMediaRec.AddField("sz" & KeyBP + "_OtherInfo", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            RootTaMediaRec.setPropertybyName("sz" & KeyBP + "_OtherInfo", OtherInfo)

            If pe.Voided = True And Not pe.Invalid Then
                ' --> VOIDED -> Stornato
                _TotValVoided += pe.DecimalValue
            ElseIf Not pe.Invalid Then
                ' --> PAYED -> Pagato
                _TotValPayed += pe.DecimalValue
            Else ' --> INVALID -> Non Contabilizzato
                _TotValNotValid += pe.DecimalValue
            End If

        Next

        ' Come meta riporto il restpo delle info per il raggruppamento
        RootTaMediaRec.AddField("ibp_GROUPED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_GROUPED", 1)


        RootTaMediaRec.AddField("ibp_TOT_BP_NOTVALID", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_NOTVALID", _NumNotValidT)                      '  resultData.totalBPNotValid  

        RootTaMediaRec.AddField("dbp_TOT_NOTVALID", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_NOTVALID", _TotValNotValid)                       '  resultData.totalNoContabilizated   Eventuali elementi scartati non validi


        RootTaMediaRec.AddField("ibp_TOT_BP_VOIDED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_VOIDED", _NumVoidedT)                          '  resultData.totalBPUsedToVoid

        RootTaMediaRec.AddField("dbp_TOT_VOIDED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_VOIDED", _TotValVoided)                           '  resultData.totalVoidedWithBP


        RootTaMediaRec.AddField("ibp_TOT_BP_PAYED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_PAYED", _NumCurrT)                            '  resultData.totalBPUsedToPay

        RootTaMediaRec.AddField("dbp_TOT_PAYED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_PAYED", _TotValPayed)                             '  resultData.totalPayedWithBP      Che in Pagamento dovrebbe essere uguale a resultData.totalPayedWithBP

        RootTaMediaRec.AddField("dbp_TOT_EXCEDEED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_EXCEDEED", resultData.totalExcedeedWithBP)        '  resultData.totalBPUsedToPay


        RootTaMediaRec.AddField("ibp_TOT_BP_USED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_USED", _NumCurrT + _NumVoidedT)                ' resultData.totalBPUsedToPay + resultData.totalBPUsedToVoid

        '
        ' Se nella sessione c'è stato un eccesso
        ' rispetto al valore del buono quindi se
        ' opzionalmente c'è o meno da dare il resto
        ' (pe.Value <> pe.FaceValue) <-- Questo mi dice esatamente quale BP è in difetto rispetto al pagabile.
        '
        If resultData.totalExcedeedWithBP <> 0 AndAlso RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember Then

            Dim pe As New PaidEntry("Excdeed_BPE_GROUP", resultData.totalExcedeedWithBP)
            pe.InfoExtra = "Item BPD excdeed total for session"

            '
            ' Su un eccesso rispetto al  BP  corrente
            ' lo aggiungo alla transazione e con esso
            ' l'opzione che ha segnato di crearlo.
            ' E' sempre e solo una voce solitamente alla fine dei BP scansionati
            '
            PeExcedeedRec = FillMediaForExcedeed(RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember, pe)

        End If

        Return PeExcedeedRec

    End Function


    ''' <summary>
    '''     Crea una serie di MediaRec per
    '''     ogni BP che è  stato elaborato
    '''     tramite il proxy di Argentea.
    ''' </summary>
    ''' <param name="RootTa">La TA di root in corso dove accodare tutte le TA di pagmaneto per ogni singolo BP</param>
    ''' <param name="resultData">Il set di risultati ottenuti dopo il processo di elaborazione sul proxy</param>
    ''' <param name="IdSessionReferement">Ad ogni elmento sarà apposta l'etichetta di id di riferimento della sessione che ha creato la serie</param>
    ''' <param name="TotalsCountInSession">Il Numero di Titoli coinvolti nella sessione che si stanno impilando nella TA</param>
    ''' <returns>Se nel riepilogare i bp ci accorgiamo che è stato superato per eccesso l'importo in pagamento restituiamo in un nuovo Media il resto eventuale da gestire</returns>
    Private Function SaveAndPutItemsOnMediaGroup(ByRef RootTa As TA, ByRef resultData As ClsProxyArgentea.DataResponse, IdSessionReferement As String, TotalsCountInSession As Integer) As TaMediaRec
        Dim NewTaMediaRec As TaMediaRec

        ' Durante l'aggiunta  delle TA Media per
        ' ogni buono che ha partecipato a pagare
        ' si è superato in   qualche   modo  per 
        ' eccesso il pagamento richiesto restituiamo 
        ' in forma di TA nuova la voce della differenza.
        Dim PeExcedeedRec As TaMediaRec = Nothing
        Dim x As Integer = 0

        ' 
        ' Scorro per tutti i BP nell'elenco dei
        ' BP usati nella sessione proxy per aggiungere
        ' un MEDIA per ogni BP utilizzato.
        '
        '       NOTA.: Per ogni Buono viene aggiunto un MediaRecord.
        '
        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            ' Per la sessione scrivo in Info Extra il Numero di Per es.: 1/3 che sta a a 1 di 3 poi 2 di tre etc.
            pe.InfoExtra = CStr(x + 1) & "/" & TotalsCountInSession
            x = x + 1

            ' Aggiungo sulla Transazione corrente la TA relativa al Media di pagamento "Non Raggruppato"
            NewTaMediaRec = AddNewTaMedia(pe, resultData.typeBPElaborated.ToString(), False)

            ' Identifico che questa voce è singola e non raggruppante
            'NewTaMediaRec.AddField("ibp_GROUPED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            'NewTaMediaRec.setPropertybyName("ibp_GROUPED", 0)

            '
            ' Definisco per questo elemento i suoi Metadata
            ' che servono per l'eventuale annullo e storno.
            '
            Dim KeyCBP As String = "bp_grp_itm"

            ' Fillo le Meta Property

            ' itm BarCode
            NewTaMediaRec.AddField("sz" & KeyCBP, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewTaMediaRec.setPropertybyName("sz" & KeyCBP, pe.Barcode)

            ' itm Value
            NewTaMediaRec.AddField("d" & KeyCBP + "_Value", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            NewTaMediaRec.setPropertybyName("d" & KeyCBP + "_Value", pe.DecimalValue)

            ' itm FaceValue
            NewTaMediaRec.AddField("d" & KeyCBP + "_FaceValue", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            NewTaMediaRec.setPropertybyName("d" & KeyCBP + "_FaceValue", pe.DecimalFaceValue)

            ' itm IDCrcTransaction associata (Da Argentea)   '''' 'szbp_grp_itm_IDTransaction'
            NewTaMediaRec.AddField("sz" & KeyCBP + "_IDTransaction", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewTaMediaRec.setPropertybyName("sz" & KeyCBP + "_IDTransaction", pe.IDTransactionCrc)

            ' itm IdGroupReferement associata (Da Sessione)
            NewTaMediaRec.AddField("sz" & KeyCBP + "_IDSessionRef", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewTaMediaRec.setPropertybyName("sz" & KeyCBP + "_IDSessionRef", IdSessionReferement)

            ' itm Emitter e CodeIssuer
            Dim OtherInfo As String = pe.Emitter.Replace("-", "_") & "-" & pe.CodeIssuer.Replace("-", "_") & "-" & pe.NameIssuer.Replace("-", "_")
            NewTaMediaRec.AddField("sz" & KeyCBP + "_OtherInfo", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewTaMediaRec.setPropertybyName("sz" & KeyCBP + "_OtherInfo", OtherInfo)

            '
            ' Aggiungo alla transazione l'elenco
            ' del media record di tipo BP appena 
            ' creato con i dati riportati.
            '
            RootTa.Add(NewTaMediaRec)

            '
            ' Se nella sessione c'è stato un eccesso
            ' rispetto al valore del buono quindi se
            ' opzionalmente c'è o meno da dare il resto
            ' (pe.Value <> pe.FaceValue) <-- Questo mi dice esatamente quale BP è in difetto rispetto al pagabile.
            '
            If resultData.totalExcedeedWithBP <> 0 And (pe.Value <> pe.FaceValue) AndAlso NewTaMediaRec.PAYMENTinMedia.lChangeMediaMember Then

                '
                ' Su un eccesso rispetto al  BP  corrente
                ' lo aggiungo alla transazione e con esso
                ' l'opzione che ha segnato di crearlo.
                ' E' sempre e solo una voce solitamente alla fine dei BP scansionati
                '
                PeExcedeedRec = FillMediaForExcedeed(NewTaMediaRec.PAYMENTinMedia.lChangeMediaMember, pe)

            End If

        Next

        Return PeExcedeedRec

    End Function


    ''' <summary>
    '''     Aggiunge un MediaRecord usato come valore
    '''     di eccesso su una transazione in corso.
    ''' </summary>
    ''' <param name="ChangeMediaMember">L'indice della opzione per usare nella voce di Media sul perchè gestire l'eccedenza</param>
    ''' <param name="PeOnExcedeed">Il Buono coinvolto che ha fatto un eccedenza rispetto al rimanente</param>
    Private Function FillMediaForExcedeed(ChangeMediaMember As Integer, PeOnExcedeed As PaidEntry) As TaMediaRec
        Dim SelectedMedia As clsSelectMedia = Nothing
        Dim NewExcedeed As TaMediaRec = Nothing
        Dim funcname As String = "FillMediaForExcedeed"

        '
        ' Se l'opzione per gestire il troncamento sul resto
        ' rispetto al pagato in buoni è presente oppure no.
        '

        LOG_Info(getLocationString("HandleControllerArgentea"), "Manage exceed for voucher " & PeOnExcedeed.Barcode)

        Try

            '
            ' Riprendo il Media per riferimento alla transazione
            ' attribuendo dei parametri che lo  identificano
            ' come eccesso rispetto al pagabile per gestire
            ' in seguito situazioni di resto al cliente.
            '
            SelectedMedia = createPosModelObject(Of clsSelectMedia)(m_TheModcntr, "clsSelectMedia", 0, True)

            '
            ' Il Media TA in Eccedenza sul Pagabile.
            '
            NewExcedeed = m_taobj.CreateTaObject(Of TaMediaRec)(PosDef.TARecTypes.iTA_MEDIA)
            NewExcedeed.theHdr.lTaRefToCreateNmbr = 0
            NewExcedeed.theHdr.lTaCreateNmbr = 0
            NewExcedeed.dTaQty = 1
            NewExcedeed.dReturn = (Convert.ToDecimal(PeOnExcedeed.FaceValue) - Convert.ToDecimal(PeOnExcedeed.Value))
            NewExcedeed.AddField("szBuono" + TYPE_SPECIFIQUE, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewExcedeed.setPropertybyName("szBuono" + TYPE_SPECIFIQUE, PeOnExcedeed.Barcode)
            NewExcedeed.AddField("szFaceValue", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewExcedeed.setPropertybyName("szFaceValue", PeOnExcedeed.FaceValue)

            ' Utilizzo il metodo apposito per fillare
            ' tutti gli altri attributi necessari al media.
            If SelectedMedia.FillPaymentDataFromID(
                                    m_TheModcntr,
                                    NewExcedeed.PAYMENTinMedia,
                                    ChangeMediaMember,
                                    m_taobj,
                                    m_taobj.colObjects
                    ) Then

                ' MEDIA di resto creato ed accodato
                LOG_Info(getLocationString("HandleControllerArgentea"), "Exceed managed with success for voucher " & PeOnExcedeed.Barcode)

            Else

                '
                ' Se non si  riesce  a  creare la 
                ' voce di riferimento all'eccesso
                ' in ogni caso la transazione  si
                ' conclude.
                '
                NewExcedeed = Nothing

                '
                ' Solleviamo l'eccezione relativa al media
                ' per l'eccesso su totale non riuscito.
                '
                Throw New Exception(INT_ERROR_NOT_CREATE_EXCEDEED)

            End If

        Catch Ex As Exception

            '
            ' Se fallisce la creazione  della 
            ' voce di riferimento all'eccesso
            ' in ogni caso la transazione  si
            ' conclude.
            '
            NewExcedeed = Nothing

            ' Signal (Errore nel Creare il Media come Media di Notazione di Resto su BP Pagati in eccesso rispetto al Totale)
            _SetOperationStatus(funcname,
                GLB_ERROR_ON_CREATE_EXCEDEED,
                "Errore interno non gestito - Exception on create Media Return Note excedeed -- (Chiamare assistenza)",
                PosDef.TARMessageTypes.TPSTOP, True
            )

            ' Log locale
            LOG_Error(getLocationString(funcname), m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), Ex)

        End Try

        Return NewExcedeed

    End Function

#End Region

End Class


''' <summary>
'''     BUONI PASTO CARTACEI
''' </summary>
Public Class BPCController
    Inherits BPControllerBase
    Implements IBPDematerialize

    ' * ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   SPECIFICITA' del CONTROLLER
    '
    ' * ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' VARIANTE PER CARTACEI

    ''' <summary>
    '''     Controller di tipo Buoni Pasto Cartacei
    ''' </summary>
    Public Sub New()
        MyBase.New("BPC", ClsProxyArgentea.enTypeProxy.Service)
    End Sub

End Class

''' <summary>
'''     BUONI PASTO ELETTRONICI
''' </summary>
Public Class BPEController
    Inherits BPControllerBase
    Implements IBPDematerialize

    ' * ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   SPECIFICITA' del CONTROLLER
    '
    ' * ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' VARIANTE PER ELETTRONICI

    ''' <summary>
    '''     Controller di tipo Buoni Pasto Elettronici
    ''' </summary>
    Public Sub New()
        MyBase.New("BPE", ClsProxyArgentea.enTypeProxy.Pos)
    End Sub

End Class
