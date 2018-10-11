Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports TPDotnet.IT.Common.Pos.EFT
Imports TPDotnet.IT.Common.Pos

Public Class BPCController
    Implements IBPDematerialize

    ' VARIANTE PER ELETTRONICI O CARTACEI

    ' Identifica nei metadata la Key per tipo
    Private Const TYPE_SPECIFIQUE As String = "BPC"                                                         ' <-- Costante per identificare il sottotipo nell'applicazione (Cartacei)

    ' Modalità del servizio Proxy di Argentea
    Private Const MODE_PROXY_USE As ClsProxyArgentea.enTypeProxy = ClsProxyArgentea.enTypeProxy.Service      ' <-- Il Proxy servizio avviato in modalità Emulazione Software


#Region "CONST di INFO e ERRORE Private"

    ' Messaggeria per codifica segnalazioni ID di errore remoti 
    Private msgUtil As New TPDotnet.IT.Common.Pos.Common

    ' Su Errore Pagabile rispetto alla vendita corrente, non sono 
    ' stati trovati prodotti nella vendita con applicato il filtro
    ' xslt e quindi vendita a 0 usciamo dal flow
    Private Const GLB_ERROR_FILTER As String = "ERROR_FILTER_PAYABLES_BP"

    ' Su Errore di configurazione o di sistema quando la 
    ' procedura non trova il file xslt di trasformazione 
    ' usato per riprendere i totali dei prodotti che  si
    ' possono pagare tramite BP
    Private Const GLB_ERROR_FILE_FILTER As String = "ERROR_FILE_FILTER_PAYABLES_BP"

    ' Per la corretta istanza di questo modello applicativo
    ' sono sempre necessari nei parametri caricati dinamicamente
    ' op er istanza di questa classe sia il controller ModCntr
    ' da cui viene avviato e la transazione in corso.
    Private Const GLB_ERROR_INTERNAL_START As String = "ERRPR_START_PROXY_INTERNAL"

    ' Su Errore Parsing utilizzata per segnalazione
    ' di errore su protocollo non previsto.
    '   Errore classificato su risposta Argentea .: RetVal.CodeResult 
    Private Const GLB_ERROR_PARSING As String = "ERROR-PARSING"

    ' Su Errore Pagabile rispetto all'ammontare del pagamento
    ' assegno errore di non valido per uscire dal pagamento BP
    Private Const GLB_ERROR_PAYABLE As String = "ERROR_PAYABLE"

    ' Su Errore Pagabile rispetto all'ammontare già totalmente pagato
    ' assegno errore di non valido per uscire dal pagamento BP
    Private Const GLB_ERROR_PAYABLE2 As String = "ERROR_PAYABLE2"

    ' Nel Flow della nel momento in cui si trova a creare un
    ' MEDIA di resto se c'è un eccezione interna enon viene creato.
    Private Const GLB_ERROR_ON_CREATE_EXCEDEED As String = "Error-ERROR-CREATE-EXCEDEED"

    ' Nel Flow della nel momento in cui si trova a creare un
    ' MEDIA di resto se c'è un errore di valutazione interna enon viene creato.
    Private Const GLB_ERROR_NOT_CREATE_EXCEDEED As String = "Error-EXCEPTION-CREATE-EXCEDEED"

    ' Nell'istanziare la classe servizio verso Argentea (proxy mode)
    ' il controller di gestione corrente ha sollevato un eccezione.
    Private Const GLB_ERROR_INSTANCE_SERVICE As String = "EXCEPTION-PROXY-INSTANCE"

    ' Nel riprendere i dati provenienti dal proxy di Argentea (proxy mode)
    ' il controller di gestione corrente ha sollevato un eccezione.
    Private Const GLB_ERROR_COLLECT_DATA_SERVICE As String = "EXCEPTION-COLLECT-PROCY-DATA"

    ' Nel Flow della funzione Entry il Throw non
    ' previsto.
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

#End Region

#Region ".ctor"

    ''' <summary>
    '''     .ctor
    ''' </summary>
    ''' <param name="theModCntr">controller -> Il Controller per riferimento dal chiamante</param>
    ''' <param name="taobj">transaction -> La TA per riferimento dal chiamante</param>
    'Public Sub New(ByRef theModCntr As ModCntr, ByRef taobj As TA)
    'Public Sub New()

    'm_taobj = Nothing
    'm_TheModcntr = Nothing

    'End Sub

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

                ' Etichettiamo l'errore per la gestione
                m_LastStatus = GLB_ERROR_FILE_FILTER
                m_LastErrorMessage = "File di trasformazione per la vendita con Buoni Pasto non valido o non presente"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            End If

            ' Se invece siamo con un totale a 0 da pagare perchè tutti i prodotti
            ' sono esclusi dalla vendita possibile tramite i Buoni Pasto.
            Dim TotalTransaction As Decimal = CDec(sTotalTransaction) / 100

            If TotalTransaction = 0 Then

                ' Valorizzo l'errore nei Params per uscire dall
                ' iterazione del modulo corrente
                m_LastStatus = GLB_ERROR_FILTER
                m_LastErrorMessage = "Non sono presenti prodotti in questa vendita che è possibile pagare con i Buoni Pasto"

                ' Messagio per l'utente
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

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

                    ' Valorizzo l'errore nei Params per uscire dall
                    ' iterazione del modulo corrente
                    m_LastStatus = GLB_ERROR_PAYABLE
                    m_LastErrorMessage = "Il pagamento eccede sul totale"

                    ' Messagio per l'utente
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                    ' Log locale
                    LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

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
                End If
                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

            Else

                ' Valorizzo l'errore nei Params per uscire dall
                ' iterazione del modulo corrente
                m_LastStatus = GLB_ERROR_PAYABLE2
                m_LastErrorMessage = "Il pagamento ricopre già l'intero totale"

                ' Messaggio per l'utente
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                Dematerialize = IBPReturnCode.KO
                Exit Function

            End If

        Catch ex As Exception

            ' Signal come errore non previsto anulla l'operazione di pagamento
            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
            m_LastErrorMessage = "Errore interno non previsto --exception on call void in controller bp with argentea-- (Chiamare assistenza)"

            ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

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
            If m_CurrMedia Is Nothing _
                OrElse (m_CurrMedia.dTaPaid >= 0 Or m_CurrMedia.PAYMENTinMedia Is Nothing) _
                OrElse (m_CurrMedia.PAYMENTinMedia.szExternalID <> TYPE_SPECIFIQUE) Then

                ' Etichettiamo l'errore per la gestione
                m_LastStatus = GLB_ERROR_FILE_FILTER
                m_LastErrorMessage = "Funzione di storno non richiamabile per chiamata a storno non congruente"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Void Argentea ::KO:: Local")

                Return IBPReturnCode.KO
            Else

                ' Importo di storno
                sTotalVoid = m_CurrMedia.dTaPaid

            End If

            ' Se invece siamo con uno storno a >= 0 da fare perchè tutti 
            ' l'erore è nella procedura di configuraione.
            Dim TotalToVoid As Decimal = CDec(sTotalVoid) '/ 100

            If TotalToVoid >= 0 Then

                ' Valorizzo l'errore nei Params per uscire dall
                ' iterazione del modulo corrente
                m_LastStatus = GLB_ERROR_FILTER
                m_LastErrorMessage = "Errore di forma nello storno"

                ' Messaggio per l'utente
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Void Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            Else

                ' Questo è quello che si può pagare in totale.
                m_VoidAmount = +(TotalToVoid)

                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **
                '
                ' Richiamiamo il Metodo per l'Azione per visualizzarlo
                ' passando l'intera transazione il Controller corrente
                ' Il Metodo di pagamento passato come argomento
                '
                If HandleVoidBPCall(m_VoidAmount) Then
                    Void = IBPReturnCode.OK
                End If
                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

            End If

        Catch ex As Exception

            ' Signal come errore non previsto non permette di continuare con lo storno
            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
            m_LastErrorMessage = "Errore interno non previsto --exception on call void in controller bp with argentea-- (Chiamare assistenza)"

            ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

        Finally
            LOG_FuncExit(getLocationString(funcName), Void.ToString())
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

            ' Signal come errore non previsto non permette di continuare con lo storno
            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
            m_LastErrorMessage = "Errore interno non previsto --exception on call void in controller bp with argentea-- (Chiamare assistenza)"

            If Not SilentMode Then

                ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            End If

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

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
    Private Function _initializeParameters(ByRef funcName As String, ByRef parameters As Dictionary(Of String, Object), showOperatorError As Boolean) As IBPReturnCode

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

                ' Etichettiamo l'errore per la gestione
                m_LastStatus = GLB_ERROR_INTERNAL_START
                m_LastErrorMessage = "Per la corretta esecuzione del proxy sono necessari il controller e la transazione in corso"

                If showOperatorError Then

                    ' Se per qualche motivo o perchè manca il file di trasformazione
                    ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                    msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

                End If

                ' Log locale
                LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage + " - " + "Transaction Dematerialize Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            End If

        Catch ex As Exception

            ' Signal come errore non previsto anulla l'operazione di pagamento
            m_LastStatus = GLB_ERROR_NOT_UNEXPECTED
            m_LastErrorMessage = "Errore interno non previsto --exception on initialize parameter in controller bp with argentea-- (Chiamare assistenza)"

            If showOperatorError Then

                ' message box: atenzione non sono riuscito a stampare la ricevuta ma la transazione è valida
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)
            End If

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

        Finally
            LOG_FuncExit(getLocationString(funcName), _initializeParameters.ToString())
        End Try


    End Function


#End Region

#Region "Handler Internal Form Action"

    ''' <summary>
    '''     Handle per azionare il pagamento tramite Form
    '''     POS simulato in questa cassa.
    ''' </summary>
    ''' <returns>Boolean True (Conslusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandlePaymentBPCall(ByVal dPayableAmount As Decimal) As Boolean
        Dim funcName As String = "HandlePaymentBPCall"
        Dim proxyPos As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandlePaymentBPCall = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al service di Argentea.
            '
            If proxyPos Is Nothing Then
                proxyPos = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                pParams.TransactionID,                  '   <-- L'id della transazione in corso
                m_CurrMedia.dTaPaidTotal                '   <-- Il Pagato fino adesso insieme agli altri media
            )

                '
                ' Preparo ad accettare l'handler degli eventi gestiti
                ' prima e dopo la comunicazione con il POS locale.
                '
                AddHandler proxyPos.Event_ProxyCollectDataTotalsAtEnd, AddressOf ProxyCollectDataTotalsAtEnd_Handler

            End If

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

            ' Etichettiamo l'errore per la gestione
            m_LastStatus = GLB_ERROR_INSTANCE_SERVICE
            m_LastErrorMessage = "Eccezione non gestita nell'istanza della classe servizio di Argentea"

            ' Se per qualche motivo o perchè manca il file di trasformazione
            ' o per errori in esecuzione non applica il filtro esco dalla gestione.
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPSTOP)

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
        proxyPos.Command = ClsProxyArgentea.enCommandToCall.Payment
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

        Else

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
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al proxy di Argentea.
            '
            If proxyPos Is Nothing Then
                proxyPos = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                pParams.TransactionID,                  '   <-- L'id della transazione in corso
                m_CurrMedia.dTaPaidTotal                '   <-- Il Pagato fino adesso insieme agli altri media
            )

                '
                ' Preparo ad accettare l'handler degli eventi gestiti
                ' prima e dopo la comunicazione con il POS locale.
                '
                AddHandler proxyPos.Event_ProxyCollectDataVoidedAtEnd, AddressOf ProxyCollectDataVoidedAtEnd_Handler

            End If


            '
            ' Preparo l'oggetto a quello che si deve a  spettare
            ' come totale da stornare e quello stornato
            '
            proxyPos.AmountVoid = 0
            proxyPos.AmountVoidable = dVoidableAmount

            '
            ' Dato che è uno storno vediamo se sul POS
            ' di servizio vogliamo ripresentare  i  BP
            ' se questi sono in un  raggruppamento  da 
            ' usare su un FORM prefillato con l'elenco.
            '
            If Not m_CurrMedia Is Nothing Then
                proxyPos.PrefillVoidable = CheckAndFillListOfGroupItems(proxyPos.FractParsing)
            End If

            ' Definisco questa  variabile  Privata 
            ' per il conteggio dei Buoni eventuali
            ' già presenti nella TA che sono stati
            ' usati in precedenza dall'operatore.
            _InitialBPPayed = pParams.GetAlreadyOnTAScanned()

        Catch ex As Exception

            LOG_Debug(getLocationString(funcName), "Instance proxyPos Argentea ::KO:: Local")

            ' Etichettiamo l'errore per la gestione
            m_LastStatus = GLB_ERROR_INSTANCE_SERVICE
            m_LastErrorMessage = "Eccezione non gestita nell'istanza della classe servizio di Argentea"

            ' Se per qualche motivo o perchè manca il file di trasformazione
            ' o per errori in esecuzione non applica il filtro esco dalla gestione.
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPSTOP)

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
            proxyPos.Command = ClsProxyArgentea.enCommandToCall.Void
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
        Dim funcName As String = "HandleVoidBPCall"
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
            If proxyPos Is Nothing Then
                proxyPos = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                MODE_PROXY_USE,                         '   <-- Il Proxy servizio avviato in modalità
                pParams.TransactionID,                  '   <-- L'id della transazione in corso
                m_CurrMedia.dTaPaidTotal                '   <-- Il Pagato fino adesso insieme agli altri media
            )

                proxyPos.SilentMode = SilentMode

            End If

        Catch ex As Exception

            LOG_Debug(getLocationString(funcName), "Instance proxyPos Argentea ::KO:: Local")

            ' Etichettiamo l'errore per la gestione
            m_LastStatus = GLB_ERROR_INSTANCE_SERVICE
            m_LastErrorMessage = "Eccezione non gestita nell'istanza della classe servizio di Argentea"

            If Not SilentMode Then

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPSTOP)

            End If

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
    Function CheckAndFillListOfGroupItems(FractToValues As Integer) As Dictionary(Of String, PaidEntry)
        Dim _NumCurrT As Integer = 1
        Dim KeyCBP As String
        Dim KeyVBP As String
        Dim _ARVoided As PaidEntry

        CheckAndFillListOfGroupItems = Nothing

        '
        If m_CurrMedia.GetPropertybyName("ibp_GROUPED") = 1 Then

            '
            '   Di tipo 1 Sono i BP cartacei solitamente
            '   listati e non raggruppati nei Metatag
            '

            Dim itmBarCode As String
            Dim itmFaceVal As Decimal

            CheckAndFillListOfGroupItems = New Dictionary(Of String, PaidEntry)

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
                itmFaceVal = m_CurrMedia.GetPropertybyName("d" & KeyCBP + "_Value") / FractToValues

                CheckAndFillListOfGroupItems.Add(itmBarCode, New PaidEntry(itmBarCode, itmFaceVal))
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
                    itmFaceVal = m_CurrMedia.GetPropertybyName("d" & KeyVBP + "_Value") / FractToValues

                    _ARVoided = New PaidEntry(itmBarCode, itmFaceVal)
                    _ARVoided.Emitter = "[[VOIDED]]"

                    CheckAndFillListOfGroupItems.Add(itmBarCode, _ARVoided)
                    _NumCurrT += 1

                Next

            End If


        ElseIf m_CurrMedia.GetPropertybyName("ibp_GROUPED") = 2 Then

            '
            '   Di tipo 2 Sono i BP elettronici solitamente
            '   raggruppati per taglio e numero di pezzi nei Metatag
            '

            CheckAndFillListOfGroupItems = New Dictionary(Of String, PaidEntry)

            Dim TypeEdges As String() = m_CurrMedia.GetPropertybyName("szbp_EDGES").Split("|")

            For x As Integer = 0 To UBound(TypeEdges) - 1

                Dim KeyQTA As String = "ibp_QUANTITY_" + TypeEdges(x)
                Dim KeyTOT As String = "dbp_AMOUNT_" + TypeEdges(x)

                Dim ValQTA As Integer = m_CurrMedia.GetPropertybyName(KeyQTA)
                Dim ValTOT As Decimal = m_CurrMedia.GetPropertybyName(KeyTOT) / FractToValues

                CheckAndFillListOfGroupItems.Add(CStr(ValQTA), New PaidEntry(CStr(ValQTA), ValTOT))

            Next


        End If
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

        ' Se non ci sono stati movimenti 
        ' ritorno al chiamante non mi serve
        ' nulla.
        If resultData.totalBPUsed = 0 Then

            ' Ripristino il pagamento esterno
            ' nella TA di origine ed esco dato
            ' che non ho nulla.
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
            OptAccorpateMediaForBP = True

            If OptAccorpateMediaForBP Then

                '
                ' Aggiungo una sola TA con un set di 
                ' BP riepilogati per Taglio che sono
                ' stati elaborati nel processo sul proxy
                '
                '       NOTA.: Viene creato Un solo MediaRecord con un riepilogo dei BP utilizzati.
                '
                Dim NmRnd As Integer = Rnd(999999999999999)
                Dim ItmPe As PaidEntry = New PaidEntry("Grouped_BP_" + CStr(NmRnd), resultData.totalPayedWithBP)

                ' Aggiungo sulla Transazione corrente
                ' la TA relativa al Media di pagamento.
                NewTaMediaRec = AddNewTaMedia(ItmPe)

                '
                ' Aggiungo il Riepilogo alla TA appena creata
                ' e controllo all'uscita se mi ha prodotto un
                ' resto da accodare come Media al pagamento
                '
                If resultData.typeBPElaborated = ClsProxyArgentea.enTypeBP.TicketsCard Then

                    ' Per i Buoni Pasti Elettronici riepilogo
                    ' per Taglio e valuta il numero dei conumati
                    PeExcedeedRec = AddGroupMetaDataForBpCutoffElaborated(NewTaMediaRec, resultData)

                ElseIf resultData.typeBPElaborated = ClsProxyArgentea.enTypeBP.TicketsRestaurant Then

                    ' Peri Buoni Pasti Cartacei riepilogo 
                    ' per ogni singolo Buono con riportato
                    ' il suo relativo BarCode utilizzato.
                    PeExcedeedRec = AddGroupMetaDataForBpBarcodeElaborated(NewTaMediaRec, resultData)

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
                PeExcedeedRec = AddMediaRecsForBpElaborated(m_taobj, resultData)

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
            m_LastErrorMessage = "Eccezione non gestita nel restituire i dati al chiamate dal set di dati Proxy"

            ' Log locale
            LOG_Error(funcName, m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), ex)

            ' Forzo il Proxy allo stato di errore in corso
            CType(sender, ClsProxyArgentea).SetStatusInError(
                funcName,
                m_LastStatus,
                m_LastErrorMessage,
                True
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

        '
        '   Richiamo lo stesso Handler di evento di aggiunta
        ' dato che mi dovrebbe eliminare l'originale, faccio
        ' in modo che se sto venendo da void l'originale lo lascio.
        '
        ProxyCollectDataTotalsAtEnd_Handler(sender, resultData)

    End Sub

#End Region

#Region "Functions Common"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

#Region "Functions Private"

    ''' <summary>
    '''     Crea un tipo MediaRec compatibile
    '''     per essere inserito nella Transazione
    '''     in corso come riga di pagamento.
    ''' </summary>
    ''' <param name="PeTo">Il new entry per completare la riga <see cref="PaidEntry"/></param>
    ''' <returns>Una Media compatibile con la transazione <see cref="TaMediaRec"/></returns>
    Private Function AddNewTaMedia(PeTo As PaidEntry) As TaMediaRec

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

            ' Nell'ordine di un elemento in quantità sempre 1 per ogni BP
            ' aggiungo la Proprietà di tipo Stringa .:   BarCode relativo al buono
            .dTaQty = 1
            .AddField("sz" + TYPE_SPECIFIQUE, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("sz" + TYPE_SPECIFIQUE, PeTo.Barcode)

            ' Aggiungo il valore di facciata che mi ha restituito Argentea
            ' ripresa dalla gridlia del form dove ho annotato.
            .AddField("szFaceValue", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("szFaceValue", PeTo.FaceValue)

            'CP#1337781:1:  default there's not rest, the mv was truncated
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
    Private Function AddGroupMetaDataForBpCutoffElaborated(ByRef RootTaMediaRec As TaMediaRec, ByRef resultData As ClsProxyArgentea.DataResponse) As TaMediaRec

        ' Durante il raggruppamento se
        ' un set di buoni ha  superato
        ' per eccesso il pagamento richiesto
        ' restituiamo in forma di TA nuova
        ' la voce della differenza.
        Dim PeExcedeedRec As TaMediaRec = Nothing

        ' Totalizzatori per Taglio
        Dim _NumCurrT As Integer = 0
        Dim _ValCurrT As Decimal = 0
        Dim szValue As String = String.Empty
        Dim lIndex As Integer = 0
        Dim _edges As String = ""

        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            '
            '   Accorpo nei metatag dell'unico  media
            '   in corso per ogni taglio il numero di
            '   tagli usati e l'importo totale. Cutoff
            '
            If (szValue <> pe.FaceValue.Replace(",", "_")) Then
                szValue = pe.FaceValue.Replace(",", "_")
                lIndex += 1
            End If
            Dim KeyQTA As String = "lBP_QUANTITY_" + lIndex.ToString
            Dim KeyTOT As String = "dBP_AMOUNT_" + lIndex.ToString
            Dim KeyVALUE As String = "dBP_VALUE_" + lIndex.ToString

            If Not RootTaMediaRec.ExistField(KeyQTA) Then
                RootTaMediaRec.AddField(KeyQTA, DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                RootTaMediaRec.AddField(KeyTOT, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                RootTaMediaRec.AddField(KeyVALUE, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                _NumCurrT = 0
                _ValCurrT = 0
                _edges += pe.FaceValue.Replace(",", "_").Trim() + "|"   ' Ogni nuovo Taglio
            Else
                _NumCurrT = RootTaMediaRec.GetPropertybyName(KeyQTA)
                _ValCurrT = RootTaMediaRec.GetPropertybyName(KeyTOT) / 100
            End If
            '
            RootTaMediaRec.setPropertybyName(KeyQTA, _NumCurrT + 1)
            RootTaMediaRec.setPropertybyName(KeyTOT, _ValCurrT + CDec(pe.Value))
            RootTaMediaRec.setPropertybyName(KeyVALUE, CDec(pe.Value))

            '
            ' Se nella sessione c'è stato un eccesso
            ' rispetto al valore del buono quindi se
            ' opzionalmente c'è o meno da dare il resto
            ' (pe.Value <> pe.FaceValue) <-- Questo mi dice esatamente quale BP è in difetto rispetto al pagabile.
            '
            If resultData.totalExcedeedWithBP <> 0 And (pe.Value <> pe.FaceValue) AndAlso RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember Then

                '
                ' Su un eccesso rispetto al  BP  corrente
                ' lo aggiungo alla transazione e con esso
                ' l'opzione che ha segnato di crearlo.
                ' E' sempre e solo una voce solitamente alla fine dei BP scansionati
                '
                PeExcedeedRec = FillMediaForExcedeed(RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember, pe)

            End If

        Next

        ' Come meta riporto il restpo delle info per il raggruppamento
        RootTaMediaRec.AddField("ibp_GROUPED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_GROUPED", 2)

        RootTaMediaRec.AddField("szbp_EDGES", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
        RootTaMediaRec.setPropertybyName("szbp_EDGES", _edges)

        RootTaMediaRec.AddField("dbp_TOT_PAYED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_PAYED", CDec(resultData.totalPayedWithBP))

        RootTaMediaRec.AddField("dbp_TOT_EXCEDEED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_EXCEDEED", CDec(resultData.totalExcedeedWithBP))

        RootTaMediaRec.AddField("ibp_TOT_BP_USED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_USED", CDec(resultData.totalBPUsed))

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
    Private Function AddGroupMetaDataForBpBarcodeElaborated(ByRef RootTaMediaRec As TaMediaRec, ByRef resultData As ClsProxyArgentea.DataResponse) As TaMediaRec

        ' Durante il raggruppamento se
        ' un set di buoni ha  superato
        ' per eccesso il pagamento richiesto
        ' restituiamo in forma di TA nuova
        ' la voce della differenza.
        Dim PeExcedeedRec As TaMediaRec = Nothing
        Dim _NumCurrT As Integer = 0
        Dim _NumVoidedT As Integer = 0
        Dim _TotValVoided As Decimal = 0
        Dim KeyCBP As String
        Dim KeyVBP As String

        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            '
            '   Elenco come metatag ogni singolo bp
            '   con relativo barcode usato nell'insieme
            '   di quelli elaborati per pagare.
            '
            KeyCBP = "bp_itm_" + CStr(_NumCurrT + 1)
            KeyVBP = "bp_itm_voided_" + CStr(_NumVoidedT + 1)

            ' Se è etichettato come Emitter = 'VOIDED'
            ' Vuol dire che siamo in una operazione di
            ' storno e quindi lo etichettiamo in  modo
            ' diverso per il presentation
            If pe.Emitter = "[[VOIDED]]" Then

                _NumVoidedT += 1

                ' itm BarCode
                RootTaMediaRec.AddField("sz" & KeyVBP, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                RootTaMediaRec.setPropertybyName("sz" & KeyVBP, pe.Barcode)

                ' itm Value
                RootTaMediaRec.AddField("d" & KeyVBP + "_Value", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                RootTaMediaRec.setPropertybyName("d" & KeyVBP + "_Value", CDec(pe.Value))

                _TotValVoided += CDec(pe.Value)

            Else

                _NumCurrT += 1

                ' itm BarCode
                RootTaMediaRec.AddField("sz" & KeyCBP, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                RootTaMediaRec.setPropertybyName("sz" & KeyCBP, pe.Barcode)

                ' itm Value
                RootTaMediaRec.AddField("d" & KeyCBP + "_Value", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                RootTaMediaRec.setPropertybyName("d" & KeyCBP + "_Value", CDec(pe.Value))


            End If

            '
            ' Se nella sessione c'è stato un eccesso
            ' rispetto al valore del buono quindi se
            ' opzionalmente c'è o meno da dare il resto
            ' (pe.Value <> pe.FaceValue) <-- Questo mi dice esatamente quale BP è in difetto rispetto al pagabile.
            '
            If resultData.totalExcedeedWithBP <> 0 And pe.Emitter <> "[[VOIDED]]" And (pe.Value <> pe.FaceValue) AndAlso RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember Then

                '
                ' Su un eccesso rispetto al  BP  corrente
                ' lo aggiungo alla transazione e con esso
                ' l'opzione che ha segnato di crearlo.
                ' E' sempre e solo una voce solitamente alla fine dei BP scansionati
                '
                PeExcedeedRec = FillMediaForExcedeed(RootTaMediaRec.PAYMENTinMedia.lChangeMediaMember, pe)

            End If

        Next

        ' Come meta riporto il restpo delle info per il raggruppamento
        RootTaMediaRec.AddField("ibp_GROUPED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_GROUPED", 1)

        If _NumVoidedT >= 1 Then

            RootTaMediaRec.AddField("ibp_TOT_BP_VOIDED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            RootTaMediaRec.setPropertybyName("ibp_TOT_BP_VOIDED", _NumVoidedT)

            RootTaMediaRec.AddField("dbp_TOT_VOIDED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            RootTaMediaRec.setPropertybyName("dbp_TOT_VOIDED", _TotValVoided)

        End If

        RootTaMediaRec.AddField("dbp_TOT_EXCEDEED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_EXCEDEED", CDec(resultData.totalExcedeedWithBP))

        RootTaMediaRec.AddField("dbp_TOT_PAYED", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
        RootTaMediaRec.setPropertybyName("dbp_TOT_PAYED", CDec(resultData.totalPayedWithBP))

        RootTaMediaRec.AddField("ibp_TOT_BP_USED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
        RootTaMediaRec.setPropertybyName("ibp_TOT_BP_USED", _NumCurrT + _NumVoidedT) ' CDec(resultData.totalBPUsed)

        Return PeExcedeedRec

    End Function


    ''' <summary>
    '''     Crea una serie di MediaRec per
    '''     ogni BP che è  stato elaborato
    '''     tramite il proxy di Argentea.
    ''' </summary>
    ''' <param name="RootTa">La TA di root in corso dove accodare tutte le TA di pagmaneto per ogni singolo BP</param>
    ''' <param name="resultData">Il set di risultati ottenuti dopo il processo di elaborazione sul proxy</param>
    ''' <returns>Se nel riepilogare i bp ci accorgiamo che è stato superato per eccesso l'importo in pagamento restituiamo in un nuovo Media il resto eventuale da gestire</returns>
    Private Function AddMediaRecsForBpElaborated(ByRef RootTa As TA, ByRef resultData As ClsProxyArgentea.DataResponse) As TaMediaRec
        Dim NewTaMediaRec As TaMediaRec

        ' Durante l'aggiunta  delle TA Media per
        ' ogni buono che ha partecipato a pagare
        ' si è superato in   qualche   modo  per 
        ' eccesso il pagamento richiesto restituiamo 
        ' in forma di TA nuova la voce della differenza.
        Dim PeExcedeedRec As TaMediaRec = Nothing

        ' 
        ' Scorro per tutti i BP nell'elenco dei
        ' BP usati nella sessione proxy per aggiungere
        ' un MEDIA per ogni BP utilizzato.
        '
        '       NOTA.: Per ogni Buono viene aggiunto un MediaRecord.
        '
        For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' proxyPos.PaidEntryBindingSource

            ' Aggiungo sulla Transazione corrente
            ' la TA relativa al Media di pagamento.
            NewTaMediaRec = AddNewTaMedia(pe)

            ' Identifico che questa voce è singola e non raggruppante
            NewTaMediaRec.AddField("ibp_GROUPED", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            NewTaMediaRec.setPropertybyName("ibp_GROUPED", 0)

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
                Throw New Exception(GLB_ERROR_NOT_CREATE_EXCEDEED)

                ' MEDIA di resto non creato per errore interno
                'm_LastStatus = GLB_ERROR_NOT_CREATE_EXCEDEED
                'm_LastErrorMessage = "Errore interno non previsto --exception on FillPaymentDataFromID to create excedeed media-- for voucher " & PeOnExcedeed.Barcode & ", return to default value."

                ' Log e segnale non aggiornato in uscita e chiusura 
                'LOG_Debug(getLocationString(funcName), m_LastErrorMessage + "--" + ex.InnerException.ToString())
                'msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            End If

        Catch Ex As Exception

            '
            ' Se fallisce la creazione  della 
            ' voce di riferimento all'eccesso
            ' in ogni caso la transazione  si
            ' conclude.
            '
            NewExcedeed = Nothing

            ' MEDIA di resto non creato per errore interno
            m_LastStatus = GLB_ERROR_ON_CREATE_EXCEDEED
            m_LastErrorMessage = "Errore interno non previsto --exception on create media excedeed-- (Chiamare assistenza)"

            ' Msg Utente
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            ' Log locale
            LOG_Error(getLocationString(funcname), m_LastStatus + " - " + m_LastErrorMessage)
            LOG_ErrorInTry(getLocationString("HandleControllerArgentea"), Ex)

        End Try

        Return NewExcedeed

    End Function

#End Region

End Class
