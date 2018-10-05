Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports TPDotnet.IT.Common.Pos.EFT

Public Class BPEController
    Implements IBPDematerialize

    ' COSTANTI PARAMETRI UTILIZZATE in Operator su IT.Parameter
    Private Const OPT_BPAccorpateBPe As String = "OPT_BPAccorpateBPe"           ' <-- Accorpa nei Media l'insieme dei BP usati o meno

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
    Private pParams As BPParameters                     ' Parametri Interni per BP

    ' 
    ' Interni per gestione
    '
    Protected m_taobj As TA
    Protected m_TheModcntr As ModCntr

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
        Dematerialize = IBPReturnCode.KO
        Dim funcName As String = "Dematerialize"

        '
        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPCParameters.LoadCommonFunctionParameter
        '
        pParams = New BPParameters()

        Try
            LOG_Debug(getLocationString(funcName), "We are entered In Argentea IExternalGiftCardActivation Function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            ' Recuperiamo e valorizziamo i parametri per reflection mode byref
            pParams.LoadCommonFunctionParameter(Parameters)

            ' Imporftanti per la gestione sul controller
            ' corrente e che non possono mancare mai.
            m_TheModcntr = pParams.Controller
            m_taobj = pParams.Transaction

            If m_TheModcntr Is Nothing Or m_taobj Is Nothing Then

                ' Etichettiamo l'errore per la gestione
                pParams.Status = GLB_ERROR_INTERNAL_START
                pParams.ErrorMessage = "Per la corretta esecuzione del proxy sono necessari il controller e la transazione in corso"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            End If

            ' Riprendo dalla TA solo i prodotti relativi a quelli 
            ' che possono essere pagati con i Buoni Pasto
            Dim sTotalTransaction As String = Nothing
            If Not Common.ApplyFilterStyleSheet(m_TheModcntr, m_taobj, "BPCType.xslt", sTotalTransaction) Then

                ' Etichettiamo l'errore per la gestione
                pParams.Status = GLB_ERROR_FILE_FILTER
                pParams.ErrorMessage = "File di trasformazione per la vendita con Buoni Pasto non valido o non presente"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            End If

            ' Se invece siamo con un totale a 0 da pagare perchè tutti i prodotti
            ' sono esclusi dalla vendita possibile tramite i Buoni Pasto.
            Dim TotalTransaction As Decimal = CDec(sTotalTransaction) / 100

            If TotalTransaction = 0 Then

                ' Valorizzo l'errore nei Params per uscire dall
                ' iterazione del modulo corrente
                pParams.Status = GLB_ERROR_FILTER
                pParams.ErrorMessage = "Non sono presenti prodotti in questa vendita che è possibile pagare con i Buoni Pasto"
                '
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            Else

                ' Questo è quello che si può pagare in totale.
                m_PayableAmout = TotalTransaction

                ' Questo è quello che ho già pagato
                Dim m_PayedAmount As Decimal = pParams.MediaRecord.dTaPaidTotal

                ' Quindi controlliamo se il totale intanto è maggiore di 0 quindi pagabile.
                If pParams.MediaRecord.dTaPaidTotal > 0 Then

                    ' Ricavo l'Amount rispetto al Totale
                    m_PayableAmout = Math.Min(m_taobj.GetTotal - (m_taobj.GetTotal - m_PayedAmount), m_PayableAmout)

                End If

            End If

            ' Quindi controlliamo se il totale intanto è maggiore di 0 quindi pagabile.
            If pParams.MediaRecord.dTaPaidTotal > 0 Then

                ' E se l'ammontare del pagmaneto non eccede il totale dovuto
                If m_PayableAmout <= 0 Then

                    ' Valorizzo l'errore nei Params per uscire dall
                    ' iterazione del modulo corrente
                    pParams.Status = GLB_ERROR_PAYABLE
                    pParams.ErrorMessage = "Il pagamento eccede sul totale"
                    '
                    msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                    LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

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
                If HandlePaymentBPElettronico(m_PayableAmout) Then
                    Dematerialize = IBPReturnCode.OK
                End If
                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

            Else

                ' Valorizzo l'errore nei Params per uscire dall
                ' iterazione del modulo corrente
                pParams.Status = GLB_ERROR_PAYABLE2
                pParams.ErrorMessage = "Il pagamento ricopre già l'intero totale"
                '
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

                Dematerialize = IBPReturnCode.KO
                Exit Function

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
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
        Void = IBPReturnCode.KO
        Dim funcName As String = "Void"

        '
        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPCParameters.LoadCommonFunctionParameter
        '
        pParams = New BPParameters()

        Try
            LOG_Debug(getLocationString(funcName), "We are entered In Argentea IExternalGiftCardActivation Function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            ' Recuperiamo e valorizziamo i parametri per reflection mode byref
            pParams.LoadCommonFunctionParameter(Parameters)

            ' Imporftanti per la gestione sul controller
            ' corrente e che non possono mancare mai.
            m_TheModcntr = pParams.Controller
            m_taobj = pParams.Transaction

            If m_TheModcntr Is Nothing Or m_taobj Is Nothing Then

                ' Etichettiamo l'errore per la gestione
                pParams.Status = GLB_ERROR_INTERNAL_START
                pParams.ErrorMessage = "Per la corretta esecuzione del proxy sono necessari il controller e la transazione in corso"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Void Argentea ::KO:: Local")

                Return IBPReturnCode.KO

            End If

            ' Riprendo l'importo da stornare per l'azione
            ' sulla riga corrente della TA (Con i controli del caso)
            Dim sTotalVoid As String = Nothing
            If pParams.MediaRecord Is Nothing Or
                pParams.MediaRecord.dTaPaid >= 0 Or
                pParams.MediaRecord.PAYMENTinMedia.szExternalID <> "BPE" Then

                ' Etichettiamo l'errore per la gestione
                pParams.Status = GLB_ERROR_FILE_FILTER
                pParams.ErrorMessage = "Funzione di storno non richiamabile per chiamata a storno non congruente"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Void Argentea ::KO:: Local")

                Return IBPReturnCode.KO
            Else

                ' Importo di storno
                sTotalVoid = pParams.MediaRecord.dTaPaid

            End If

            ' Se invece siamo con uno storno a >= 0 da fare perchè tutti 
            ' l'erore è nella procedura di configuraione.
            Dim TotalToVoid As Decimal = CDec(sTotalVoid) '/ 100

            If TotalToVoid >= 0 Then

                ' Valorizzo l'errore nei Params per uscire dall
                ' iterazione del modulo corrente
                pParams.Status = GLB_ERROR_FILTER
                pParams.ErrorMessage = "Errore di forma nello storno"
                '
                msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                LOG_Debug(getLocationString(funcName), "Transaction Void Argentea ::KO:: Local")

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
                If HandleVoidBPElettronico(m_VoidAmount) Then
                    Void = IBPReturnCode.OK
                End If
                '
                '** >>>>>>>> Handle sul service remoto Argentea <<<<<<<<<<<<< **

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            LOG_FuncExit(getLocationString(funcName), Void.ToString())
        End Try

    End Function
#End Region

#Region "Handler Internal Form Action"

    ''' <summary>
    '''     Handle per azionare il pagamento tramite Terminale
    '''     POS collegato a questa cassa.
    ''' </summary>
    ''' <returns>Boolean True (Conslusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandlePaymentBPElettronico(ByVal dPayableAmount As Decimal) As Boolean
        Dim funcName As String = "HandlePaymentBPCartaceo"
        Dim service As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandlePaymentBPElettronico = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al service di Argentea.
            '
            If service Is Nothing Then
                service = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                ClsProxyArgentea.enTypeProxy.Pos,       '   <-- Il Proxy servizio avviato in modalità
                pParams.TransactionID,                  '   <-- L'id della transazione in corso
                pParams.MediaRecord.dTaPaidTotal        '   <-- Il Pagato fino adesso insieme agli altri media
            )

                '
                ' Preparo ad accettare l'handler degli eventi gestiti
                ' prima e dopo la comunicazione con il POS locale.
                '
                AddHandler service.Event_ProxyCollectDataTotalsAtEnd, AddressOf ProxyCollectDataTotalsAtEnd_Handler
                'AddHandler service.Event_ProxyCollectDataReturnedAtEnd, AddressOf ProxyCollectDataReturnedAtEnd_Handler

            End If

            '
            ' Preparo l'oggetto a quello che si deve a  spettare
            ' come totale da pagare e quello pagabile
            '
            service.Paid = 0
            service.Payable = dPayableAmount

            ' Definisco questa  variabile  Privata 
            ' per il conteggio dei Buoni eventuali
            ' già presenti nella TA che sono stati
            ' usati in precedenza dall'operatore.
            _InitialBPPayed = pParams.GetAlreadyOnTAScanned()

        Catch ex As Exception

            ' Etichettiamo l'errore per la gestione
            pParams.Status = GLB_ERROR_INSTANCE_SERVICE
            pParams.ErrorMessage = "Eccezione non gestita nell'istanza della classe servizio di Argentea"

            ' Se per qualche motivo o perchè manca il file di trasformazione
            ' o per errori in esecuzione non applica il filtro esco dalla gestione.
            msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPSTOP)
            LOG_Debug(getLocationString(funcName), "Instance Service Argentea ::KO:: Local")

            Return False

        Finally
            ''
        End Try


        ' >>>> ***************************************** <<<<<<
        '
        ' RUN -> Avvio il FORM Locale ed attendo!! with try entrapment
        '
        service.Command = ClsProxyArgentea.enCommandToCall.Payment
        If Not service.IsLive Then
            service.Connect()
        Else
            service.Unpark()
        End If
        '
        ' >>>> ***************************************** <<<<<<

        ' Del resto concludo
        Dim StatusResult As ClsProxyArgentea.enProxyStatus = service.ProxyStatus
        service.Close()

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
    '''     Handle per azionare lo storno tramite Terminale
    '''     POS collegato a questa cassa.
    ''' </summary>
    ''' <returns>Boolean True (Conslusa operazione senza eccezioni)</returns>
    Protected Overridable Function HandleVoidBPElettronico(ByVal dVoidableAmount As Decimal) As Boolean
        Dim funcName As String = "HandleVoidBPCartaceo"
        Dim service As ClsProxyArgentea = Nothing

        ' GO
        Try

            ' START & LOG
            HandleVoidBPElettronico = False
            LOG_FuncStart(getLocationString(funcName))

            '
            ' Istanza principale del Form relativo
            ' alla gestione della scansione dei codici
            ' dei Buoni Pasto e relativa validazione
            ' tramite chiamata al service di Argentea.
            '
            If service Is Nothing Then
                service = New ClsProxyArgentea(
                m_TheModcntr,                           '   <-- Il Controller di base (la cassa)
                m_taobj,                                '   <-- Il Riferimento alla transazione (per altri dati che servono)
                ClsProxyArgentea.enTypeProxy.Pos,       '   <-- Il Proxy servizio avviato in modalità
                pParams.TransactionID,                  '   <-- L'id della transazione in corso
                pParams.MediaRecord.dTaPaidTotal        '   <-- Il Pagato fino adesso insieme agli altri media
            )

                '
                ' Preparo ad accettare l'handler degli eventi gestiti
                ' prima e dopo la comunicazione con il POS locale.
                '
                AddHandler service.Event_ProxyCollectDataReturnedAtEnd, AddressOf ProxyCollectDataReturnedAtEnd_Handler

            End If


            '
            ' Preparo l'oggetto a quello che si deve a  spettare
            ' come totale da stornare e quello stornato
            '
            service.AmountVoid = 0
            service.AmountVoidable = dVoidableAmount

            ' Definisco questa  variabile  Privata 
            ' per il conteggio dei Buoni eventuali
            ' già presenti nella TA che sono stati
            ' usati in precedenza dall'operatore.
            _InitialBPPayed = pParams.GetAlreadyOnTAScanned()

        Catch ex As Exception

            ' Etichettiamo l'errore per la gestione
            pParams.Status = GLB_ERROR_INSTANCE_SERVICE
            pParams.ErrorMessage = "Eccezione non gestita nell'istanza della classe servizio di Argentea"

            ' Se per qualche motivo o perchè manca il file di trasformazione
            ' o per errori in esecuzione non applica il filtro esco dalla gestione.
            msgUtil.ShowMessage(m_TheModcntr, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPSTOP)
            LOG_Debug(getLocationString(funcName), "Instance Service Argentea ::KO:: Local")

            Return False

        Finally
            ''
        End Try


        ' >>>> ***************************************** <<<<<<
        '
        ' RUN -> Avvio il FORM Locale ed attendo!! with try entrapment
        '
        If Not service.IsLive Then
            service.Command = ClsProxyArgentea.enCommandToCall.Void
            service.Connect()
        Else
            service.Unpark()
        End If
        '
        ' >>>> ***************************************** <<<<<<

        ' Del resto concludo
        Dim StatusResult As ClsProxyArgentea.enProxyStatus = service.ProxyStatus
        service.Close()

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
            Dim OptAccorpateMediaForBP As Boolean = Microsoft.VisualBasic.IIf(m_TheModcntr.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPAccorpateBPe).Trim().ToUpper() = "Y", True, False)

            OptAccorpateMediaForBP = True

            If OptAccorpateMediaForBP Then

                ' Accorpo per l'insieme
                Dim NmRnd As Integer = Rnd(999999999999999)
                Dim ItmPe As PaidEntry = New PaidEntry("Riepilogo_BPe_" + CStr(NmRnd), resultData.totalPayedWithBP)

                ' Aggiungo sulla Transazione corrente
                ' la TA relativa al Media di pagamento.
                NewTaMediaRec = AddNewTaMedia(ItmPe)

                ' Totalizzatori per Taglio
                Dim _NumCurrT As Integer = 0
                Dim _ValCurrT As Decimal = 0

                For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' service.PaidEntryBindingSource

                    '
                    '   Accorpo nei metatag dell'unico  media
                    '   in corso per ogni taglio il numero di
                    '   tagli usati e l'importo totale.
                    '
                    Dim KeyQTA As String = "bpe_QUANTITY_" + pe.FaceValue.Replace(",", "_").Trim()
                    Dim KeyTOT As String = "bpe_AMOUNT_" + pe.FaceValue.Replace(",", "_").Trim()
                    If Not NewTaMediaRec.ExistField(KeyQTA) Then
                        NewTaMediaRec.AddField(KeyQTA, DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                        NewTaMediaRec.AddField(KeyTOT, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                        _NumCurrT = 0
                        _ValCurrT = 0
                    Else
                        _NumCurrT = NewTaMediaRec.GetPropertybyName(KeyQTA)
                        _ValCurrT = NewTaMediaRec.GetPropertybyName(KeyTOT)
                    End If
                    '
                    NewTaMediaRec.setPropertybyName(KeyQTA, _NumCurrT + 1)
                    NewTaMediaRec.setPropertybyName(KeyTOT, _ValCurrT + CDec(pe.Value))

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

                '
                ' Almeno uno è stato  usato  flaggo
                ' per rimuovere il media della coda
                ' che in origine ha iniziato.
                '
                RemoveFirstOriginalMedia = True

                '
                ' Aggiungo alla transazione l'elenco
                ' del media record di tipo BP appena 
                ' creato con i dati riportati.
                '
                m_taobj.Add(NewTaMediaRec)


            Else

                '
                ' Scorro per tutti i BP nell'elenco dei
                ' BP usati nella sessione proxy per aggiungere
                ' un MEDIA per ogni BP utilizzato.
                '
                '       NOTA.: Per ogni Buono viene aggiunto un MediaRecord.
                '
                For Each pe As PaidEntry In resultData.PaidEntryBindingSource    ' service.PaidEntryBindingSource

                    ' Aggiungo sulla Transazione corrente
                    ' la TA relativa al Media di pagamento.
                    NewTaMediaRec = AddNewTaMedia(pe)

                    '
                    ' Almeno uno è stato  usato  flaggo
                    ' per rimuovere il media della coda
                    ' che in origine ha iniziato.
                    '
                    RemoveFirstOriginalMedia = True

                    '
                    ' Aggiungo alla transazione l'elenco
                    ' del media record di tipo BP appena 
                    ' creato con i dati riportati.
                    '
                    m_taobj.Add(NewTaMediaRec)

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
            If RemoveFirstOriginalMedia Then

                m_taobj.RemoveWithRefs(pParams.MediaRecord.theHdr.lTaCreateNmbr)
                m_taobj.TARefresh(False)

            End If

            'service.Disconnect()

        Catch ex As Exception

            ' Etichettiamo l'errore per la gestione
            pParams.Status = GLB_ERROR_COLLECT_DATA_SERVICE
            pParams.ErrorMessage = "Eccezione non gestita nel restituire i dati al chiamate dal set di dati Proxy"
            LOG_Debug(getLocationString(funcName), "Instance Service Argentea ::KO:: Local")

            ' Forzo il Proxy allo stato di errore in corso
            CType(sender, ClsProxyArgentea).SetStatusInError(
                funcName,
                pParams.Status,
                pParams.ErrorMessage,
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
    Private Sub ProxyCollectDataReturnedAtEnd_Handler(ByRef sender As Object, ByRef resultData As ClsProxyArgentea.DataResponse)
        Throw New NotImplementedException()
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
            .Clone(pParams.MediaRecord, pParams.MediaRecord.theHdr.lTaCreateNmbr)

            ' Porto a non linkato l'intestazione
            ' del nodo che non dipende da altri.
            .theHdr.lTaRefToCreateNmbr = 0
            .theHdr.lTaCreateNmbr = 0

            ' Nell'ordine di un elemento in quantità sempre 1 per ogni BP
            ' aggiungo la Proprietà di tipo Stringa .:   BarCode relativo al buono
            .dTaQty = 1
            .AddField("szBPC", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            .setPropertybyName("szBPC", PeTo.Barcode)

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

        LOG_Info(getLocationString("HandlePayment"), "Manage exceed for voucher " & PeOnExcedeed.Barcode)

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
            NewExcedeed.AddField("szBuonoCartaceo", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            NewExcedeed.setPropertybyName("szBuonoCartaceo", PeOnExcedeed.Barcode)
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
                LOG_Info(getLocationString("HandlePayment"), "Exceed managed with success for voucher " & PeOnExcedeed.Barcode)

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

            ' Log e segnale non aggiornato in uscita e chiusura 
            LOG_Debug(getLocationString(funcname), m_LastErrorMessage + "--" + Ex.InnerException.ToString())
            msgUtil.ShowMessage(m_TheModcntr, m_LastErrorMessage, "LevelITCommonModArgentea_" + m_LastStatus, PosDef.TARMessageTypes.TPERROR)

            LOG_ErrorInTry(getLocationString("HandlePayment"), Ex)

        End Try

        Return NewExcedeed

    End Function



#End Region

End Class
