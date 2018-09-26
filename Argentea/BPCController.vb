Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Windows.Forms

Public Class BPCController
    Implements IBPCDematerialize


#Region "Internal Members"

    ' COSTANTI PARAMETRI UTILIZZATE in Operator su IT.Parameter
    Private Const OPT_BPParameterRuppArgentea As String = "BP_ParameterRuppArgentea"        ' <-- Parametro Stringa RUPP per protocollo Argentea

    ' COSTANTI PARAMETRI UTILIZZATE in Operator su Parameter
    Private Const OPT_BPAcceptExcedeedValues As String = "BP_AcceptExcedeedValues"          ' <-- Accetta o meno il Resto sui BP Y o N
    Private Const OPT_BPNumMaxPayablesOnVoid As String = "BP_NumMaxPayablesOnVoid"          ' <-- Numero massimo di Buoni Pasto utilizzati per la vendita in corso 0 o ^n

    ' Messaggeria per codifica segnalazioni ID di errore remoti 
    Private msgUtil As New TPDotnet.IT.Common.Pos.Common

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


    ' Su Errore Pagabile rispetto alla vendita corrente, non sono 
    ' stati trovati prodotti nella vendita con applicato il filtro
    ' xslt e quindi vendita a 0 usciamo dal flow
    Private Const GLB_ERROR_FILTER As String = "ERROR_FILTER_PAYABLES_BP"

    ' Su Errore di configurazione o di sistema quando la 
    ' procedura non trova il file xslt di trasformazione 
    ' usato per riprendere i totali dei prodotti che  si
    ' possono pagare tramite BP
    Private Const GLB_ERROR_FILE_FILTER As String = "ERROR_FILE_FILTER_PAYABLES_BP"

    ' Nell'istanziare e castare il Form di appoggio
    ' a scansione dei Barcode solleva questa eccezione.
    Private Const GLB_ERROR_FORM_INSTANCE As String = "Error-INSTANCE-FORM"

    ' Nell'aggiungere elementi alla griglia o per 
    ' motivi legati alla gestione del form non
    ' previsti solleva questa eccezione interna.
    Private Const GLB_ERROR_FORM_DATA As String = "Error-FORM-FLOWDATA"

    ' BarCode già utilizzato in precedenza evitiamo
    ' di richiamare argentea per il controllo
    Private Const GLB_INFO_CODE_ALREADYINUSE As String = "Error-BARCODE-ALREADYINUSE"

    ' BarCode da rimuovere da quelli già scanditi in precedenza 
    ' non presente in elenco
    Private Const GLB_INFO_CODE_NOTPRESENT As String = "Error-BARCODE-NOTPRESENT"

    ' Nel Flow della funzione Entry il Throw non
    ' previsto.
    Private Const GLB_ERROR_NOT_UNEXPECTED As String = "Error-EXCEPTION-UNEXPECTED"

    'Private m_Transaction As TPDotnet.Pos.TA            ' internal reference to the current TA object
    'Private m_ModCntr As ModCntr                        ' internal reference to the current ModCntr object
    'Private m_TAMediaRec As TPDotnet.Pos.TaMediaRec     ' internal refefence to media payment
    Private m_VoucherValues As Decimal = 0              ' total value ticket paid
    Private m_PayableAmout As Decimal = 0               ' payable amount

    Private _InitialBPPayed As Integer = 0              ' All'ingresso il conteggio dei BP già usati nella vendita

    Private IsOpened As Boolean                         ' Open First call barcode BCP
    Private pParams As BPCParameters                    ' Parametri Interni per BPC

    Dim m_FirstCall As Boolean = False
    Dim m_TotalPayed As Decimal = 0


#End Region

#Region "Argentea specific"

    Protected ArgenteaCOMObject As ARGLIB.argpay

#End Region

    ''' <summary>
    '''     Gestiamo un wrap verso un Form da visualizzare
    '''     con l'handler degli eventi chiave.
    ''' </summary>
    ''' <param name="Parameters">Dictionary di Parametri dinamici</param>
    ''' <returns></returns>
    Public Function Dematerialize(ByRef Parameters As Dictionary(Of String, Object)) As IBPReturnCode Implements IBPCDematerialize.Dematerialize
        Dematerialize = IBPReturnCode.KO
        Dim funcName As String = "Dematerialize"

        ' Prendo attraverso il costrutto dei parametri
        ' dinamici per allocare i parametri passati  a
        ' questa funzione.
        ' I Parametri dinamici in dictionary passati a 
        ' questa funzione vengono reimmessi in Reflection 
        ' dal metodo principe di BPCParameters.LoadCommonFunctionParameter
        pParams = New BPCParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered In Argentea IExternalGiftCardActivation Function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            ' call in check mode
            ArgenteaCOMObject = Nothing
            ArgenteaCOMObject = New ARGLIB.argpay()

            ' Recuperiamo e valorizziamo i parametri per reflection mode byref
            pParams.LoadCommonFunctionParameter(Parameters)

            ' Riprendo dalla TA solo i prodotti relativi a quelli 
            ' che possono essere pagati con i Buoni Pasto
            Dim sTotalTransaction As String = Nothing
            If Not Common.ApplyFilterStyleSheet(pParams.Controller, pParams.Transaction, "BPCType.xslt", sTotalTransaction) Then

                ' Etichettiamo l'errore per la gestione
                pParams.Status = GLB_ERROR_FILE_FILTER
                pParams.ErrorMessage = "File di trasformazione per la vendita con Buoni Pasto non valido o non presente"

                ' Se per qualche motivo o perchè manca il file di trasformazione
                ' o per errori in esecuzione non applica il filtro esco dalla gestione.
                msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
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
                msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
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
                    m_PayableAmout = Math.Min(pParams.Transaction.GetTotal - (pParams.Transaction.GetTotal - m_PayedAmount), m_PayableAmout)

                End If

            End If

            ' QUindi controlliamo se il totale intanto è maggiore di 0 quindi pagabile.
            If pParams.MediaRecord.dTaPaidTotal > 0 Then

                ' E se l'ammontare del pagmaneto non eccede il totale dovuto
                If m_PayableAmout <= 0 Then

                    ' Valorizzo l'errore nei Params per uscire dall
                    ' iterazione del modulo corrente
                    pParams.Status = GLB_ERROR_PAYABLE
                    pParams.ErrorMessage = "Il pagamento eccede sul totale"
                    '
                    msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                    LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

                    Dematerialize = IBPReturnCode.KO
                    Exit Function
                End If

                ' Richiamiamo il Form di Azione per visualizzarlo
                ' passando l'intera transazione il Controller corrente
                ' Il Metodo di pagamento passato come argomento
                If HandlePaymentBPCartaceo(m_PayableAmout) Then
                    Dematerialize = IBPReturnCode.OK
                End If

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            LOG_FuncExit(getLocationString(funcName), Dematerialize.ToString())
        End Try

    End Function



#Region "Handler Form Action"

    ''' <summary>
    '''     Handle the payment of one or more Paper Voucher
    ''' </summary>
    ''' <returns>Boolean</returns>
    Protected Overridable Function HandlePaymentBPCartaceo(ByVal dPayableAmount As Decimal) As Boolean
        Dim funcName As String = "GetPaidAmount"
        Dim frm As FormBuonoChiaro = Nothing
        Dim MyBuonoCartaceoTaMediaRec As TaMediaRec
        Dim SelectedMedia As clsSelectMedia = Nothing
        Dim Excedeed As TaMediaRec = Nothing

        m_VoucherValues = 0

        Try

            HandlePaymentBPCartaceo = False
            LOG_FuncStart(getLocationString(funcName))

            frm = pParams.Controller.GetCustomizedForm(GetType(FormBuonoChiaro), STRETCH_TO_SMALL_WINDOW)
            frm.theModCntr = pParams.Controller
            frm.taobj = pParams.Transaction

            AddHandler frm.BarcodeRead, AddressOf BarcodeReadHandler
            AddHandler frm.BarcodeRemove, AddressOf BarcodeRemoveHandler
            frm.Paid = Format(0, pParams.Controller.getFormatString4Price())
            frm.Payable = Format(dPayableAmount, pParams.Controller.getFormatString4Price())
            frm.Show() ' non modal VB Dialog

            ' Dispongo le proprietà del Controller
            ' per cominciare ad accettare le scansioni
            pParams.Controller.DialogActiv = True
            pParams.Controller.DialogFormName = frm.Text
            pParams.Controller.SetFuncKeys((False))

            ' Definisco questa variabile Privata 
            ' per il conteggio dei Buoni eventuali
            ' già presenti nella TA che sono stati
            ' usati in precedenza dall'operatore.
            _InitialBPPayed = pParams.GetAlreadyOnTAScanned()

            ' Valorizziamo il RUPP Account dallla confifurazione utente sul DB
            pParams.RUPP = pParams.Controller.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPParameterRuppArgentea)

            ' dialog is running, we will wait until Dialog is completed
            frm.bDialogActive = True
            Do While frm.bDialogActive = True
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()
            Loop


            ' Infine ricreo il Media Record da consegnareù
            ' alla cassa per il pagmaneto in corso.
            '
            '       NOTA.: Per ogni Buono viene aggiunto un MediaRecord.
            '
            For Each pe As PaidEntry In frm.PaidEntryBindingSource

                ' Preparo la MediaRecord iniziale con cui
                ' sono entrato per farne un clone di rimpiazzo.
                MyBuonoCartaceoTaMediaRec = pParams.Transaction.CreateTaObject(Of TaMediaRec)(PosDef.TARecTypes.iTA_MEDIA)

                ' Ne clono e aggiungo le propietà
                ' relative a quelle del Buono Pasto.
                With MyBuonoCartaceoTaMediaRec
                    .Clone(pParams.MediaRecord, pParams.MediaRecord.theHdr.lTaCreateNmbr)
                    .theHdr.lTaRefToCreateNmbr = 0
                    .theHdr.lTaCreateNmbr = 0

                    ' Nell'ordine di un elemento 
                    ' aggiungo la Proprietà di tipo Stringa .:   BarCode relativo al buono
                    .dTaQty = 1
                    .AddField("szBPC", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                    .setPropertybyName("szBPC", pe.Barcode)

                    ' aggiungo il valore di facciata che mi ha restituito Argentea
                    .AddField("szFaceValue", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                    .setPropertybyName("szFaceValue", pe.FaceValue)


                    'CP#1337781:1:  default there's not rest, the mv was truncated
                    .dTaPaid = Convert.ToDecimal(pe.Value)
                    .dTaPaidTotal = Convert.ToDecimal(pe.Value)
                    .dPaidForeignCurr = Convert.ToDecimal(pe.Value)
                    .dPaidForeignCurrTotal = Convert.ToDecimal(pe.Value)

                    'CP#1337781:3: if need exceed force the rest amount, set the media paid as mv face value
                    If pe.Value <> pe.FaceValue AndAlso .PAYMENTinMedia.lChangeMediaMember > 0 Then
                        LOG_Info(getLocationString("HandlePayment"), "Manage exceed for voucher " & pe.Barcode)
                        .dTaPaid = Convert.ToDecimal(pe.FaceValue)
                        .dTaPaidTotal = Convert.ToDecimal(pe.FaceValue)
                        .dPaidForeignCurr = Convert.ToDecimal(pe.FaceValue)
                        .dPaidForeignCurrTotal = Convert.ToDecimal(pe.FaceValue)

                        Try
                            SelectedMedia = createPosModelObject(Of clsSelectMedia)(pParams.Controller, "clsSelectMedia", 0, True)

                            Excedeed = pParams.Transaction.CreateTaObject(Of TaMediaRec)(PosDef.TARecTypes.iTA_MEDIA)
                            Excedeed.theHdr.lTaRefToCreateNmbr = 0
                            Excedeed.theHdr.lTaCreateNmbr = 0
                            Excedeed.dTaQty = 1
                            Excedeed.dReturn = (Convert.ToDecimal(pe.FaceValue) - Convert.ToDecimal(pe.Value))
                            Excedeed.AddField("szBuonoChiaro", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                            Excedeed.setPropertybyName("szBuonoChiaro", pe.Barcode)
                            Excedeed.AddField("szFaceValue", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                            Excedeed.setPropertybyName("szFaceValue", pe.FaceValue)

                            ' fill with payment exceed container
                            If SelectedMedia.FillPaymentDataFromID(pParams.Controller, Excedeed.PAYMENTinMedia,
                                                            .PAYMENTinMedia.lChangeMediaMember, pParams.Transaction, pParams.Transaction.colObjects) Then
                                LOG_Info(getLocationString("HandlePayment"), "Exceed managed with success for voucher " & pe.Barcode)
                            Else
                                Throw New Exception("Exceed managed with error in FillPaymentDataFromID for voucher " & pe.Barcode & ", return to default value.")
                            End If

                        Catch Ex As Exception
                            'CP#1337781:3: if error occured return to default value to continue the transaction
                            ' eg: szModulName clsSelectMedia not found in POSModel
                            .dTaPaid = Convert.ToDecimal(pe.Value)
                            .dTaPaidTotal = Convert.ToDecimal(pe.Value)
                            .dPaidForeignCurr = Convert.ToDecimal(pe.Value)
                            .dPaidForeignCurrTotal = Convert.ToDecimal(pe.Value)
                            Excedeed = Nothing
                            LOG_ErrorInTry(getLocationString("HandlePayment"), Ex)
                        End Try
                    End If
                End With

                ' Chiudo alla transazione l'elenco
                ' dei media record appena creati.
                pParams.Transaction.Add(MyBuonoCartaceoTaMediaRec)

                ' CP#1337781:3: if exceed exists add it to transaction
                If Not Excedeed Is Nothing Then
                    pParams.Transaction.Add(Excedeed)
                End If

                HandlePaymentBPCartaceo = True
            Next

            If HandlePaymentBPCartaceo Then
                ' remove original media record
                pParams.Transaction.RemoveWithRefs(pParams.MediaRecord.theHdr.lTaCreateNmbr)
                pParams.Transaction.TARefresh(False)
            End If

            frm.Hide()
            System.Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            LOG_ErrorInTry(getLocationString(funcName), ex)
        Finally
            Try
                If Not frm Is Nothing Then
                    pParams.Controller.DialogActiv = False
                    pParams.Controller.DialogFormName = ""
                    pParams.Controller.SetFuncKeys((True))
                    pParams.Controller.EndForm()
                    frm.Close()
                    frm = Nothing
                End If
            Catch ex As Exception
                LOG_Error(getLocationString(funcName), ex)
            End Try
        End Try

        LOG_FuncExit(getLocationString(funcName), "return : " + HandlePaymentBPCartaceo.ToString)

    End Function

#End Region

#Region "Handler BarCode Action"

    ''' <summary>
    '''     Stati di risposta in azione 
    '''     sulle operazioni di Modulo.
    ''' </summary>
    Enum StatusCode

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

        ''' <summary>
        '''     Errore del parser su protocollo
        '''     di risposte argentea
        ''' </summary>
        ERRORPARSING

    End Enum


    ''' <summary>
    '''     Event handler for the payment operation. A barcode has been scanned or manually inserted.
    ''' </summary>
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
                pParams.Barcode = barcode

                '
                ' Opzione Max BP pagabili per vendita.: 
                '       Se il Numero di Buoni pagabili per una vendita
                '       è superiore al numero di buoni passato procediamo
                '       con la sgnazlazione.
                Dim OptPayablesBP As Integer = CInt(pParams.Controller.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPNumMaxPayablesOnVoid).Trim())

                '
                If OptPayablesBP <> 0 And (((pParams.BPCList.Count + 1) > OptPayablesBP) Or ((_InitialBPPayed + 1) > OptPayablesBP)) Then

                    pParams.Status = GLB_OPT_ERROR_NUMEBP_EXCEDEED
                    pParams.ErrorMessage = "Il numero di buoni pasto per questa vendita è stato superato!!"

                    msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPINFORMATION)
                    Return

                End If

                'Controllo se nell'elenco è già stato considerato il BarCode
                If pParams.BPCList.ContainsKey(pParams.Barcode) Then

                    ' Status di Errore interno da segnalare
                    pParams.Status = GLB_INFO_CODE_ALREADYINUSE
                    pParams.ErrorMessage = "Il barcode è già stato usato per questa vendita!!"

                    msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPINFORMATION)
                    Return
                End If

                ' La prima chiamata Apre la sessione in Argentea
                If Not m_FirstCall Then

                    ' Chiama per  la  Dematirializzazione
                    ' e incrementa di uno il numero delle
                    ' chiamate interne.
                    FormHelper.ShowWaitScreen(pParams.Controller, False, sender)
                    Inizializated = Me.CallInitialization(funcName)
                    m_FirstCall = True

                End If

                If Inizializated Then

                    ' Mostriamo il Wait
                    FormHelper.ShowWaitScreen(pParams.Controller, False, sender)

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
                    Dim OptAcceptExceeded As Boolean = IIf(pParams.Controller.getParam(PARAMETER_MOD_CNTR + "." + "Argentea" + "." + OPT_BPAcceptExcedeedValues).Trim().ToUpper() = "Y", True, False)
                    pParams.ValueExcedeed = 0
                    '
                    If _CallDematerialize <> StatusCode.KO Then

                        If OptAcceptExceeded Then
                            '
                            '       --> Accetta eccesso su Totale da Pagare
                            '               Alla fine scrive il media don i due riporti
                            '               concludendo il pagamento a totale.
                            '
                            pParams.ValueExcedeed = pParams.Value - pParams.MediaRecord.dTaPaidTotal

                        Else

                            '
                            '       --> Non Accetta eccesso su Totale da Pagare
                            '               Richiama Argentea per fare l'annullo
                            '               alla demateriliazzazione fatta in precedenza
                            '
                            pParams.ValueExcedeed = pParams.MediaRecord.dTaPaidTotal - pParams.Value

                            If pParams.ValueExcedeed < 0 Then
                                ' Status di Errore interno da segnalare
                                pParams.Status = GLB_OPT_ERROR_VALUE_EXCEDEED
                                pParams.ErrorMessage = "Il Valore del Buono Pasto eccede il valore rispetto al totale (non è possibile dare resto)"

                                ' Segnalo Operatore di Cassa
                                msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                                LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Excedeed")

                                ' Immediatamente annullo verso il sistema argnetea l'operazione
                                ' Per rimuoverlo tramite il metodo stesso per l'annullo
                                pParams.UndoBPCForExcedeed = True  ' <-- permette di riutilizzare la funzione di remove senza eccezioni
                                Me.BarcodeRemoveHandler(sender, pParams.Barcode)
                                pParams.UndoBPCForExcedeed = False ' <-- Ripristino per le chiamate succesive

                                ' Torno all'nseirmento eventualemnete per optare su altri 
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
                    If _CallDematerialize = StatusCode.OK And _CallConfirmation = StatusCode.OK Then

                        ' Aggiungo in una collection specifica in uso
                        ' interno l'elemento Buono appena accodato in
                        ' modo univoco rispetto al suo BarCode.
                        Dim NewBPC As BPCType
                        NewBPC.szKey = pParams.Barcode
                        NewBPC.szVal = pParams.TerminalID
                        pParams.BPCList.Add(NewBPC.szKey, NewBPC)
                        _InitialBPPayed += 1                         ' <-- Conteggio numero di bpc usati in local per ogni ingresso sulla vendita

                        ' Per il Form in azione corrente mi
                        ' riprendo il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        faceValue = pParams.Value
                        paidValue = GetUsedValue(faceValue)
                        m_TotalPayed += paidValue

                        Try
                            ' Riprendo il sender che p il Form
                            ' dove voglio aggiungere alla lista
                            ' l'n elemento appena validato.
                            formBC = TryCast(sender, FormBuonoChiaro)
                            If formBC Is Nothing Then Throw New Exception(GLB_ERROR_FORM_INSTANCE)

                            ' Aggiungo l'elemento al controllo Griglia
                            formBC.PaidEntryBindingSource.Add(New PaidEntry(pParams.Barcode, paidValue.ToString("###,##0.00"), faceValue.ToString("###,##0.00"), ""))

                            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
                            formBC.Paid = m_TotalPayed.ToString("###,##0.00")

                        Catch ex As Exception
                            Throw New Exception(GLB_ERROR_FORM_INSTANCE, ex)
                        End Try

                    Else

                        ' Errata Dematerializzione o Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        FormHelper.ShowWaitScreen(pParams.Controller, True, sender)
                        msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                        LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Remote")
                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(pParams.Controller, True, sender)
                    msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                    LOG_Debug(getLocationString(funcName), "Transaction Dematerialize Argentea ::KO:: Local")

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto
                Throw New Exception(GLB_ERROR_FORM_INSTANCE)
            End If

        Catch ex As Exception

            FormHelper.ShowWaitScreen(pParams.Controller, True, sender)
            SetExceptionsStatus(funcName, ex)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            FormHelper.ShowWaitScreen(pParams.Controller, True, sender)

            ' Riporto la firstcall a false
            ' per le istanze successive.
            m_FirstCall = False

            ' Svuoto il controllo del barcode
            formBC.txtBarcode.Text = String.Empty

        End Try

    End Sub

    ''' <summary>
    '''     Event handler for remove payment from payed list inserted on previous operation call.
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
                pParams.Barcode = barcode

                'Controllo se nell'elenco è già stato considerato il BarCode
                If Not pParams.UndoBPCForExcedeed And Not pParams.BPCList.ContainsKey(pParams.Barcode) Then

                    pParams.Status = GLB_INFO_CODE_NOTPRESENT
                    pParams.ErrorMessage = "Il BarCode non è presente tra le scelte possibili!!"

                    msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPINFORMATION)
                    Return
                End If

                If Inizializated Then

                    FormHelper.ShowWaitScreen(pParams.Controller, False, sender)

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
                    If _CallUndoDematerialize = StatusCode.OK And _CallConfirmation = StatusCode.OK Then

                        ' Argormento Opzione per Opzione 
                        ' su Flow operatore se non accetta
                        ' Sulla griglia e il form non deve
                        ' fare altro dato che non è stato aggiunto.
                        If pParams.UndoBPCForExcedeed Then
                            Return
                        End If


                        ' Rimuovo dalla collection specifica in uso
                        ' interno l'elemento Buono da annullare individuandolo
                        ' in modo univoco rispetto al suo BarCode con cui era 
                        ' stato registrato all'aggiunta dell'handler di ADD.
                        Dim NewBPC As BPCType
                        NewBPC.szKey = pParams.Barcode
                        NewBPC.szVal = pParams.TerminalID
                        pParams.BPCList.Remove(NewBPC.szKey)
                        _InitialBPPayed -= 1                         ' <-- Conteggio numero di bpc usati in local per ogni ingresso sulla vendita

                        ' Per il Form in azione corrente mi
                        ' aggiorno il Totale da Pagare rispetto a
                        ' quelli già in elenco
                        faceValue = pParams.Value
                        paidValue = GetUsedValue(faceValue)
                        m_TotalPayed -= paidValue

                        Try
                            ' Riprendo il sender che p il Form
                            ' dove voglio aggiungere alla lista
                            ' l'n elemento appena validato.
                            formBC = TryCast(sender, FormBuonoChiaro)
                            If formBC Is Nothing Then Throw New Exception(GLB_ERROR_FORM_INSTANCE)

                            ' Sul Form rimuovo dalla griglia l'elemento
                            formBC.PaidEntryBindingSource.RemoveCurrent()

                            ' Ed aggiorno anche il campo sul form per  il totale che rimane.
                            formBC.Paid = m_TotalPayed.ToString("###,##0.00")

                        Catch ex As Exception
                            Throw New Exception(GLB_ERROR_FORM_INSTANCE, ex)
                        End Try

                    Else

                        ' Errata Reverse per Dematerializzione o Reverse Confirm su Dematerializzazione
                        ' data dalla risposta argentea quindi su segnalazione remota.
                        FormHelper.ShowWaitScreen(pParams.Controller, True, sender)
                        msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                        LOG_Debug(getLocationString(funcName), "Transaction Reverse Dematerialize Argentea ::KO::")
                        Return

                    End If

                Else

                    ' Tutti i messaggi di errata inizializzazione sono
                    ' stati già dati loggo comunque questa informazione.
                    FormHelper.ShowWaitScreen(pParams.Controller, True, sender)
                    msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
                    LOG_Debug(getLocationString(funcName), "Transaction Reverse Dematerialize Argentea ::KO:: NOT INIZIALIZATED")

                End If

            Else
                ' Chiamata a questo Handler da un Form non previsto
                Throw New Exception(GLB_ERROR_FORM_INSTANCE)
            End If

        Catch ex As Exception

            FormHelper.ShowWaitScreen(pParams.Controller, True, sender)
            SetExceptionsStatus(funcName, ex)

        Finally

            ' In ogni caso chiudo se rimane aperto su eccezione
            FormHelper.ShowWaitScreen(pParams.Controller, True, sender)

            ' Svuoto il controllo del barcode
            If Not pParams.UndoBPCForExcedeed Then formBC.txtBarcode.Text = String.Empty

        End Try

    End Sub

    ''' <summary>
    '''     Inizializza la Sessione verso Argentea
    '''     e parte la numerazione interna delle chiamate
    '''     da 1
    ''' </summary>
    Private Function CallInitialization(funcName As String) As Boolean

        CallInitialization = False

        ' Partiamo che non sia OK l'esito su chiamata remota Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO

        ' Active to first Argentea COM communication
        retCode = ArgenteaCOMObject.OpenTicketBC(
            pParams.ProgressiveCall,
            pParams.ReceiptNumber,
            pParams.CodeCashDevice,
            pParams.MessageOut
        )

        ''' Per Test
        'pParams.MessageOut = "OK--TICKET APERTO-----0---" ' <-- x test 
        'retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

        ' Riprendiamo la Risposta cos' come è stata
        ' data per il log di debug grezza
        LOG_Debug(getLocationString(funcName), "ReturnCode: " & retCode.ToString & ". BPC: " & pParams.Barcode & ". Output: " & pParams.MessageOut)

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            Me.ParseErrorAndMapToParams(funcName, retCode, pParams.MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Activation check for BPC with  " & pParams.Barcode & " returns error: " & pParams.ErrorMessage & ". The message raw output is: " & pParams.MessageOut)
            Return False

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Me.ParseResponseAndMapToParams(funcName, pParams.MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If pParams.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                pParams.IncrementProgressiveCall()

                ' ** INIZIALIZZATA e corretamente chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BPC inizialization " & pParams.Barcode & " successfuly on call first with message " & pParams.SuccessMessage)
                Return True

            Else

                ' Non inizializzata da parte di Argentea per
                ' errore remoto in risposta a questo codice.
                LOG_Debug(getLocationString(funcName), "BPC inizialization " & pParams.Barcode & " remote failed on first call to argentea with message code " & pParams.Status & " relative to " & pParams.ErrorMessage)
                Return False

            End If

        End If

    End Function


    ''' <summary>
    '''     Esegue la chiamata di Dematerializzazione secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    Private Function CallDematerialize(funcName As String) As StatusCode

        ' CSV for Argumet return Status Call to Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallDematerialize = StatusCode.KO

        ' Active to first Argentea COM communication                                **** DEMATERIALIZZAZIONE
        retCode = ArgenteaCOMObject.DematerializzazioneBP(
                    pParams.GetCodifiqueReceipt(BPCParameters.TypeCodifiqueProtocol.Dematerialization),
                    pParams.MessageOut
                )

        ''' Per Test questo è il suio CSV
        'pParams.MessageOut = "OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--"    ' <-- x test 
        'pParams.MessageOut = "KO-3-Buono pasto gia' rientrato-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"       ' <-- x test su questo signal
        'pParams.MessageOut = "KO-903-Sequenza non valida-68123781901001800003069451200529-529-ARGENTEA-201809201733577-0-202--"            ' <-- x test su questo signal
        'retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            Me.ParseErrorAndMapToParams(funcName, retCode, pParams.MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Dematerialization for BPC with  " & pParams.Barcode & " returns error: " & pParams.ErrorMessage & ". The message raw output is: " & pParams.MessageOut)

            ' Esco dal  flow immediatamente
            Return CallDematerialize = StatusCode.KO

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Me.ParseResponseAndMapToParams(funcName, pParams.MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If pParams.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                pParams.IncrementProgressiveCall()

                ' Se la risposta argenta richiede un ulteriore 
                ' conferma allora procedo ad uscire per il flow.
                If pParams.CommittRequired Then

                    ' ** DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BPC dematirializated with wait confirm " & pParams.Barcode & " successfuly on call with message " & pParams.SuccessMessage)

                    ' RICHIESTO CONFERMA
                    Return CallDematerialize = StatusCode.CONFIRMREQUEST

                Else

                    ' ** DEMATERIALIZZATO corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BPC dematirializated " & pParams.Barcode & " successfuly on call with message " & pParams.SuccessMessage)

                    ' COMPLETATO
                    Return CallDematerialize = StatusCode.OK

                End If

            Else

                ' Non dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BPC dematirializated " & pParams.Barcode & " remote failed on call to argentea with message code " & pParams.Status & " relative to " & pParams.ErrorMessage)

                ' NON EFFETTUATO
                Return CallDematerialize = StatusCode.KO

            End If

        End If

    End Function

    ''' <summary>
    '''     Esegue la chiamata di Reverse da uno già Dematerializzato secondo
    '''     le specifiche Argentea al sistema remoto
    ''' </summary>
    Private Function CallReverseMaterializated(funcName As String) As Boolean

        ' CSV for Argumet return Status Call to Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallReverseMaterializated = StatusCode.KO

        ' Active to first Argentea COM communication                                **** ANNULLO BUONO GIA' MATERIALIZZATO
        retCode = ArgenteaCOMObject.ReverseTransactionDBP(
                    pParams.GetCodifiqueReceipt(BPCParameters.TypeCodifiqueProtocol.Reverse),
                    pParams.MessageOut
                )

        ''' Per Test
        'pParams.MessageOut = "OK-0 - BUONO STORNATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
        'retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            Me.ParseErrorAndMapToParams(funcName, retCode, pParams.MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Reverse Dematerialization for BPC with  " & pParams.Barcode & " returns error: " & pParams.ErrorMessage & ". The message raw output is: " & pParams.MessageOut)

            ' Esco dal  flow immediatamente
            Return CallReverseMaterializated = StatusCode.KO

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Me.ParseResponseAndMapToParams(funcName, pParams.MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If pParams.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                pParams.IncrementProgressiveCall()

                ' Se la risposta argenta richiede un ulteriore 
                ' conferma allora procedo ad uscire per il flow.
                If pParams.CommittRequired Then

                    ' ** REVERSE SU DEMATERIALIZZATO in check corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BPC reverse dematirializated with wait confirm " & pParams.Barcode & " successfuly on call with message " & pParams.SuccessMessage)

                    ' RICHIESTO CONFERMA
                    Return CallReverseMaterializated = StatusCode.CONFIRMREQUEST

                Else

                    ' ** REVERSE SU DEMATERIALIZZATO corretamente da chiamata ad Argentea
                    LOG_Debug(getLocationString(funcName), "BPC reverse dematirializated " & pParams.Barcode & " successfuly on call with message " & pParams.SuccessMessage)

                    ' COMPLETATO
                    Return CallReverseMaterializated = StatusCode.OK

                End If

            Else

                ' Non reverse su dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BPC reverse dematirializated " & pParams.Barcode & " remote failed on call to argentea with message code " & pParams.Status & " relative to " & pParams.ErrorMessage)

                ' NON EFFETTUATO
                Return CallReverseMaterializated = StatusCode.KO

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


        ' CSV for Argumet return Status Call to Argentea
        Dim retCode As ArgenteaFunctionsReturnCode = ArgenteaFunctionsReturnCode.KO
        CallConfirmOperation = StatusCode.KO

        ' Active to first Argentea COM communication                                **** CONFERMA DEMATERIALIZZAZIONE o REVERSE
        retCode = ArgenteaCOMObject.CommitTransactionDBP(
                    pParams.GetCodifiqueReceipt(BPCParameters.TypeCodifiqueProtocol.Confirm),
                    pParams.MessageOut
                )

        ''' Per Test
        'pParams.MessageOut = "OK-0 - BUONO CONFERMATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--" ' <-- x test 
        'retCode = ArgenteaFunctionsReturnCode.OK
        ''' to remove:

        If retCode <> ArgenteaFunctionsReturnCode.OK Then

            ' Su risposta da COM  in  negativo
            ' in ogni formatto il returnString
            ' ma con la variante che già mi filla
            ' l'attributro ErrorMessage
            Me.ParseErrorAndMapToParams(funcName, retCode, pParams.MessageOut)

            ' Non inizializzata su Errori di comunicazione
            ' o per risposta remota data da Argentea KO.
            LOG_Error(getLocationString(funcName), "Confirm " & sConfirmOperation & " for BPC with  " & pParams.Barcode & " returns error: " & pParams.ErrorMessage & ". The message raw output is: " & pParams.MessageOut)

            ' Esco dal  flow immediatamente
            Return CallConfirmOperation = StatusCode.KO

        Else

            ' Riprendiamo la Risposta da protocollo Argentea
            Me.ParseResponseAndMapToParams(funcName, pParams.MessageOut)

            ' Se Argentea mi dà Successo Procedo altrimenti 
            ' sono un un errore remoto, su eccezione locale
            ' di parsing esco a priori e non passo.
            If pParams.Successfull Then

                ' Incrementiamo di uno l'azione al numero di chiamate verso argentea
                pParams.IncrementProgressiveCall()

                ' ** CONFIRM su REVERSE o DEMATERIALIZZATO effettuata corretamente da chiamata ad Argentea
                LOG_Debug(getLocationString(funcName), "BPC confirm " & sConfirmOperation & " for " & pParams.Barcode & " successfuly on call with message " & pParams.SuccessMessage)

                ' COMPLETATO
                Return CallConfirmOperation = StatusCode.OK

            Else

                ' Non confirm su reverse o dematerializzato da risposta Argentea per
                ' errore remoto in relazione a questo codice.
                LOG_Debug(getLocationString(funcName), "BPC confirm " & sConfirmOperation & " for " & pParams.Barcode & " remote failed on call to argentea with message code " & pParams.Status & " relative to " & pParams.ErrorMessage)

                ' NON EFFETTUATO
                Return CallConfirmOperation = StatusCode.KO

            End If

        End If

    End Function

    ''' <summary>
    '''     Sui OK di Argentea eseguo il Parsing della Risposta
    '''     per formulare la risposta in esito (Success/Not Success) in codifica da protocollo.
    '''     Mappa gli Attributi su pParams del modulo per lo sbroglio del Flow
    ''' </summary>
    ''' <param name="funcName">Il Nome della Funzione che ha catturato la risposta</param>
    ''' <param name="CSV">Il MessageOut da codificare</param>
    Private Sub ParseResponseAndMapToParams(funcName As String, CSV As String)

        ' Tipo di codifica generalizzata Argentea wrappatra su un ReturnObject
        Dim objTPTAHelperArgentea(0), ResponseArgentea As ArgenteaFunctionReturnObject
        objTPTAHelperArgentea(0) = New ArgenteaFunctionReturnObject()

        ' Parsiamo la risposta argentea
        If CSV = "ERRORE SOCKET" Then

            pParams.Successfull = False

            ' E riporto nell'ordine corretto il messaggio di stato.
            pParams.SuccessMessage = ""
            pParams.ErrorMessage = "Errore di connessione"
            pParams.Status = 9001
            Return
        End If

        ' Parsiamo la risposta argentea per l'azione BPC
        If (Not CSVHelper.ParseReturnString(CSV, InternalArgenteaFunctionTypes.BPCPayment, objTPTAHelperArgentea, "-")) Then

            LOG_Debug(getLocationString(funcName), "BPC Parsing Protcol Argentea Fail to Parse 'Message Response' for this " & funcName & " response in MessageOut")

            ' Su Errore di Parsing solleviamo immediatamente l'eccezione per uscire dalla
            ' gestione della comunicazione Argentea.
            Throw New Exception(GLB_ERROR_PARSING)

        Else

            ' RIPORTO SUL FLOW quelli concerni allo Stato di OK Success o KO Error

            ' Risposta Codicficata da Risposta Raw Argentea
            ResponseArgentea = objTPTAHelperArgentea(0)

            ' Quindi mi riporto lo stato dell'operazione
            ' data dalla risposta remota di argentea.
            pParams.Successfull = ResponseArgentea.Successfull

            ' E riporto nell'ordine corretto il messaggio di stato.
            pParams.SuccessMessage = ""
            pParams.ErrorMessage = ""
            If pParams.Successfull Then
                pParams.SuccessMessage = (ResponseArgentea.Description & " " & ResponseArgentea.Result).Trim()
            Else
                pParams.ErrorMessage = (ResponseArgentea.Description & " " & ResponseArgentea.Result).Trim()
            End If

            ' E per questa specifica Azione fortunatamente
            ' abbiamo il codice di Stato
            pParams.Status = ResponseArgentea.CodeResult

            ' Riprendo queste  notazioni  rispettivamente
            ' per il TerminalID che ha eseguito  la trans
            ' il Valore del Buono Pasto dato da  Argentea
            ' e soprattutto se richiede un ulteriore call
            ' verso argentea di conferma alla trans.
            pParams.TerminalID = ResponseArgentea.TerminalID
            pParams.Value = CDec(ResponseArgentea.Amount) / 100
            pParams.CommittRequired = CBool(ResponseArgentea.RequireCommit)

        End If

    End Sub

    ''' <summary>
    '''     Sui KO di Argentea eseguo il Parsing della Risposta
    '''     per formulare l'errore in codifica da protocollo.
    '''     Mappa gli Attributi su pParams del modulo per lo sbroglio del Flow
    ''' </summary>
    ''' <param name="funcName">Il Nome della Funzione che ha sollevato l'errore</param>
    ''' <param name="CSV">Il MessageOut da codificare</param>
    Private Sub ParseErrorAndMapToParams(funcName As String, retCode As ArgenteaFunctionsReturnCode, CSV As String)

        ' Tipo di codifica generalizzata Argentea wrappatra su un ReturnObject
        Dim objTPTAHelperArgentea(0), ResponseArgentea As ArgenteaFunctionReturnObject
        objTPTAHelperArgentea(0) = New ArgenteaFunctionReturnObject()

        ' Parsiamo la risposta argentea
        If CSV = "ERRORE SOCKET" Then

            pParams.Successfull = False

            ' E riporto nell'ordine corretto il messaggio di stato.
            pParams.SuccessMessage = ""
            pParams.ErrorMessage = "Errore di connessione"
            pParams.Status = 9001
            Return
        End If

        If (Not CSVHelper.ParseReturnString(CSV, InternalArgenteaFunctionTypes.BPCPayment, objTPTAHelperArgentea, "-")) Then

            LOG_Debug(getLocationString(funcName), "BPC Parsing Protcol Argentea Fail to Parse 'Error' for this " & funcName & " response in MessageOut")

            ' Su Errore di Parsing solleviamo immediatamente l'eccezione per uscire dalla
            ' gestione della comunicazione Argentea.
            Throw New Exception(GLB_ERROR_PARSING)

        Else

            ' RIPORTO SUL FLOW quelli concerni all'errore

            ' Risposta Codicficata da Risposta Raw Argentea
            ResponseArgentea = objTPTAHelperArgentea(0)

            ' Quindi mi riporto lo stato dell'operazione
            ' data dalla risposta remota di argentea.
            pParams.Successfull = False

            ' E riporto nell'ordine corretto il messaggio di stato.
            pParams.SuccessMessage = ""
            pParams.ErrorMessage = (ResponseArgentea.Description & " " & ResponseArgentea.Result).Trim()

            ' E per questa specifica Azione fortunatamente
            ' abbiamo il codice di Stato
            pParams.Status = retCode.ToString() & "-" & ResponseArgentea.CodeResult

        End If

    End Sub

    Private Sub SetExceptionsStatus(funcname As String, ex As Exception)

        If ex.Message = GLB_ERROR_PARSING Then

            ' Se una funzione nel Flow mi ha dato picca
            ' ed ha sollevato questa eccezione di parsing
            ' sulla chiamat interna.
            pParams.Status = GLB_ERROR_PARSING
            pParams.ErrorMessage = "Errore di parsing sul protocollo Argentea (Chiamare assistenza)"

        ElseIf ex.Message = GLB_ERROR_FORM_INSTANCE Then

            ' Nell'instanziare il Form è successo qualche       *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            ' errore di valutazione sul tipo specifico.
            pParams.Status = GLB_ERROR_FORM_INSTANCE
            pParams.ErrorMessage = "Errore interno alla procedura --istance or object-- (Chiamare assistenza)"

        ElseIf ex.Message = GLB_ERROR_FORM_DATA Then

            ' Eventuali errori interni sull'iterazione con      *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            ' i controlli grigli e altro nel Form legato alla
            ' presentazione dei dati con errore non previsto
            pParams.Status = GLB_ERROR_FORM_DATA
            pParams.ErrorMessage = "Errore interno alla procedura --data values or stream-- (Chiamare assistenza)"

        Else

            ' Altro non previsto in questa funzione             *** Prestare attenzione qui potrebbe essere che la transazione sia stata comunque completata
            pParams.Status = GLB_ERROR_FORM_DATA
            pParams.ErrorMessage = "Errore interno alla procedura --exception unexcpted-- (Chiamare assistenza)"

        End If
        '
        LOG_Debug(getLocationString(funcname), pParams.ErrorMessage + "--" + ex.InnerException.ToString())
        msgUtil.ShowMessage(pParams.Controller, pParams.ErrorMessage, "LevelITCommonModArgentea_" + pParams.Status, PosDef.TARMessageTypes.TPERROR)
        '

    End Sub


    ''' <summary>
    '''     Gets the voucher have to use, if FaceValue minor than Payable, voucher can be truncated
    ''' </summary>
    ''' <returns>Amount of Totale subtract of current BPC</returns>
    Protected Function GetUsedValue(ByVal FaceValue As Decimal) As Decimal
        Dim UsedValue As Decimal = FaceValue

        If m_PayableAmout < FaceValue Then
            UsedValue = m_PayableAmout
        End If

        Return UsedValue
    End Function

    Private Function ValidationVoucherRequest(barcode As String) As Boolean
        'Logic Comunication Barcode at Argentea Supplier

        ValidationVoucherRequest = False


    End Function


#End Region

#Region "Functions Common"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region


End Class
