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


''' <summary>
'''     Di supporto alla codifica del protocollo
'''     Argentea definisce tutto lo stile CSV in
'''     returnString alla chiamata remota per 
'''     poi essere formattato secondo un oggetto
'''     interno globale da usare nei flow applicativi interni
'''     (di appoggio alla funzione ParseReturnString del <see cref="CSVHelper"/>) 
''' </summary>
Public Class ArgenteaFunctionReturnObject

    ''' <summary>
    '''     Il Nome della funzione dove eseguire 
    '''     per la risposta la corretta codifica.
    '''     In Base al Metodo Argentea saranno usate
    '''     le rispettive posizioni del CSV per
    '''     valorizzare gli attributi di questo oggetto.
    ''' </summary>
    Public ArgenteaFunction As InternalArgenteaFunctionTypes

    ''' <summary>
    '''     Il Campo chiave extra che in modo universale
    '''     nelle risposte Argentea definisce l'OK o il KO Tecnico. 
    '''     Ripreso dal Result in base alla posizione solitamente in posizione 0.
    ''' </summary>
    Public Successfull As Boolean


#Region ".ctor"

    Friend Sub New(Optional ByVal StatusCurrent As Integer = 0)

        If StatusCurrent <> 0 Then

            ' ERRORI su chiamata

            If StatusCurrent = 9001 Then

                ' ERRORE GENERALE DI SOCKET
                Me.Description = "SOCKET-ERROR"

            ElseIf StatusCurrent = 9002 Then

                ' ERRORE GENERALE DI CONFIGURAZIONE MONETICA INI    
                Me.Description = "MONETICA-CONFIG-ERROR"

            ElseIf StatusCurrent = 9003 Then

                ' ERRORE POS IN TIMEOUT
                Me.Description = "POS-TIMEOUT-ERROR"

            ElseIf StatusCurrent = 9004 Then

                ' SEND DATA FAILED
                Me.Description = "POS-SEND-DATA-FAILED"

            ElseIf StatusCurrent = 9005 Then

                ' OPERAZIONE DA UTENTE ANNULLATA
                Me.Description = "POS-OPERATION-ABORTED-BY-USER"

            ElseIf StatusCurrent = 9006 Then

                ' OPERAZIONE DA UTENTE NON HA SELEZIONATO BUONI PASTO
                Me.Description = "POS-OPERATION-NOINPUTDATA-BY-USER"

            ElseIf StatusCurrent = 9007 Then

                ' OPERAZIONE CARTA BUONI PASTO NON VALIDA
                Me.Description = "POS-TICKETCARD-NOTVALID"

            ElseIf StatusCurrent = 9008 Then

                ' OPERAZIONE NON SUPPORTATA
                Me.Description = "POS-OPERATION-NOT-IMPLEMENTATED"

            ElseIf StatusCurrent = 9009 Then

                ' OPERAZIONE DATI NON RICEVUTI
                Me.Description = "POS-OPERATION-NOINPUTDATA-RECEIVED"

            ElseIf StatusCurrent = 9010 Then

                ' ERRORE GENERALE DI PARSING SU PROTOCOLLO
                Me.Description = "GEN-PARSING-ERROR"

            Else

                ' ERRORE GENERALE UKNOWED SU RISPOSTA METODO
                Me.Description = "UKNOW-ERROR"

            End If

            ' Unsuccessfull
            Me.Successfull = False

            ' Status Code Error Object Unsuccssfull
            _Status = StatusCurrent

        Else

            ' STATUS OK KO di Risposta

        End If


    End Sub


#End Region

#Region "Membri Friends"

    Private _Status As Integer

    ''' <summary>
    '''     Il Messaggio in KO quando la risposta
    '''     data ha avuto esito secondo le specifiche
    '''     di protocollo intesa con KO tecnico
    ''' </summary>
    ''' <returns>Stringa congtenente il messaggio di KO per esteso</returns>
    Public ReadOnly Property SuccessMessage() As String
        Get
            If Me.Successfull Then
                Return (Me.Description & " " & Me.Result).Trim()
            Else
                Return ""
            End If
        End Get
    End Property

    ''' <summary>
    '''     Il Messaggio in KO quando la risposta
    '''     data ha avuto esito secondo le specifiche
    '''     di protocollo intesa con KO tecnico
    ''' </summary>
    ''' <returns>Stringa congtenente il messaggio di KO per esteso</returns>
    Public ReadOnly Property ErrorMessage() As String
        Get
            If Not Me.Successfull Then
                Return (Me.Description & " " & Me.Result).Trim()
            Else
                Return ""
            End If
        End Get
    End Property

    ''' <summary>
    '''     Il Messaggio di Stato relativo alla risposta
    '''     data secondo le specifiche di protocollo.
    ''' </summary>
    ''' <returns>Stringa congtenente il messaggio di Stato per esteso</returns>
    Public ReadOnly Property Status() As Integer
        Get
            Return _Status
        End Get
    End Property

    ''' <summary>
    '''     Il Valore ripreso dall'Amount stringa
    '''     riportato in Decimal.
    ''' </summary>
    ''' <returns>Stringa congtenente il messaggio di Stato per esteso</returns>
    Public Function GetAmountValue(ByVal Fract As Integer) As Decimal
        If Me.Amount = "" Or Me.Amount = Nothing Then
            Return 0
        Else
            Return CDec(Me.Amount) / Fract
        End If
    End Function

    ''' <summary>
    '''     Restituisce codificato lo stato dell'operazione
    '''     rispetto al suo ReturnCode da protocollo e lo stato
    '''     interno di risposta.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetStatusOperation() As String
        Return Me.Result & "-" & Me.CodeResult
    End Function

#End Region

#Region "CSV All Fields"

    ''' <summary>
    '''     Richiesta di conferma a 0 o a 1 per i BPC o BPE
    '''     0 OK 1 Richiede altra conferma operatore
    ''' </summary>
    Public RequireCommit As Boolean = False

    ''' <summary>
    '''     EsitoCode su azione effettuata in argentea 
    '''     indica il codice esito se OK solitamente 0
    '''     se KO indica il codice di Errore
    ''' </summary>
    Public CodeResult As String = ""

    ''' <summary>
    '''     OK o KO da Argentea solitamente in posizione 0
    ''' </summary>
    Public Result As String = ""

    ''' <summary>
    '''     Descrizione (se KO la descrizione dell'errore remoto in OK un Messaggio relativo all'azione)
    ''' </summary>
    Public Description As String = ""

    ''' <summary>
    '''     L'ID della Transazione in corso rilasciato
    '''     dal sistema remoto Argentea
    ''' </summary>
    Public TerminalID As String = ""

    ''' <summary>
    '''     L'Amount sia su Pagamenti o il Valore del 
    '''     Buono su BPC o BPE o l'importo che si sta 
    '''     pagando in un azione di pagamento.
    ''' </summary>
    Public Amount As String = ""

    ''' <summary>
    '''     
    ''' </summary>
    Public Acquirer As String = ""

    ''' <summary>
    '''     Per i pagamenti è lo scontrino elettronico rilasciato dal POS
    ''' </summary>
    Public Receipt As String = ""

    ''' <summary>
    '''     Per i pagamenti idntifica l'ABI bancario
    ''' </summary>
    Public Abi As String = ""

    ''' <summary>
    '''     Indica il circuito del tipo di buono cartacet
    ''' </summary>
    Public Provider As String = ""

    ''' <summary>
    '''     Da protocollo corrisponde a Codici Emettitori di Ticket (RFU)
    ''' </summary>
    Public CodeIssuer As String = ""

    ''' <summary>
    '''     Da protocollo corrisponde al nome dell'Emettiotre dei Buoni Pasto 
    ''' </summary>
    Public NameIssuer As String = ""

    ''' <summary>
    '''     Nelle chiamate e risposte dal terminale hardware
    '''     il numero dei buoni che sono stati usati sul sistema.
    ''' </summary>
    Public NumBPEvalutated As Integer = 0

    ''' <summary>
    '''     Nelle chiamate e risposte dal terminale hardware
    '''     L'elenco create fittizio di riporto alle funzioni 
    '''     suddiviso per ogni taglio consumato.
    ''' </summary>
    Public ListBPsEvaluated As System.Collections.Generic.Dictionary(Of String, Decimal)

    Public NodeXML As String = ""

    Public CouponCode As String = ""
    Public CouponCancelReason As String = ""
    Public CouponTransID As String = ""
    Public SkuSold As String = ""
    Public SkuList As String = ""
    Public SkuSaleNum As String = ""
    Public ClientCode As String = ""
    Public PosData As String = ""
#End Region

End Class
