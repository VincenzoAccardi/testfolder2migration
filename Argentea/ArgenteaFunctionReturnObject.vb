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


#Region "CSV All Fields"

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
    '''     Richiesta di conferma a 0 o a 1 per i BPC o BPE
    '''     0 OK 1 Richiede altra conferma operatore
    ''' </summary>
    Public RequireCommit As String = ""

    ''' <summary>
    '''     Indica il circuito del tipo di buono cartacet
    ''' </summary>
    Public Provider As String = ""

    ''' <summary>
    '''     Da protocollo corrisponde a Codici Emettitori di Ticket (RFU)
    ''' </summary>
    Public CodeIssuer As String = ""

#End Region

End Class
