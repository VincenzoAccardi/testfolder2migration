#Region "Documentation"
' ********** ********** ********** **********
' IBPDematerialize
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region

Imports TPDotnet.Pos
Imports TPDotnet.IT.Common.Pos


#Region "IBPDematerialize"
Public Interface IBPDematerialize

    ''' <summary>
    '''     Azione di Dematerializzazione dei BP sia
    '''     Elettronici (Richiamerà il proxy in modalità Hardware POS)
    '''     Cartacei (Richiamerà il roxy in modalità Software SERVICE)
    ''' </summary>
    ''' <param name="Parameters">
    '''     Il Set di parametri dinamici ad uso e consumo
    '''     del Controller che implementa il metodo passati
    '''     in modo dinamico previsti sul DB di BackStore
    ''' </param>
    ''' <returns></returns>
    Function Dematerialize(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IBPReturnCode

    ''' <summary>
    '''     Azione di Storno
    '''     (direttamente da voce di selezione in TA principale)
    ''' </summary>
    ''' <param name="Parameters">
    '''     Il Set di parametri dinamici ad uso e consumo
    '''     del Controller che implementa il metodo passati
    '''     in modo dinamico previsti sul DB di BackStore
    ''' </param>
    ''' <returns></returns>
    Function Void(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IBPReturnCode


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
    ''' <returns>True se l'azione API ha dato esito OK altrimenti False</returns>
    Function Close(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object), SilentMode As Boolean) As Boolean


    ''' <summary>
    '''     Se Visualizzare o Meno i Messagi di Stato
    '''     delle operazioni in corso.
    ''' </summary>
    ''' <returns>True se è impostato per visualizzare i messaggi di stato altrimenti False</returns>
    Property SilentMode As Boolean


End Interface
#End Region