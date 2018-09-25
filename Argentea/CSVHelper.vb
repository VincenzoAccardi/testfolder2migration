﻿Imports System
Imports TPDotnet.Pos
Imports Microsoft.VisualBasic

Public Class CSVHelper

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
    '''     Esegue il parsing di una risposta dettata dal Protocollo
    '''     Argentea.
    '''     In Ordine del CSV del protocollo Argentea in punti chiave
    '''     della risposta sono.:
    '''         Successufully   Format to Boolean True/False  da Stringa   -OK/KO-   <-- Per tutti i Metodi Argentea
    '''         CodeResult      Format to Stringa 000         da Stringa   -000-     <-- Riporta il Codice di Errore o Success
    '''         Description     Format to Stringa "...."      da Stringa   -....-    <-- Descrizione dell'errore o del Success
    '''     gli altri in base all'azione
    '''         ripresi dalla posizione sul CSV e interpretati
    '''         per definire gli attributi dell'oggetto intrno
    ''          di risposta al chaiamante.
    ''' </summary>
    ''' <param name="returnString">La Stringa di risposta dopo la chiamata verso Argentea da interpretare</param>
    ''' <param name="argenteaFunction">Il Nome del metodo Argentea e quindi dell'azione su cui elaborare la risposta</param>
    ''' <param name="argenteaFunctionReturnObject">L'oggetto da fillare con gli attributi dati dalla risposta dopo codifica</param>
    ''' <param name="szCharSeparator">Il separatore nella ripsota CSV remota che solitamente è - </param>
    ''' <returns>Se True (se non ci sono cambiamenti nel protocollo Argentea errori di Parsing non ci dovrebbero essere) Altrimenti è un Errore si Parsing sul Protocollo di risposte Argentea</returns>
    Public Shared Function ParseReturnString(ByVal returnString As String, ByVal argenteaFunction As InternalArgenteaFunctionTypes, ByRef argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject, Optional ByRef szCharSeparator As String = ";") As Boolean
        ParseReturnString = False
        Dim I, J As Integer
        Dim funcName As String = "ParseReturnString"
        Dim CSV(200) As String '= Nothing
        Dim StepNum As Integer

        ' Riprendo il riferimento da restituire con gli attributi codificati
        Dim MyRefRet As ArgenteaFunctionReturnObject = argenteaFunctionReturnObject(0)

        ' Formatta il campo scpeciale Recipt senza vincoli dei vbcrlf di vb
        Dim ReplaceVbCRLF As Func(Of String, String) = Function(ByVal strFieldCSV As String) 'As String

                                                           strFieldCSV = Replace(strFieldCSV,
                                                    Microsoft.VisualBasic.vbCrLf,
                                                    Microsoft.VisualBasic.vbLf)
                                                           strFieldCSV = Replace(strFieldCSV,
                                                    Microsoft.VisualBasic.vbCr,
                                                    Microsoft.VisualBasic.vbLf)
                                                           strFieldCSV = Replace(strFieldCSV,
                                                                    Microsoft.VisualBasic.vbLf,
                                                                    Microsoft.VisualBasic.vbCrLf)
                                                           Return strFieldCSV
                                                       End Function

        ' Formatta l'attributo speciale di ritorno a boolean in base a OK KO
        Dim SetSuccessufully As Func(Of String, Boolean) = Function(ByVal strFieldCSV As String) 'As Boolean
                                                               If strFieldCSV = "OK" Then Return True Else Return False
                                                           End Function

        ' Per i valori attesi numerici  se stringa vuota riporta "0"
        Dim SetNumeric As Func(Of String, String) = Function(ByVal strFieldCSV As String) 'As String
                                                        If strFieldCSV = String.Empty Then Return "0" Else Return strFieldCSV
                                                    End Function


        Try

            ' Log che stiamo parsando la risposta
            LOG_Debug(getLocationString(funcName), "Parsing result")

            ' Gli Attributi per formattare la risposta corrente da restituire al chiamante
            'ReDim argenteaFunctionReturnObject(0)
            MyRefRet.Successfull = False
            MyRefRet.ArgenteaFunction = argenteaFunction

            ' Il CSV della risposta data dal comando Argentea da Interpretare
            CSV = returnString.Split(szCharSeparator)

            ' Cicla su ogni campo del CSV
            For I = 0 To CSV.Length - 1

                ' Ed in Base alla funzione dell'azione Argentea che ci 
                ' ha dato la risposta preleviamo il rispettivo  valore
                ' da riportare nel Nostro Oggetto di Risposta codificato.
                Select Case argenteaFunction

                    ' AZIONE DEL CASO: Pagamento o Annullo di Pagamento
                    Case InternalArgenteaFunctionTypes.EFTPayment,
                         InternalArgenteaFunctionTypes.EFTVoid

                        MyRefRet.TerminalID = CSV(0)
                        MyRefRet.Amount = CSV(1)
                        MyRefRet.Successfull = SetSuccessufully(CSV(2))
                        MyRefRet.Description = CSV(3)
                        MyRefRet.Acquirer = CSV(4)
                        MyRefRet.Receipt = ReplaceVbCRLF(CSV(5))
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONE DEL CASO: Stampa Totali dal POS o Chiusura del POS
                    Case InternalArgenteaFunctionTypes.EFTGetTotals,
                         InternalArgenteaFunctionTypes.EFTClose

                        If CInt(CSV(0)) > 1 Then ReDim Preserve argenteaFunctionReturnObject(CInt(CSV(0)) - 1)

                        ' In questo caso speciale scorro sull'array
                        ' perchè il returnString contiene più risposte
                        ' da codificare.
                        For J = 0 To argenteaFunctionReturnObject.GetUpperBound(0)
                            If J <> 0 Then argenteaFunctionReturnObject(J) = New ArgenteaFunctionReturnObject
                            StepNum = 6 * J
                            argenteaFunctionReturnObject(J).TerminalID = CSV(StepNum + 1)
                            argenteaFunctionReturnObject(J).Abi = CSV(StepNum + 2)
                            argenteaFunctionReturnObject(J).Amount = CSV(StepNum + 3)
                            argenteaFunctionReturnObject(J).Successfull = SetSuccessufully(CSV(StepNum + 4))
                            argenteaFunctionReturnObject(J).Amount = CSV(StepNum + 3)
                            argenteaFunctionReturnObject(J).Description = CSV(StepNum + 5)
                            argenteaFunctionReturnObject(J).Receipt = ReplaceVbCRLF(CSV(StepNum + 6))
                        Next J
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONDE DEL CASO: ?? non usato
                    Case InternalArgenteaFunctionTypes.EFTConfirm
                        Exit Select

                    ' AZIONDE DEL CASO: GIFT CARDS su Protocollo
                    Case InternalArgenteaFunctionTypes.GiftCardActivationPreCheck,
                        InternalArgenteaFunctionTypes.GiftCardActivation,
                        InternalArgenteaFunctionTypes.GiftCardRedeemPreCkeck,
                        InternalArgenteaFunctionTypes.GiftCardRedeem,
                        InternalArgenteaFunctionTypes.GiftCardRedeemCancel,
                        InternalArgenteaFunctionTypes.GiftCardBalance

                        MyRefRet.ArgenteaFunction = argenteaFunction
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.Receipt = ReplaceVbCRLF(CSV(1))
                        MyRefRet.Result = CSV(2)
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONDE DEL CASO: Ricariche Telefoniche da Protocollo
                    Case InternalArgenteaFunctionTypes.PhoneRechargeCheck,
                        InternalArgenteaFunctionTypes.PhoneRechargeActivation

                        MyRefRet.ArgenteaFunction = argenteaFunction
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.Receipt = ReplaceVbCRLF(CSV(1))
                        MyRefRet.Result = CSV(2)
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONDE DEL CASO: GIFTCARDS tipo smartbox applepay da protocollo
                    Case InternalArgenteaFunctionTypes.ExternalGiftCardActivation,
                         InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation

                        MyRefRet.ArgenteaFunction = argenteaFunction
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.Receipt = ReplaceVbCRLF(CSV(1))
                        MyRefRet.Result = CSV(2)
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONDE DEL CASO: Pagamenti da SatisPay da protocollo
                    Case InternalArgenteaFunctionTypes.ADVPayment

                        MyRefRet.ArgenteaFunction = argenteaFunction
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.Description = CSV(1)
                        MyRefRet.TerminalID = CSV(2)
                        MyRefRet.Amount = CSV(3)
                        MyRefRet.Receipt = ReplaceVbCRLF(CSV(4))
                        MyRefRet.Result = CSV(0)
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONDE DEL CASO: Storno dei Pagamenti da SatisPay da protocollo
                    Case InternalArgenteaFunctionTypes.ADVVoid

                        MyRefRet.ArgenteaFunction = argenteaFunction
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.Description = CSV(1)
                        MyRefRet.TerminalID = CSV(2)
                        MyRefRet.Amount = CSV(3)
                        MyRefRet.Receipt = ReplaceVbCRLF(CSV(4))
                        MyRefRet.Result = CSV(0)
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    ' AZIONDE DEL CASO: Buoni Pasto Cartacei da protocollo
                    Case InternalArgenteaFunctionTypes.BPCPayment

                        MyRefRet.ArgenteaFunction = argenteaFunction
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.CodeResult = SetNumeric(CSV(1))
                        MyRefRet.Description = CSV(2)
                        ' il 3 c'è già
                        MyRefRet.Amount = SetNumeric(CSV(4))
                        MyRefRet.Provider = CSV(5)
                        MyRefRet.TerminalID = CSV(6)
                        MyRefRet.RequireCommit = CSV(7)
                        MyRefRet.CodeIssuer = CSV(8)
                        MyRefRet.Result = CSV(0)
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case Else

                        ' AZIONDE DEL CASO: Azione non prevista esco con errore Parsing di protocollo
                        MyRefRet.Successfull = False
                        MyRefRet.Result = "KO"
                        Exit For

                End Select

            Next I

        Catch ex As Exception
            ' ESCO Con azione di parsing in Errore
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & ParseReturnString.ToString)
        End Try

    End Function

    Protected Shared Function getLocationString(ByRef actMethode As String) As String
        getLocationString = "CSVHelper" & "." & actMethode & " "
    End Function


End Class
