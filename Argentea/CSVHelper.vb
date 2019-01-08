Imports System
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
    Public Shared Function ParseReturnString(ByVal returnString As String,
                                             ByVal argenteaFunction As InternalArgenteaFunctionTypes,
                                             ByRef argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject,
                                             Optional ByVal szCharSeparator As String = ";",
                                             Optional ByVal iFractParser As Integer = 1) As Boolean
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

        Dim SetTypeSuccessufully As Func(Of String, Boolean) = Function(ByVal strFieldCSV As String) 'As Boolean
                                                                   If strFieldCSV = "000" Then Return True Else Return False
                                                               End Function

        ' Per i valori attesi numerici  se stringa vuota riporta "0"
        Dim SetNumeric As Func(Of String, String) = Function(ByVal strFieldCSV As String) 'As String
                                                        If strFieldCSV = String.Empty Then Return "0" Else Return strFieldCSV
                                                    End Function

        ' Per i valori attesi numerici  se stringa vuota riporta False o se "0" Riporta False o se "1" Riporta True
        Dim SetBoolState As Func(Of String, String) = Function(ByVal strFieldCSV As String) 'As String
                                                          If strFieldCSV = String.Empty OrElse strFieldCSV.Trim() = "0" Then Return False Else Return True
                                                      End Function


        Dim SetErrResponse As Func(Of String, EFT.ArgenteaFunctionReturnObject) = Function(msg As String) ' As ArgenteaFunctionReturnObject
                                                                                      Dim CSVE(200) As String
                                                                                      CSVE = returnString.Split("-") 'szCharSeparator)
                                                                                      Dim MyRefErr As New ArgenteaFunctionReturnObject
                                                                                      '"KO-903-PROGRESSIVO FUORI SEQUENZA-----0---"  ' <-- x test (su ko remoti)
                                                                                      MyRefErr.ArgenteaFunction = argenteaFunction
                                                                                      MyRefErr.Successfull = False
                                                                                      MyRefErr.CodeResult = CSVE(1)
                                                                                      MyRefErr.Description = CSVE(2)
                                                                                      MyRefErr.Receipt = msg
                                                                                      MyRefErr.Amount = "0"
                                                                                      MyRefErr.Provider = ""
                                                                                      MyRefErr.TerminalID = ""
                                                                                      MyRefErr.RequireCommit = False
                                                                                      MyRefErr.CodeIssuer = ""
                                                                                      MyRefErr.Result = CSVE(0)
                                                                                      Return MyRefErr
                                                                                  End Function

        Dim GetResultAndDictBPs As Func(Of String,
            Tuple(Of Decimal, Integer, Collections.Generic.Dictionary(Of String, Decimal))) =
            Function(msg As String) As Tuple(Of Decimal, Integer, Collections.Generic.Dictionary(Of String, Decimal))
                Dim _Partial As Decimal, _NumB As Integer
                Dim _DictBPs As New Collections.Generic.Dictionary(Of String, Decimal)

                Dim Itms(200) As String
                ReDim Itms(CInt(SetNumeric(CSV(2).Split("|")(0))))
                Itms = CSV(2).Split("|")
                ' il 3 è il numero di BP evasi in Tagli
                _NumB = 0
                For X As Integer = 1 To (CInt(SetNumeric(Itms(0))) + 1) Step 2
                    If Not X = CInt(SetNumeric(Itms(0))) Then
                        For Y As Integer = 0 To CInt(SetNumeric(Itms(X)) - 1)
                            _DictBPs.Add("terminal_bp_" + CStr(_NumB + 1), CDec(SetNumeric(Itms(X + 1)) / iFractParser))
                            _Partial += CDec(SetNumeric(Itms(X + 1)))
                            _NumB += 1
                        Next
                    End If
                    '_Partial = _Partial + (CInt(Itms(X)) * CInt(Itms(X + 1)))
                Next
                Return Tuple.Create(Of Decimal, Integer, Collections.Generic.Dictionary(Of String, Decimal))(
                    _Partial,
                    _NumB,
                    _DictBPs
                )

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

                    Case InternalArgenteaFunctionTypes.Initialization_AG
                        '"OK--TICKET APERTO-----0---"    ' <-- x test 
                        '"KO-903-PROGRESSIVO FUORI SEQUENZA-----0---"  ' <-- x test (su ko remoti)
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.CodeResult = "0"
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Receipt = ""
                            MyRefRet.Amount = "0"
                            MyRefRet.Provider = ""
                            MyRefRet.TerminalID = ""
                            MyRefRet.RequireCommit = False
                            MyRefRet.CodeIssuer = ""
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.ResetCounter_AG
                        '"OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--"    ' <-- x test 
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.CodeResult = SetNumeric(CSV(1))
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(3))
                            MyRefRet.Amount = CDec(SetNumeric(CSV(4))) / iFractParser
                            MyRefRet.Provider = CSV(5)
                            MyRefRet.TerminalID = CSV(6)
                            MyRefRet.RequireCommit = SetBoolState(CSV(7))
                            MyRefRet.CodeIssuer = CSV(8)
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.Confirmation_AG
                        '"OK-0 -CONFERMATO CON SUCCESSO-----0-2--"    ' <-- x test 
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.CodeResult = SetNumeric(CSV(1))
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(3))
                            MyRefRet.Amount = CDec(SetNumeric(CSV(4))) / iFractParser
                            MyRefRet.Provider = CSV(5)
                            MyRefRet.TerminalID = CSV(6)
                            MyRefRet.RequireCommit = False
                            MyRefRet.CodeIssuer = CSV(8)
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.Close_AG
                        '"OK-0 -OPERAZIONE COMPLETATA---ARGENTEA--0---"    ' <-- x test 
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.CodeResult = SetNumeric(CSV(1))
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(3))
                            MyRefRet.Amount = CDec(SetNumeric(CSV(4))) / iFractParser
                            MyRefRet.Provider = CSV(5)
                            MyRefRet.TerminalID = CSV(6)
                            MyRefRet.RequireCommit = False
                            MyRefRet.CodeIssuer = CSV(8)
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.Check_BP
                        '  0      1   2           3       4    5         6             7
                        '"Buono-000-Buono Valido-529-8897456-12345687-201809201733577-ARGENTEA-"            ' <-- x test su questo signal
                        MyRefRet.Successfull = SetTypeSuccessufully(CSV(1))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.Type = SetNumeric(CSV(0))
                            MyRefRet.CodeResult = SetNumeric(CSV(1))
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Amount = CDec(SetNumeric(CSV(3))) / iFractParser
                            MyRefRet.Receipt = CSV(4)
                            MyRefRet.CodeIssuer = CSV(5)
                            MyRefRet.TerminalID = CSV(6)
                            MyRefRet.Provider = CSV(7)
                            MyRefRet.RequireCommit = "0"
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.SinglePaid_BP
                        '"OK-0 - BUONO VALIDATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--"    ' <-- x test 
                        ' AZIONDE DEL CASO: Buoni Pasto Cartacei da protocollo in risposta dal SERVICE remoto Argentea
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.CodeResult = SetNumeric(CSV(1))
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(3))
                            MyRefRet.Amount = CDec(SetNumeric(CSV(4))) / iFractParser
                            MyRefRet.Provider = CSV(5)
                            MyRefRet.TerminalID = CSV(6)
                            MyRefRet.RequireCommit = SetBoolState(CSV(7))
                            MyRefRet.CodeIssuer = CSV(8)
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.MultiPaid_BP
                        '"OK;TRANSAZIONE ACCETTATA;2|5|10|1|4;104;PELLEGRINI;  PAGAMENTO BUONO PASTO "
                        ' AZIONE DEL CASO: Buoni Pasto Elettronici  da protocollo in risposta dal POS locale fornito da Argentea
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            Dim Collect As Tuple(Of Decimal, Integer, Collections.Generic.Dictionary(Of String, Decimal)) = GetResultAndDictBPs("Res")
                            MyRefRet.ListBPsEvaluated = Collect.Item3
                            MyRefRet.Amount = Collect.Item1 / iFractParser
                            MyRefRet.NumBPEvalutated = Collect.Item2
                            MyRefRet.TerminalID = "POS"
                            MyRefRet.RequireCommit = False
                            MyRefRet.CodeIssuer = CSV(J + 3)
                            MyRefRet.NameIssuer = CSV(J + 4)
                            MyRefRet.Provider = "ARGENTEA"
                            MyRefRet.Description = MyRefRet.Description '& " - " & CSV(J + 5)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(5))
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.MultiItemsIC_BP
                        '"OK;OPERAZIONE ACCETTATA;2|5|10|1|4;104;PELLEGRINI;  INFO BUONI PASTO "
                        ' AZIONE DEL CASO: Buoni Pasto Elettronici  da protocollo in risposta dal POS locale fornito da Argentea per le info sulla Card
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            Dim Collect As Tuple(Of Decimal, Integer, Collections.Generic.Dictionary(Of String, Decimal)) = GetResultAndDictBPs("Res")
                            MyRefRet.ListBPsEvaluated = Collect.Item3
                            MyRefRet.Amount = Collect.Item1 / iFractParser
                            MyRefRet.NumBPEvalutated = Collect.Item2
                            MyRefRet.TerminalID = "POS"
                            MyRefRet.RequireCommit = False
                            MyRefRet.CodeIssuer = CSV(J + 3)
                            MyRefRet.NameIssuer = CSV(J + 4)
                            MyRefRet.Provider = "ARGENTEA"
                            MyRefRet.Description = MyRefRet.Description '& " - " & CSV(J + 5)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(5))
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.MultiVoid_BP
                        '"OK;TRANSAZIONE ACCETTATA;2|5|10|1|4;104;PELLEGRINI;  STORNO BUONO PASTO "
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction

                        If MyRefRet.Successfull Then
                            Dim Collect As Tuple(Of Decimal, Integer, Collections.Generic.Dictionary(Of String, Decimal)) = GetResultAndDictBPs("Res")
                            MyRefRet.ListBPsEvaluated = Collect.Item3
                            MyRefRet.Amount = Collect.Item1 / iFractParser
                            MyRefRet.NumBPEvalutated = Collect.Item2
                            MyRefRet.TerminalID = "POS"
                            MyRefRet.RequireCommit = False
                            MyRefRet.CodeIssuer = CSV(J + 3)
                            MyRefRet.NameIssuer = CSV(J + 4)
                            MyRefRet.Provider = "ARGENTEA"
                            MyRefRet.Description = MyRefRet.Description '& " - " & CSV(J + 5)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(5))
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
                        ParseReturnString = True
                        Exit For
                        Exit Select

                    Case InternalArgenteaFunctionTypes.SingleVoid_BP
                        '"OK-0 - BUONO STORNATO CON SUCCESSO-68195717306007272725069219400700-700-ARGENTEA-201809181448517-0-202--"
                        MyRefRet.Successfull = SetSuccessufully(CSV(0))
                        MyRefRet.ArgenteaFunction = argenteaFunction
                        If MyRefRet.Successfull Then
                            MyRefRet.Successfull = SetSuccessufully(CSV(0))
                            MyRefRet.CodeResult = SetNumeric(CSV(1))
                            MyRefRet.Description = CSV(2)
                            MyRefRet.Receipt = ReplaceVbCRLF(CSV(3))
                            MyRefRet.Amount = CDec(SetNumeric(CSV(4))) / iFractParser
                            MyRefRet.Provider = CSV(5)
                            MyRefRet.TerminalID = CSV(6)
                            MyRefRet.RequireCommit = SetBoolState(CSV(7))
                            MyRefRet.CodeIssuer = CSV(8)
                            MyRefRet.Result = CSV(0)
                        Else
                            argenteaFunctionReturnObject(I) = SetErrResponse("err argentea")
                        End If
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
