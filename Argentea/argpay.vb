#Region "Documentation"
' ********** ********** ********** **********
' Internal dummy class to simulate the answer by Argentea
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region

Namespace PAGAMENTOLib_TESTOFFLINE

    Public Class argpay

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="nOperazione">not used</param>
        ''' <param name="importo">payment amount in euro cents</param>
        ''' <param name="tmo">card swipe timeout</param>
        ''' <param name="transaction_identifier">transaction identifier containing: 
        '''     - Codifica tracciato – Char – 5 bytes;
        '''     - Matricola Fiscale della cassa – Char – 10 bytes;
        '''     - Codice Cassiere – Char – 8 bytes;
        '''     - Data dell'operazione – Char – 8 bytes (AAAAMMGG);
        '''     - Ora dell'operazione – Char – 6 bytes (HHMMSS);
        '''     - Numero scontrino fiscale – Char – 5 bytes (“0” riempitivi);
        ''' </param>
        ''' <param name="cMsg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function PagamentoPlus(ByVal nOperazione As Integer, _
                                        ByVal importo As Integer, _
                                        ByVal tmo As Integer, _
                                        ByRef transaction_identifier As String, _
                                        ByRef cMsg As String) As Integer
            'Dim newgr As New PAGAMENTOLib.argpay
            'Dim pippo = newgr.PagamentoPlus(nOperazione, importo, tmo, transaction_identifier, cMsg)
            PagamentoPlus = 1
            System.Threading.Thread.Sleep(15000)
            cMsg = "12345678" + _
                    ";" + importo.ToString("000000000000") + _
                    ";" + "OK" + _
                    ";" + "Transazione eseguita" + _
                    ";" + "00" + _
                    ";" + "123456789012345678901234" + Microsoft.VisualBasic.vbCrLf + _
                        "123456789012345678901234" + Microsoft.VisualBasic.vbCrLf + _
                        "123456789012345678901234" + Microsoft.VisualBasic.vbCrLf + _
                        "123456789012345678901234" + Microsoft.VisualBasic.vbCrLf + _
                        "123456789012345678901234" + Microsoft.VisualBasic.vbCrLf

        End Function

        Public Function StornoPlus(ByVal nOperazione As Integer, _
                                        ByVal importo As Integer, _
                                        ByVal tmo As Integer, _
                                        ByRef transaction_identifier As String, _
                                        ByRef cMsg As String) As Integer

            StornoPlus = 1
            System.Threading.Thread.Sleep(15000)
            cMsg = " 12345602;000000000001;OK;;; " & _
                    "      TERMINALE TEST      " & Microsoft.VisualBasic.vbCrLf & _
                    "         ARGENTEA         " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    "     STORNO ACQUISTO      " & Microsoft.VisualBasic.vbCrLf & _
                    "           VISA           " & Microsoft.VisualBasic.vbCrLf & _
                    " DATA 08/05/14  ORA 18:01 " & Microsoft.VisualBasic.vbCrLf & _
                    " ESERC.        8888888888 " & Microsoft.VisualBasic.vbCrLf & _
                    " ACQ.ID       88105000003 " & Microsoft.VisualBasic.vbCrLf & _
                    " N.OP.000029 TML 12345602 " & Microsoft.VisualBasic.vbCrLf & _
                    " CAUSALE 400     A.C. 000 " & Microsoft.VisualBasic.vbCrLf & _
                    " PAN     402360******1123 " & Microsoft.VisualBasic.vbCrLf & _
                    " EXP                 **** " & Microsoft.VisualBasic.vbCrLf & _
                    " STAN 000012  AUT. 66258  " & Microsoft.VisualBasic.vbCrLf & _
                    " IMPORTO  EUR        3,38 " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    " STORNO  CARTE DI CREDITO " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    "   ARRIVEDERCI E GRAZIE   "

        End Function

        Public Function Conferma(ByVal nOperazione As Integer) As Integer

            Conferma = 1

        End Function

        Public Function RichiestaTotaliHost(ByVal nOperazione As Integer, _
                                            ByVal ChiusuraSessione As Integer,
                                            ByRef szOut As String) As Integer

            RichiestaTotaliHost = 1
            System.Threading.Thread.Sleep(15000)
            szOut = " 01;12345602;;0000000000000003042;OK;; " & _
                    "      TERMINALE TEST      " & Microsoft.VisualBasic.vbCrLf & _
                    "         ARGENTEA         " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    "       TOTALI HOST        " & Microsoft.VisualBasic.vbCrLf & _
                    " DATA 08/05/14  ORA 18:03 " & Microsoft.VisualBasic.vbCrLf & _
                    " STAN 000013 TML 12345602 " & Microsoft.VisualBasic.vbCrLf & _
                    " CAUSALE 870              " & Microsoft.VisualBasic.vbCrLf & _
                    " OPER.N. 000030           " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    " Mastercard               " & Microsoft.VisualBasic.vbCrLf & _
                    " ------------------------ " & Microsoft.VisualBasic.vbCrLf & _
                    " TOTALE EUR          0,00 " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    " Maestro                  " & Microsoft.VisualBasic.vbCrLf & _
                    " ------------------------ " & Microsoft.VisualBasic.vbCrLf & _
                    " TOTALE EUR          0,00 " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    " VISA                     " & Microsoft.VisualBasic.vbCrLf & _
                    " ACQUISTI           33,80 " & Microsoft.VisualBasic.vbCrLf & _
                    " (num. acq. 10)           " & Microsoft.VisualBasic.vbCrLf & _
                    " STORNI              3,38 " & Microsoft.VisualBasic.vbCrLf & _
                    " (num. storni 1)          " & Microsoft.VisualBasic.vbCrLf & _
                    " ------------------------ " & Microsoft.VisualBasic.vbCrLf & _
                    " TOTALE EUR         30,42 " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    " TOTALE GENERALE          " & Microsoft.VisualBasic.vbCrLf & _
                    " EUR                30,42 " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    "   TRANSAZIONE ESEGUITA   "

        End Function

        Public Function Chiusura(ByVal nOperazione As Integer, _
                                    ByRef szOut As String) As Integer

            Chiusura = 1
            System.Threading.Thread.Sleep(15000)
            szOut = " 01;12345602;;0000000000000003042;OK;; " & _
                    "      TERMINALE TEST      " & Microsoft.VisualBasic.vbCrLf & _
                    "         ARGENTEA         " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    "         CHIUSURA         " & Microsoft.VisualBasic.vbCrLf & _
                    " DATA 08/05/14  ORA 18:04 " & Microsoft.VisualBasic.vbCrLf & _
                    " STAN 000014 TML 12345602 " & Microsoft.VisualBasic.vbCrLf & _
                    " TRANSAZIONI N.        11 " & Microsoft.VisualBasic.vbCrLf & _
                    " CAUSALE 871              " & Microsoft.VisualBasic.vbCrLf & _
                    " OPER.N. 000031           " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    " TOTALE GENERALE          " & Microsoft.VisualBasic.vbCrLf & _
                    " EUR                30,42 " & Microsoft.VisualBasic.vbCrLf & _
                    "                          " & Microsoft.VisualBasic.vbCrLf & _
                    "   TRANSAZIONE ESEGUITA   "

        End Function

    End Class

End Namespace
