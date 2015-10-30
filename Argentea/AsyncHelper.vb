Imports System
Imports TPDotnet
Imports TPDotnet.Pos

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

'operation 1 pagamento ==> PagamentoPlus
Public Delegate Function PagamentoPlusDelegate(ByVal nOperazione As Integer, _
                                        ByVal importo As Integer, _
                                        ByVal tmo As Integer, _
                                        ByRef transaction_identifier As String, _
                                        ByRef cMsg As String) As Integer

'operation 2 storno ==> StornoPlus
Public Delegate Function StornoPlusDelegate(ByVal nOperazione As Integer, _
                                        ByVal importo As Integer, _
                                        ByVal tmo As Integer, _
                                        ByRef transaction_identifier As String, _
                                        ByRef cMsg As String) As Integer

'operation 3 richiesta totali ==> RichiestaTotaliHost
Public Delegate Function RichiestaTotaliHostDelegate(ByVal nOperazione As Integer, _
                                        ByVal ChiusuraSessione As Integer, _
                                        ByRef szOut As String) As Integer


'operation 4 chiusura ==> RichiestaTotaliHost
Public Delegate Function ChiusuraDelegate(ByVal nOperazione As Integer, _
                                            ByRef szOut As String) As Integer


'confirmation return 
Public Delegate Function ConfermaDelegate(ByVal nOperazione As Integer) As Integer



Public Class AsyncHelper

    Public Shared Sub Callback(ByVal ar As System.IAsyncResult)

    End Sub

    Public Delegate Function InvokerDelegate(ByRef parameters As System.Collections.Hashtable) As Boolean

    Public Function Invoker(ByRef parameters As System.Collections.Hashtable) As Boolean
        Invoker = False
        Dim funcName As String = parameters("FuncName").ToString
        Dim del As InvokerDelegate = parameters("Function")

        Try
            LOG_FuncStart(funcName)
            Invoker = del.Invoke(parameters)

        Catch ex As Exception
            LOG_Error(funcName, ex.Message)
        Finally
            LOG_FuncExit(funcName, Invoker.ToString)
        End Try

    End Function

End Class
