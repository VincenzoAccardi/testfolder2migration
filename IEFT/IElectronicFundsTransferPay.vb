#Region "Documentation"
' ********** ********** ********** **********
' IElectronicFundsTransferPay
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Retex SPA
' -----------------------------------
' Copyright by Retex SPA 2019
' -----------------------------------
#End Region
Public Interface IElectronicFundsTransferPay
    Function Pay(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode
    Function Check(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode
End Interface
