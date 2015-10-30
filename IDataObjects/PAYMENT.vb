#Region "Documentation"
' ********** ********** ********** **********
' IFiscalPAYMENT
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Assago, 2012, All rights reserved.
' -----------------------------------
#End Region
#Region "IFiscalPAYMENT"
Public Interface IFiscalPAYMENT
    Property bITCheckCashHalo() As Integer
    Property dITTxHALO() As Decimal
    Property bITFiscalNotPaid() As Integer
End Interface
#End Region
