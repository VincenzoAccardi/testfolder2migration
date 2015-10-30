#Region "Documentation"
' ********** ********** ********** **********
' ITaZReportRec
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Assago, 2012, All rights reserved.
' -----------------------------------
#End Region
#Region "ITaZReportRec"
Public Interface ITaZReportRec
    Property dFiscalDailyTotal() As Double
    Property dFiscalGrandTotal() As Double
    Property lZReportNmbr() As Integer
    Property lFiscalReceiptNmbr() As Integer
    Property lNonFiscalReceiptNmbr() As Integer
    Property szFiscalPrinterID() As String
    Property szFiscalPrinterFW() As String
End Interface
#End Region
