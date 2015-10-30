#Region "Documentation"
' ********** ********** ********** **********
' IFiscalTA
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Assago, 2012, All rights reserved.
' -----------------------------------
#End Region
#Region "IFiscalTA"
Public Interface IFiscalTA
    Property bFiscalPrinted() As Boolean
    Function GetPayedValueForMediaMember(ByVal lMediaMember As Integer) As Decimal
    Function GetPaidValueOfMediaMemberToCheckCashHalo() As Decimal
End Interface
#End Region
