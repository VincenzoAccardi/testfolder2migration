#Region "Documentation"
' ********** ********** ********** **********
' IGiftCardBalanceInquiry
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region
#Region "IGiftCardBalanceInquiry"
Public Interface IGiftCardBalanceInquiry
    Function GiftCardBalanceInquiry(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode
End Interface
#End Region
