﻿#Region "Documentation"
' ********** ********** ********** **********
' IGiftCardCancellationPayment
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region
#Region "IGiftCardCancellationPayment"
Public Interface IGiftCardCancellationPayment
    Function GiftCardCancellation(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IGiftCardReturnCode
End Interface
#End Region
