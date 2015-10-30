#Region "Documentation"
' ********** ********** ********** **********
' IPhoneRechargeActivation
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region
#Region "IPhoneRechargeActivation"
Public Interface IPhoneRechargeActivation
    Function ActivatePhoneRecharge(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IPhoneRechargeReturnCode
End Interface
#End Region
