#Region "Documentation"
' ********** ********** ********** **********
' IPhoneRechargeActivationPreCheck
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region
#Region "IPhoneRechargeActivationPreCheck"
Public Interface IPhoneRechargeActivationPreCheck
    Function CheckPhoneRecharge(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IPhoneRechargeReturnCode
End Interface
#End Region
