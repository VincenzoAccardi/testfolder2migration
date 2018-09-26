#Region "Documentation"
' ********** ********** ********** **********
' IBPCDematerialize
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region

#Region "IBPCDematerialize"
Public Interface IBPCDematerialize
    Function Dematerialize(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IBPReturnCode
End Interface
#End Region