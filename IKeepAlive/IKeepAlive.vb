#Region "Documentation"
' ********** ********** ********** **********
' IKeepAlive
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region
#Region "IKeepAlive"
Public Interface IKeepAlive
    Sub Run(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA)
End Interface
#End Region
