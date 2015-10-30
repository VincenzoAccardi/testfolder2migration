Imports System
Imports TPDotnet.Pos
Imports System.IO

#Region "Documentation"
' ********** ********** ********** **********
' Argentea EFT
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Basiglio, 2014, All rights reserved.
' -----------------------------------
#End Region

Public Interface ITPTAMediaSwapHelper

    Function GetMediaMemberByCardType(ByRef CardType As String) As Integer
    Sub SwapElectronicMedia(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef TheMediaRec As TPDotnet.Pos.TaMediaRec, ByRef CardType As String)

End Interface
