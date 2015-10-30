Public Class ArgenteaFunctionReturnObject

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

#Region "CVS"
    Public TerminalID As String = ""
    Public Amount As String = ""
    Public Result As String = ""
    Public Description As String = ""
    Public Acquirer As String = ""
    Public Receipt As String = ""
    Public Abi As String = ""
#End Region

#Region "CVS fields helper"



#End Region

    Public ArgenteaFunction As InternalArgenteaFunctionTypes
    Public Successfull As Boolean

End Class
