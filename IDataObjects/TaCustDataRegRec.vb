#Region "Documentation"
' ********** ********** ********** **********
' ITaCustDataRegRec
' ---------- ---------- ---------- ----------
' Author : Emanuele Gualtierotti
' Wincor Nixdorf Retail Consulting
' -----------------------------------
' Copyright by Wincor Nixdorf Retail Consulting
' 20090, Assago, 2012, All rights reserved.
' -----------------------------------
#End Region
#Region "ITaCustDataRegRec"
Public Interface ITaCustDataRegRec
    Property szFirstName() As String
    Property szLastName() As String
    Property szBirthDate() As String
    Property szBirthCity() As String
    Property szAddress() As String
    Property szState() As String
End Interface
#End Region

