Option Strict Off
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos

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

Public Class FormEFTFunctions
    Inherits FormBase
    Public Enum _Options As Short
        No_Option_Selected = -1
        Void_Last_Transaction = 0
        Close_EFT = 1
        Get_Totals = 2
    End Enum
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Tag = False
        Me.bDialogActive = False
    End Sub
    Public ReadOnly Property OptionSelected As _Options
        Get
            Dim vReturn As _Options = _Options.No_Option_Selected
            If rbVoidLT.Checked Then vReturn = _Options.Void_Last_Transaction
            If rbEFTClose.Checked Then vReturn = _Options.Close_EFT
            If rbGetTotals.Checked Then vReturn = _Options.Get_Totals
            Return vReturn
        End Get
    End Property
    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Me.Tag = True
        Me.bDialogActive = False
    End Sub
    Private Sub FormEFTHandler_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim GetLastPaymentTransactionIdentifier As String = RegistryHelper.GetLastPaymentTransactionIdentifier
        Dim LastPaymentTransactionIdentifier(0) As String
        Try
            LastPaymentTransactionIdentifier = Split(GetLastPaymentTransactionIdentifier, "|")

            lblAmountToken.Text = Format(RegistryHelper.GetLastPaymentTransactionAmount / 100, "####0.00")
            lblNumberToken.Text = CInt(LastPaymentTransactionIdentifier(1)) & "|" & _
                                    CInt(LastPaymentTransactionIdentifier(3)) & "|" & _
                                    CInt(LastPaymentTransactionIdentifier(4))
        Catch ex As Exception

        End Try
        
    End Sub
End Class