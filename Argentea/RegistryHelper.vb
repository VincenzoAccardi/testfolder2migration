Imports Microsoft.Win32
Imports System

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

Public Class RegistryHelper

    Public Shared Sub SetLastPaymentTransactionIdentifier(ByVal LastPaymentTransactionIdentifier As String)
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            REG.SetValue("LastPaymentTransactionIdentifier", LastPaymentTransactionIdentifier.ToString)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Sub

    Public Shared Function GetLastPaymentTransactionIdentifier() As String
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            GetLastPaymentTransactionIdentifier = REG.GetValue("LastPaymentTransactionIdentifier", 0)
        Catch ex As Exception
            GetLastPaymentTransactionIdentifier = String.Empty
        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function

    Public Shared Sub SetLastPaymentTransactionAmount(ByVal LastPaymentTransactionAmount As Integer)
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            REG.SetValue("LastPaymentTransactionAmount", LastPaymentTransactionAmount.ToString)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Sub

    Public Shared Function GetLastPaymentTransactionAmount() As Integer
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            GetLastPaymentTransactionAmount = REG.GetValue("LastPaymentTransactionAmount", 0)
        Catch ex As Exception
            GetLastPaymentTransactionAmount = -1
        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function

End Class
