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
#Region "EFT"
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

#End Region

#Region "ADV"
    Public Shared Sub SetLastPaymentADVTransactionIdentifier(ByVal LastPaymentADVTransactionIdentifier As String)
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            REG.SetValue("LastPaymentADVTransactionIdentifier", LastPaymentADVTransactionIdentifier.ToString)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Sub

    Public Shared Function GetLastPaymentADVTransactionIdentifier() As String
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            GetLastPaymentADVTransactionIdentifier = REG.GetValue("LastPaymentADVTransactionIdentifier", 0)
        Catch ex As Exception
            GetLastPaymentADVTransactionIdentifier = String.Empty
        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function

    Public Shared Sub SetLastPaymentADVTransactionAmount(ByVal LastPaymentADVTransactionAmount As Integer)
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            REG.SetValue("LastPaymentADVTransactionAmount", LastPaymentADVTransactionAmount.ToString)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Sub


    Public Shared Function GetLastPaymentADVTransactionAmount() As Integer
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            GetLastPaymentADVTransactionAmount = REG.GetValue("LastPaymentADVTransactionAmount", 0)
        Catch ex As Exception
            GetLastPaymentADVTransactionAmount = -1
        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function
    Public Shared Sub SetLastPaymentADVTransactionType(ByVal LastPaymentADVTransactionType As Integer)
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            REG.SetValue("LastPaymentADVTransactionType", LastPaymentADVTransactionType.ToString)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Sub
    Public Shared Function GetLastPaymentADVTransactionType() As Integer
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            GetLastPaymentADVTransactionType = REG.GetValue("LastPaymentADVTransactionType", 0)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function
#End Region
End Class
