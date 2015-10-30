Imports System
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

Public Class ModSignOff
    Inherits TPDotnet.Pos.ModSignOff
    Public Overrides Function ModBase_run(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Short
        Dim funcName As String = "ModBase_run"
        LOG_Debug(getLocationString(funcName), "We are entered in Argentea Signoff overrride ModBase_run function")
        Try
            ModBase_run = MyBase.ModBase_run(taobj, TheModCntr)
            If ModBase_run = 1 Then
                If TheModCntr.ModulNmbrExt = 0 Then
                    'Try
                    '    EFTController.Instance.GetTotals(taobj, TheModCntr)
                    'Catch ex As Exception
                    '    LOG_Debug(getLocationString(funcName), "Argentea GetTotals function raises an error: " & ex.Message)
                    'End Try
                    Try
                        EFTController.Instance.Close(taobj, TheModCntr)
                    Catch ex As Exception
                        LOG_Debug(getLocationString(funcName), "Argentea Close function raises an error: " & ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception
            LOG_Debug(getLocationString(funcName), "Short Signoff overrride ModBase_run function raises an error: " & ex.Message)
        End Try
    End Function
End Class
