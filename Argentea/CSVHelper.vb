Imports System
Imports TPDotnet.Pos
Imports Microsoft.VisualBasic

Public Class CSVHelper

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

    Public Shared Function ParseReturnString(ByVal returnString As String, ByVal argenteaFunction As InternalArgenteaFunctionTypes, ByRef argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject, Optional ByRef szCharSeparator As String = ";") As Boolean
        ParseReturnString = False
        Dim I, J As Integer
        Dim funcName As String = "ParseReturnString"
        Dim CSV() As String = Nothing
        Dim StepNum As Integer
        Try
            LOG_Debug(getLocationString(funcName), "Parsing result")
            'ReDim argenteaFunctionReturnObject(0)
            argenteaFunctionReturnObject(0).Successfull = False
            argenteaFunctionReturnObject(0).ArgenteaFunction = argenteaFunction

            CSV = returnString.Split(szCharSeparator)
            For I = 0 To CSV.Length - 1

                Select Case argenteaFunction

                    Case InternalArgenteaFunctionTypes.EFTPayment, InternalArgenteaFunctionTypes.EFTVoid
                        argenteaFunctionReturnObject(0).TerminalID = CSV(0)
                        argenteaFunctionReturnObject(0).Amount = CSV(1)
                        If CSV(2) = "OK" Then
                            argenteaFunctionReturnObject(0).Successfull = True
                        Else
                            argenteaFunctionReturnObject(0).Successfull = False
                        End If
                        argenteaFunctionReturnObject(0).Description = CSV(3) ' when payment fails it describe the reason
                        argenteaFunctionReturnObject(0).Acquirer = CSV(4)
                        CSV(5) = Replace(CSV(5),
                                        Microsoft.VisualBasic.vbCrLf,
                                        Microsoft.VisualBasic.vbLf)
                        CSV(5) = Replace(CSV(5),
                                        Microsoft.VisualBasic.vbCr,
                                        Microsoft.VisualBasic.vbLf)
                        CSV(5) = Replace(CSV(5),
                                        Microsoft.VisualBasic.vbLf,
                                        Microsoft.VisualBasic.vbCrLf)
                        argenteaFunctionReturnObject(0).Receipt = CSV(5)
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case InternalArgenteaFunctionTypes.EFTGetTotals, InternalArgenteaFunctionTypes.EFTClose
                        If CInt(CSV(0)) > 1 Then ReDim Preserve argenteaFunctionReturnObject(CInt(CSV(0)) - 1)
                        For J = 0 To argenteaFunctionReturnObject.GetUpperBound(0)
                            If J <> 0 Then argenteaFunctionReturnObject(J) = New ArgenteaFunctionReturnObject
                            StepNum = 6 * J
                            argenteaFunctionReturnObject(J).TerminalID = CSV(StepNum + 1)
                            argenteaFunctionReturnObject(J).Abi = CSV(StepNum + 2)
                            argenteaFunctionReturnObject(J).Amount = CSV(StepNum + 3)
                            If CSV(StepNum + 4) = "OK" Then
                                argenteaFunctionReturnObject(J).Successfull = True
                            Else
                                argenteaFunctionReturnObject(J).Successfull = False
                            End If
                            argenteaFunctionReturnObject(J).Amount = CSV(StepNum + 3)
                            argenteaFunctionReturnObject(J).Description = CSV(StepNum + 5)
                            CSV(StepNum + 6) = Replace(CSV(StepNum + 6),
                                            Microsoft.VisualBasic.vbCrLf,
                                            Microsoft.VisualBasic.vbLf)
                            CSV(StepNum + 6) = Replace(CSV(StepNum + 6),
                                            Microsoft.VisualBasic.vbCr,
                                            Microsoft.VisualBasic.vbLf)
                            CSV(StepNum + 6) = Replace(CSV(StepNum + 6),
                                            Microsoft.VisualBasic.vbLf,
                                            Microsoft.VisualBasic.vbCrLf)
                            argenteaFunctionReturnObject(J).Receipt = CSV(StepNum + 6)
                        Next J
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case InternalArgenteaFunctionTypes.EFTConfirm
                        Exit Select
                    Case InternalArgenteaFunctionTypes.GiftCardActivationPreCheck,
                        InternalArgenteaFunctionTypes.GiftCardActivation,
                        InternalArgenteaFunctionTypes.GiftCardRedeemPreCkeck,
                        InternalArgenteaFunctionTypes.GiftCardRedeem,
                        InternalArgenteaFunctionTypes.GiftCardRedeemCancel,
                        InternalArgenteaFunctionTypes.GiftCardBalance

                        argenteaFunctionReturnObject(0).ArgenteaFunction = argenteaFunction
                        If CSV(0) = "OK" Then
                            argenteaFunctionReturnObject(0).Successfull = True
                        Else
                            argenteaFunctionReturnObject(0).Successfull = False
                        End If
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbCrLf,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbCr,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbLf,
                                            Microsoft.VisualBasic.vbCrLf)
                        argenteaFunctionReturnObject(0).Receipt = CSV(1)
                        argenteaFunctionReturnObject(0).Result = CSV(2)
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case InternalArgenteaFunctionTypes.PhoneRechargeCheck,
                        InternalArgenteaFunctionTypes.PhoneRechargeActivation

                        argenteaFunctionReturnObject(0).ArgenteaFunction = argenteaFunction
                        If CSV(0) = "OK" Then
                            argenteaFunctionReturnObject(0).Successfull = True
                        Else
                            argenteaFunctionReturnObject(0).Successfull = False
                        End If
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbCrLf,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbCr,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbLf,
                                            Microsoft.VisualBasic.vbCrLf)
                        argenteaFunctionReturnObject(0).Receipt = CSV(1)
                        argenteaFunctionReturnObject(0).Result = CSV(2)
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case InternalArgenteaFunctionTypes.ExternalGiftCardActivation,
                         InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation

                        argenteaFunctionReturnObject(0).ArgenteaFunction = argenteaFunction
                        If CSV(0) = "OK" Then
                            argenteaFunctionReturnObject(0).Successfull = True
                        Else
                            argenteaFunctionReturnObject(0).Successfull = False
                        End If
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbCrLf,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbCr,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(1) = Replace(CSV(1),
                                            Microsoft.VisualBasic.vbLf,
                                            Microsoft.VisualBasic.vbCrLf)
                        argenteaFunctionReturnObject(0).Receipt = CSV(1)
                        argenteaFunctionReturnObject(0).Result = CSV(2)
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case InternalArgenteaFunctionTypes.ADVPayment
                        argenteaFunctionReturnObject(0).ArgenteaFunction = argenteaFunction
                        If CSV(0) = "OK" Then
                            argenteaFunctionReturnObject(0).Successfull = True
                        Else
                            argenteaFunctionReturnObject(0).Successfull = False
                        End If
                        argenteaFunctionReturnObject(0).Description = CSV(1)
                        argenteaFunctionReturnObject(0).TerminalID = CSV(2)
                        If Not String.IsNullOrEmpty(CSV(3).Trim()) Then
                            argenteaFunctionReturnObject(0).TerminalID = CSV(3)
                        End If
                        argenteaFunctionReturnObject(0).Amount = CSV(4)
                        CSV(5) = Replace(CSV(5),
                                            Microsoft.VisualBasic.vbCrLf,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(5) = Replace(CSV(5),
                                            Microsoft.VisualBasic.vbCr,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(5) = Replace(CSV(5),
                                            Microsoft.VisualBasic.vbLf,
                                            Microsoft.VisualBasic.vbCrLf)
                        argenteaFunctionReturnObject(0).Receipt = CSV(5)
                        argenteaFunctionReturnObject(0).Result = CSV(0)
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case InternalArgenteaFunctionTypes.ADVVoid
                        argenteaFunctionReturnObject(0).ArgenteaFunction = argenteaFunction
                        If CSV(0) = "OK" Then
                            argenteaFunctionReturnObject(0).Successfull = True
                        Else
                            argenteaFunctionReturnObject(0).Successfull = False
                        End If
                        argenteaFunctionReturnObject(0).Description = CSV(1)
                        argenteaFunctionReturnObject(0).TerminalID = CSV(2)

                        argenteaFunctionReturnObject(0).Amount = CSV(3)
                        CSV(4) = Replace(CSV(4),
                                            Microsoft.VisualBasic.vbCrLf,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(4) = Replace(CSV(4),
                                            Microsoft.VisualBasic.vbCr,
                                            Microsoft.VisualBasic.vbLf)
                        CSV(4) = Replace(CSV(4),
                                            Microsoft.VisualBasic.vbLf,
                                            Microsoft.VisualBasic.vbCrLf)
                        argenteaFunctionReturnObject(0).Receipt = CSV(4)
                        argenteaFunctionReturnObject(0).Result = CSV(0)
                        ParseReturnString = True
                        Exit For
                        Exit Select
                    Case Else
                        argenteaFunctionReturnObject(0).Successfull = False
                        argenteaFunctionReturnObject(0).Result = "KO"
                        Exit For

                End Select

            Next I

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & ParseReturnString.ToString)
        End Try

    End Function

    Protected Shared Function getLocationString(ByRef actMethode As String) As String
        getLocationString = "CSVHelper" & "." & actMethode & " "
    End Function


End Class
