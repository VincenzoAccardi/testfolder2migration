Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports System.Windows.Forms

Imports Microsoft.PointOfService
Imports Microsoft.PointOfService.PosCommon
Imports System.Collections.Generic


Public Class ModArgentea
    Inherits TPDotnet.Pos.ModBase

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

#Region "Overridable"

    Protected myForm As FormEFTFunctions = Nothing
    Protected LastOkAmount As Integer = 0
    Protected LastOkNumber As Integer = 0
    Protected NextNumber As Integer = 0

    Public Enum FunctionType
        EFT
        GiftCard
    End Enum

    Public Overrides Function ModBase_run(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Short
        Dim szFuncName As String = "ModBase_run"
        Dim frm As FormBase = Nothing

        Try
            ModBase_run = 1
            LOG_FuncStart(getLocationString(szFuncName))

            Select Case TheModCntr.ModulNmbrExt

                Case FunctionType.EFT
                    frm = TheModCntr.GetCustomizedForm(GetTPDotNetType(Of FormEFTFunctions), STRETCH_TO_SMALL_WINDOW)
                    Exit Select

                Case FunctionType.GiftCard
                    frm = TheModCntr.GetCustomizedForm(GetTPDotNetType(Of FormArgenteaItemInput), STRETCH_TO_SMALL_WINDOW)
                    Exit Select

                Case Else
                    LOG_Error(getLocationString(szFuncName), "Invalid option: " + TheModCntr.ModulNmbrExt.ToString)
                    Exit Function

            End Select

            'If TheModCntr.ModulNmbrExt = FunctionType.GiftCard Then
            '    frm = TheModCntr.GetCustomizedForm(GetTPDotNetType(Of FormArgenteaItemInput), STRETCH_TO_SMALL_WINDOW)
            'Else
            '    frm = TheModCntr.GetCustomizedForm(GetTPDotNetType(Of FormEFTFunctions), STRETCH_TO_SMALL_WINDOW)
            'End If

            frm.theModCntr = TheModCntr
            frm.taobj = taobj
            Dim bFrmRet As Object = frm.DisplayMe(taobj, TheModCntr, FormRoot.DialogActive.d1_DialogActive)
            frm.Close()

            If TypeOf frm Is FormArgenteaItemInput Then

                Dim dBalance As Decimal = 0
                Dim szReceipt As String = String.Empty

                InquiryGiftCard(TheModCntr, taobj, frm.Tag, dBalance, szReceipt)

            ElseIf TypeOf frm Is FormEFTFunctions Then

                If Not bFrmRet Then Exit Function

                Select Case CType(frm, FormEFTFunctions).OptionSelected

                    Case FormEFTFunctions._Options.No_Option_Selected
                        TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION, getPosTxtNew(TheModCntr.contxt, "LevelITCommonHelloWorldFailed", 0), 0, TheModCntr, "LevelITCommonHelloWorldFailed")

                        'Case FormEFTFunctions._Options.Void_Last_Transaction
                        '    If Not TPDotnet.IT.Common.Pos.EFT.EFTController.Instance.Void(taobj, TheModCntr) Then
                        '        TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION, getPosTxtNew(TheModCntr.contxt, "LevelITCommonHelloWorldFailed", 0), 0, TheModCntr, "LevelITCommonHelloWorldFailed")
                        '    End If

                        'Case FormEFTFunctions._Options.Close_EFT
                        '    If Not TPDotnet.IT.Common.Pos.EFT.EFTController.Instance.Close(taobj, TheModCntr) Then
                        '        TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION, getPosTxtNew(TheModCntr.contxt, "LevelITCommonHelloWorldFailed", 0), 0, TheModCntr, "LevelITCommonHelloWorldFailed")
                        '    End If

                        'Case FormEFTFunctions._Options.Get_Totals
                        '    If Not TPDotnet.IT.Common.Pos.EFT.EFTController.Instance.GetTotals(taobj, TheModCntr) Then
                        '        TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION, getPosTxtNew(TheModCntr.contxt, "LevelITCommonHelloWorldFailed", 0), 0, TheModCntr, "LevelITCommonHelloWorldFailed")
                        '    End If

                End Select

            End If


            ModBase_run = 0

        Catch ex As Exception

            LOG_Error(getLocationString(szFuncName), ex.Message)

        Finally
            LOG_FuncExit(getLocationString(szFuncName), "Function ModBase_run returns " & ModBase_run.ToString)
        End Try
    End Function
#End Region

#Region "IGiftCardBalanceInquiry"

    Public Function InquiryGiftCard(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA, ByRef szBarcode As String, ByRef dBalance As System.Decimal, ByRef szReceipt As String) As IGiftCardReturnCode
        InquiryGiftCard = IGiftCardReturnCode.OK
        Dim funcName As String = "CheckRedeemGiftCard"
        Dim handler As IGiftCardBalanceInquiry
        Dim parameters As Dictionary(Of String, Object)

        Try
            parameters = New Dictionary(Of String, Object) From { _
                                                             {"Controller", TheModCntr}, _
                                                             {"Transaction", taobj}, _
                                                             {"Barcode", szBarcode}, _
                                                             {"Value", dBalance}, _
                                                             {"Receipt", szReceipt}, _
                                                             {"GiftCardBalanceInternalInquiry", False} _
                                                            }

            handler = createPosModelObject(Of IGiftCardBalanceInquiry)(TheModCntr, "GiftCardController", 0, False)
            If handler Is Nothing Then
                ' gift card handler is not defined into the database
                InquiryGiftCard = IGiftCardReturnCode.KO
                Exit Function
            End If

            InquiryGiftCard = handler.GiftCardBalanceInquiry(parameters)

            ' adjustment for value data type
            szBarcode = parameters("Barcode")
            dBalance = parameters("Value")
            szReceipt = parameters("Receipt")

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "returns ")
        End Try
    End Function

#End Region

#Region "Protected Functions"
    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
#End Region

End Class
