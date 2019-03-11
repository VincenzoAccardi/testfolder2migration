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
Imports System.Xml.Linq
Imports System.Xml.XPath
Imports System.Linq

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
                    InquiryGiftCard(TheModCntr, taobj)
                    Exit Select

                Case Else
                    LOG_Error(getLocationString(szFuncName), "Invalid option: " + TheModCntr.ModulNmbrExt.ToString)
                    Exit Function

            End Select

            If TypeOf frm Is FormEFTFunctions Then

                frm.theModCntr = TheModCntr
                frm.taobj = taobj
                Dim bFrmRet As Object = frm.DisplayMe(taobj, TheModCntr, FormRoot.DialogActive.d1_DialogActive)
                frm.Close()
                If Not bFrmRet Then Exit Function

                Select Case CType(frm, FormEFTFunctions).OptionSelected

                    Case FormEFTFunctions._Options.No_Option_Selected
                        TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION, getPosTxtNew(TheModCntr.contxt, "LevelITCommonHelloWorldFailed", 0), 0, TheModCntr, "LevelITCommonHelloWorldFailed")
                End Select

            End If

            If taobj.getFtrRecNr = -1 Then
                taobj.TAEnd(fillFooterLines(TheModCntr.con, taobj, TheModCntr))
            End If
            taobj.bPrintReceipt = False
            taobj.bTAtoFile = True
            taobj.bDelete = True

            ModBase_run = 0

        Catch ex As Exception

            LOG_Error(getLocationString(szFuncName), ex.Message)

        Finally
            LOG_FuncExit(getLocationString(szFuncName), "Function ModBase_run returns " & ModBase_run.ToString)
        End Try
    End Function
#End Region

#Region "IGiftCardBalanceInquiry"

    Public Function InquiryGiftCard(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA) As IGiftCardReturnCode
        InquiryGiftCard = IGiftCardReturnCode.OK
        Dim funcName As String = "CheckRedeemGiftCard"
        Dim handler As IGiftCardBalanceInquiry
        Dim TaHelper As New TPDotnet.IT.Common.Pos.TPTAHelper
        Try

            handler = createPosModelObject(Of IGiftCardBalanceInquiry)(TheModCntr, TPDotnet.IT.Common.Pos.GiftCardItem.ToString, 0, False)
            If handler Is Nothing Then
                ' gift card handler is not defined into the database
                InquiryGiftCard = IGiftCardReturnCode.KO
                Exit Function
            End If
            Dim myheader As TPDotnet.Pos.TaHdrRec = CType(taobj.GetTALine(1), TPDotnet.Pos.TaHdrRec)
            Dim result As Integer = handler.GiftCardBalanceInquiry(New Dictionary(Of String, Object) From {
                                                                     {"Controller", TheModCntr},
                                                                     {"Transaction", taobj},
                                                                     {"CurrentRecord", myheader},
                                                                     {"CurrentDetailRecord", Nothing},
                                                                     {"Parameters", New TPDotnet.IT.Common.Pos.CommonParameters}
                                                                 })

            If result = 0 Then

                Dim ExternalServiceRec As TPDotnet.IT.Common.Pos.TaExternalServiceRec
                Dim lTaCreateNmbr As Integer = myheader.theHdr.lTaCreateNmbr
                Dim listTaExtCreateNmbr As List(Of XElement) = taobj.TAtoXDocument(False, 0, False).XPathSelectElements("/TAS/NEW_TA/EXTERNAL_SERVICE[Hdr/lTaRefToCreateNmbr=" + lTaCreateNmbr.ToString + "]/Hdr/lTaCreateNmbr").ToList

                For Each lTaExtCreateNmbr As XElement In listTaExtCreateNmbr

                    'Check if exist External Service Rercod into original TA
                    If lTaExtCreateNmbr IsNot Nothing Then
                        ExternalServiceRec = CType(taobj.GetTALine(taobj.GetPositionFromCreationNmbr(lTaExtCreateNmbr.Value)), TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                    Else
                        Throw New Exception("EXTERNAL_SERVICE_RECORD_NOT_FOUND")
                    End If

                    TaHelper.ExternalTAHandler(taobj, TheModCntr, ExternalServiceRec)
                Next


            End If
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
