Imports System
Imports TPDotnet.Pos
Imports System.IO
Imports Microsoft.VisualBasic

Public Class TPTAHelperArgentea
    Inherits TPTAHelper

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

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

    Public Overridable Function ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(ByRef taobj As TPDotnet.Pos.TA, ByVal obj As ArgenteaFunctionReturnObject) As TaExternalServiceRec
        ArgenteaFunctionReturnObjectToTaArgenteaEMVRec = Nothing
        Dim funcName As String = "ArgenteaFunctionReturnObjectToTaArgenteaEMVRec"

        Try
            ArgenteaFunctionReturnObjectToTaArgenteaEMVRec = taobj.CreateTaObject(TPDotnet.IT.Common.Pos.Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE)
            If ArgenteaFunctionReturnObjectToTaArgenteaEMVRec Is Nothing Then
                Exit Function
            End If

            If Not ArgenteaFunctionReturnObjectToTaArgenteaEMVRec.ExistField("szFunctionID") Then
                ArgenteaFunctionReturnObjectToTaArgenteaEMVRec.AddField("szFunctionID", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            End If
            ArgenteaFunctionReturnObjectToTaArgenteaEMVRec.setPropertybyName("szFunctionID", obj.ArgenteaFunction.ToString)
            ArgenteaFunctionReturnObjectToTaArgenteaEMVRec.szReceipt = obj.Receipt
            ' ... fills all properties

        Catch ex As Exception
            ' to do log
        Finally

        End Try

    End Function

    Protected Class HandlingParamter

        Private _NoOfAdditionalTaCopy As Integer
        Public Property NoOfAdditionalTaCopy() As Integer
            Get
                Return _NoOfAdditionalTaCopy
            End Get
            Set(ByVal value As Integer)
                _NoOfAdditionalTaCopy = value
            End Set
        End Property

        Private _CreateAdditionalTa As Boolean
        Public Property CreateAdditionalTa() As Boolean
            Get
                Return _CreateAdditionalTa
            End Get
            Set(ByVal value As Boolean)
                _CreateAdditionalTa = value
            End Set
        End Property

        Private _PrintWithinTa As Boolean
        Public Property PrintWithinTa() As Boolean
            Get
                Return _PrintWithinTa
            End Get
            Set(ByVal value As Boolean)
                _PrintWithinTa = value
            End Set
        End Property

        Private _LineIdentifier As String
        Public Property LineIdentifier() As String
            Get
                Return _LineIdentifier
            End Get
            Set(ByVal value As String)
                _LineIdentifier = value
            End Set
        End Property

        Private _GiftCardBalanceInternalInquiry As Boolean
        Public Property GiftCardBalanceInternalInquiry() As Boolean
            Get
                Return _GiftCardBalanceInternalInquiry
            End Get
            Set(ByVal value As Boolean)
                _GiftCardBalanceInternalInquiry = value
            End Set
        End Property

    End Class

    Protected Overridable Function LoadHandlingParamterByArgenteaFuntion(ByVal argenteaFunction As InternalArgenteaFunctionTypes,
                                                   ByVal parameters As ArgenteaParameters) As HandlingParamter

        Dim HandlingParamter As New HandlingParamter()
        Dim funcName As String = "LoadHandlingParamterByArgenteaFuntion"

        Try

            Select Case argenteaFunction

                Case InternalArgenteaFunctionTypes.GiftCardRedeem
                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.GiftCardRedeemCopies
                        .CreateAdditionalTa = parameters.GiftCardRedeemSave
                    End With

                    'HandlingParamter.PrintWithinTa = parameters.GiftCardRedeemPrintWithinTa
                    Exit Select

                Case InternalArgenteaFunctionTypes.GiftCardRedeemPreCkeck
                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.GiftCardRedeemCheckCopies
                        .CreateAdditionalTa = parameters.GiftCardRedeemCheckSave
                        .PrintWithinTa = parameters.GiftCardRedeemCheckPrintWithinTa
                    End With

                    Exit Select

                Case InternalArgenteaFunctionTypes.GiftCardActivation
                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.GiftCardActivationCopies
                        .CreateAdditionalTa = parameters.GiftCardActivationSave
                        .PrintWithinTa = parameters.GiftCardActivationPrintWtihinTa
                        'LoadHandlingParamterByArgenteaFuntion.NumberOfCopy = parameters.GiftCardRedeem
                    End With
                    Exit Select

                Case InternalArgenteaFunctionTypes.GiftCardBalance

                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.GiftCardBalanceCopies
                        .CreateAdditionalTa = parameters.GiftCardBalanceSave
                        .PrintWithinTa = parameters.GiftCardBalancePrintWithinTa
                        .LineIdentifier = parameters.GiftCardBalanceLineIdentifier
                        .GiftCardBalanceInternalInquiry = parameters.GiftCardBalanceInternalInquiry
                    End With
                    Exit Select

                Case InternalArgenteaFunctionTypes.GiftCardRedeemCancel

                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.GiftCardRedeemCancelCopies
                        .CreateAdditionalTa = parameters.GiftCardRedeemCancelSave
                        .PrintWithinTa = parameters.GiftCardBalancePrintWithinTa
                        .LineIdentifier = parameters.GiftCardBalanceLineIdentifier
                        .GiftCardBalanceInternalInquiry = parameters.GiftCardBalanceInternalInquiry
                    End With
                    Exit Select

                Case InternalArgenteaFunctionTypes.PhoneRechargeCheck

                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.PhoneRechargeCheckCopies
                        .CreateAdditionalTa = parameters.PhoneRechargeCheckSave
                        .PrintWithinTa = parameters.PhoneRechargeCheckPrintWithinTa
                    End With
                    Exit Select

                Case InternalArgenteaFunctionTypes.PhoneRechargeActivation

                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.PhoneRechargeActivationCopies
                        .CreateAdditionalTa = parameters.PhoneRechargeActivationSave
                        .PrintWithinTa = parameters.PhoneRechargeActivationPrintWithinTa
                    End With

                    Exit Select
                Case InternalArgenteaFunctionTypes.ExternalGiftCardActivation

                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.ExtGiftCardActivationCopies
                        .CreateAdditionalTa = parameters.ExtGiftCardActivationSave
                        .PrintWithinTa = parameters.ExtGiftCardActivationPrintWithinTa
                    End With
                    Exit Select
                Case InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation

                    With HandlingParamter
                        .NoOfAdditionalTaCopy = parameters.ExtGiftCardDeActivationCancelCopies
                        .CreateAdditionalTa = parameters.ExtGiftCardDeActivationCancelSave
                        .PrintWithinTa = parameters.ExtGiftCardDeActivationPrintWithinTa
                    End With
                    Exit Select


                    'Case InternalArgenteaFunctionTypes.BPCeliacPayment

                    '    With HandlingParamter
                    '        .NoOfAdditionalTaCopy = parameters.BPCeliacCopies
                    '        .PrintWithinTa = parameters.BPCeliacPrintWithinTa
                    '    End With
                    '    Exit Select

                Case Else
                    LoadHandlingParamterByArgenteaFuntion = Nothing

            End Select

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

        Return HandlingParamter

    End Function

    Public Overridable Function HandleReturnString(ByRef taobj As TPDotnet.Pos.TA,
                                                   ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                                                   ByVal returnString As String,
                                                   ByVal argenteaFunction As InternalArgenteaFunctionTypes,
                                                   ByVal parameters As ArgenteaParameters,
                                                   Optional ByRef taArgenteaEMVRec As TaExternalServiceRec = Nothing
                                                   ) As Boolean

        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject
        Dim objTPTAHelper As TPTAHelper = Nothing
        Dim argenteaTA As TPDotnet.Pos.TA = Nothing
        Dim successfull As String = String.Empty
        Dim funcName As String = "HandleReturnString"
        Dim handlingParamter As HandlingParamter

        Try
            LOG_Debug(getLocationString(funcName), "ArgenteaFunction: <" & argenteaFunction.ToString & ">")
            LOG_Debug(getLocationString(funcName), "Return string: <" & returnString & ">")

            ' init
            taArgenteaEMVRec = Nothing
            objTPTAHelper = New TPTAHelper
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            parameters.LoadParametersByReflection(TheModCntr)
            handlingParamter = LoadHandlingParamterByArgenteaFuntion(argenteaFunction, parameters)

            ' parse the CSV ans fill the ArgenteaFunctionReturnObject object
            If (Not CSVHelper.ParseReturnString(returnString,
                                                argenteaFunction,
                                                argenteaFunctionReturnObject)) Then
                Throw New Exception("CSV_INVALID")
            End If

            successfull = IIf(argenteaFunctionReturnObject(0).Successfull, "OK", "KO")

            ' create and fill the TaArgenteaEMVRec based on data in ArgenteaFunctionReturnObject
            taArgenteaEMVRec = ArgenteaFunctionReturnObjectToTaArgenteaEMVRec(taobj, argenteaFunctionReturnObject(0))
            If taArgenteaEMVRec Is Nothing Then
                Throw New Exception("CANNOT_CREATE_ARGENTEA_REC")
            End If

            ' create the TP transaction containing the TaArgenteaEMVRec
            argenteaTA = CreateTA(taobj, TheModCntr, taArgenteaEMVRec, handlingParamter.CreateAdditionalTa)
            If argenteaTA Is Nothing Then
                Throw New Exception("CANNOT_CREATE_ARGENTEA_TA")
            End If

            ' write the TP transaction containing the TaArgenteaEMVRec
            If handlingParamter.CreateAdditionalTa AndAlso Not WriteTA(argenteaTA, TheModCntr) Then
                Throw New Exception("CANNOT_WRITE_ARGENTEA_TA")
            End If

            If handlingParamter.GiftCardBalanceInternalInquiry Then
                Exit Function
            End If

            ' print the TP transaction containing the TaArgenteaEMVRec
            For i As Integer = 1 To handlingParamter.NoOfAdditionalTaCopy
                If Not PrintReceipt(argenteaTA, TheModCntr) Then
                    Throw New Exception("CANNOT_PRINT_ARGENTEA_TA")
                End If
            Next

            If handlingParamter.PrintWithinTa Then
                taArgenteaEMVRec.theHdr.lTaCreateNmbr = 0
                taArgenteaEMVRec.theHdr.lTaRefToCreateNmbr = 0
                taArgenteaEMVRec.bPrintReceipt = True
                taobj.Add(taArgenteaEMVRec)
            End If

            If argenteaFunction = InternalArgenteaFunctionTypes.GiftCardBalance Then
                If taobj.getFtrRecNr = -1 Then
                    taobj.TAEnd(fillFooterLines(TheModCntr.con, taobj, TheModCntr))
                End If
                taobj.bPrintReceipt = False
                taobj.bTAtoFile = True
                taobj.bDelete = True
            Else
                If argenteaTA.getFtrRecNr = -1 Then
                    argenteaTA.TAEnd(fillFooterLines(TheModCntr.con, taobj, TheModCntr))
                End If
                argenteaTA.bPrintReceipt = False
                argenteaTA.bTAtoFile = True
                argenteaTA.bDelete = True
            End If


        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Function

End Class
