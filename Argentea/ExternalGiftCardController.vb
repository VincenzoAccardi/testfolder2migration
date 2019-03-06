Imports System
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports Microsoft.VisualBasic
Imports System.Xml.Linq
Imports System.Xml.XPath
Imports System.Linq

Public Class ExternalGiftCardController

#Region "Public Function"

#Region "IExternalGiftCardActivation"

    Public Function ActivationExternalGiftCard(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "ActivationExternalGiftCard"
        Dim response As New ArgenteaResponse

        'Dim p As ExternalGiftCardActivationParameters = New ExternalGiftCardActivationParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IExternalGiftCardActivation function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            Dim MyTaArtSaleRec As TPDotnet.Pos.TaArtSaleRec = CType(MyCurrentRecord, TPDotnet.Pos.TaArtSaleRec)

            Dim szEanCode As String = GetEanCode(MyCurrentRecord)
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            Dim szBarcode As String = MyCurrentRecord.GetPropertybyName("szITExtGiftCardEAN")
            Dim lAmount As Integer = CInt(MyTaArtSaleRec.dTaTotal * 100)
            response.ReturnCode = ArgenteaCOMObject.Attivazione(lAmount, 0, szEanCode, szBarcode, szMessageOut, szErrorMessage)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)
            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                response.MessageOut = "KO" & ";" & szMessageOut & vbCrLf &
                                                    "!!!ERRORE DI ATTIVAZIONE!!!" & vbCrLf &
                                                    szErrorMessage & vbCrLf &
                                                    "Giftcard Serial: " & szBarcode & vbCrLf &
                                                    "Value: " & Math.Round(lAmount / 100, 2) & vbCrLf & vbCrLf & " " & ";" & szErrorMessage
                response.SetProperty("szErrorMessage", szErrorMessage)
            Else
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.ExternalGiftCardActivation


            paramArg.Copies = paramArg.ExtGiftCardActivationCopies
            paramArg.PrintWithinTA = paramArg.ExtGiftCardActivationPrintWithinTa

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response
    End Function
#End Region

#Region "IExternalGiftCardDeActivation"
    Public Function DeActivationExternalGiftCard(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "IExternalGiftCardDeActivation"
        Dim response As New ArgenteaResponse


        Try

            Dim szEanCode As String = GetEanCode(MyCurrentRecord)
            Dim szBarcode As String = MyCurrentRecord.GetPropertybyName("szITExtGiftCardEAN")
            Dim szMessageOut As String = String.Empty
            Dim szErrorMessage As String = String.Empty
            Dim dTaTotal As Decimal = CDec(MyCurrentRecord.GetPropertybyName("dTaTotal").ToString().Replace(".", ","))
            Dim lAmount As Integer = CInt(dTaTotal * 100)

            response.ReturnCode = ArgenteaCOMObject.DisAttivazione(lAmount, szEanCode, szBarcode, szMessageOut, szErrorMessage)
            LOG_Debug(getLocationString(funcName), "ReturnCode: " & response.ReturnCode & ". Giftcard: " & szBarcode & ". Error: " & szErrorMessage & ". Output: " & szMessageOut)

            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                response.MessageOut = "KO" & ";" & szMessageOut & ";" & szErrorMessage
                response.SetProperty("szErrorMessage", szErrorMessage)
            Else
                response.MessageOut = "OK" & ";" & szMessageOut & ";" & szErrorMessage
            End If
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.ExternalGiftCardDeActivation


            paramArg.Copies = paramArg.ExtGiftCardDeActivationPrintWithinTa
            paramArg.PrintWithinTA = paramArg.ExtGiftCardDeActivationCancelCopies


        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response
    End Function
#End Region

#Region "IExternalGiftCardConfirm"
    Public Function ConfirmExternalGiftCard(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse

        Dim funcName As String = "IExternalGiftCardConfirm"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As ExternalGiftCardConfirmParameters = New ExternalGiftCardConfirmParameters
        Dim CSV As String = String.Empty
        Dim retCode As Integer = ArgenteaFunctionsReturnCode.KO
        Dim response As New ArgenteaResponse

        Try
            Dim MyTaExternalService As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(MyCurrentRecord, TPDotnet.IT.Common.Pos.TaExternalServiceRec)


            response.ReturnCode = ArgenteaFunctionsReturnCode.OK
            response.MessageOut = "OK;" + MyTaExternalService.szOriginalReceipt + ";;"
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.ExternalGiftCardConfirm



            paramArg.Copies = paramArg.ExtGiftCardConfirmCopies
            paramArg.PrintWithinTA = paramArg.ExtGiftCardConfirmPrintWithinTa

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response

    End Function

#End Region
#End Region

#Region "Overridable"
    Protected Overridable Function GetEanCode(ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec) As String
        GetEanCode = String.Empty
        If TypeOf (MyCurrentRecord) Is TPDotnet.Pos.TaArtSaleRec Then
            Dim MyTaArtSaleRec As TPDotnet.Pos.TaArtSaleRec = CType(MyCurrentRecord, TPDotnet.Pos.TaArtSaleRec)
            GetEanCode = IIf(String.IsNullOrEmpty(MyTaArtSaleRec.szItemLookupCode), MyTaArtSaleRec.ARTinArtSale.szPOSItemID, MyTaArtSaleRec.szItemLookupCode)
        ElseIf TypeOf (MyCurrentRecord) Is TPDotnet.Pos.TaArtReturnRec Then
            Dim MyTaArtReturnRec As TPDotnet.Pos.TaArtReturnRec = CType(MyCurrentRecord, TPDotnet.Pos.TaArtReturnRec)
            GetEanCode = IIf(String.IsNullOrEmpty(MyTaArtReturnRec.szItemLookupCode), MyTaArtReturnRec.ARTinArtReturn.szPOSItemID, MyTaArtReturnRec.szItemLookupCode)
        End If
        Return GetEanCode
    End Function
    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
