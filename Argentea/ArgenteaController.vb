Imports System
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos
Imports Microsoft.VisualBasic
Imports System.Reflection
Imports ARGLIB = PAGAMENTOLib
Imports System.Linq
Imports TPDotnet.Pos

Public Class Controller
    Implements IElectronicFundsTransferPay
    Implements IElectronicFundsTransferVoid
    Implements IElectronicFundsTransferClosure
    Implements IElectronicFundsTransferTotals
    Implements IPhoneRechargeActivationPreCheck
    Implements IPhoneRechargeActivation
    Implements IElectronicMealVoucherPay
    Implements IElectronicMealVoucherVoid
    Implements IExternalGiftCardActivation
    Implements IExternalGiftCardDeActivation
    Implements IExternalGiftCardConfirm

#Region "Enum"
    Protected Enum MethodParameter
        ArgenteaCOMObject
        taobj
        TheModCntr
        MyCurrentRecord
        MyCurrentDetailRec
        paramArg
    End Enum
    Protected Enum Controller
        None
        EFTController
        DTPController
        ADVController
        BPCController
        BPEController
        BPCeliacController
        GiftCardController
        ExternalGiftCardController
        PhoneRechargeController
    End Enum
    Protected Enum Method
        Payment
        Void
        Check
        Closure
        Totals
        CheckPhoneRecharge
        ActivatePhoneRecharge
        ActivationExternalGiftCard
        DeActivationExternalGiftCard
        ConfirmExternalGiftCard
    End Enum

#End Region

#Region "Dictionary"

    Protected GetController As New Dictionary(Of String, Controller) From {
        {"SATISPAY", Controller.ADVController},
        {"JIFFY", Controller.ADVController},
        {"BITCOIN", Controller.ADVController},
        {"EFT", Controller.EFTController},
        {"PAYFAST", Controller.DTPController},
        {"EMEALCELIAC", Controller.BPCeliacController}
    }

    Protected GetStatus As New Dictionary(Of Method, TaExternalServiceRec.ExternalServiceStatus) From {
        {Method.Payment, TaExternalServiceRec.ExternalServiceStatus.Activated},
        {Method.Void, TaExternalServiceRec.ExternalServiceStatus.Deleted},
        {Method.Closure, TaExternalServiceRec.ExternalServiceStatus.Unkown},
        {Method.Totals, TaExternalServiceRec.ExternalServiceStatus.Unkown},
        {Method.CheckPhoneRecharge, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.ActivatePhoneRecharge, TaExternalServiceRec.ExternalServiceStatus.Activated},
        {Method.Check, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.ActivationExternalGiftCard, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.DeActivationExternalGiftCard, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.ConfirmExternalGiftCard, TaExternalServiceRec.ExternalServiceStatus.Activated}
    }

    Protected OperationParameters As New Dictionary(Of Method, MethodParameter()) From {
        {Method.Payment, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.MyCurrentDetailRec, MethodParameter.paramArg}},
        {Method.Void, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.MyCurrentDetailRec, MethodParameter.paramArg}},
        {Method.Closure, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.paramArg}},
        {Method.Totals, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.paramArg}},
        {Method.CheckPhoneRecharge, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.ActivatePhoneRecharge, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.Check, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.MyCurrentDetailRec, MethodParameter.paramArg}},
        {Method.ActivationExternalGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.DeActivationExternalGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.ConfirmExternalGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}}
    }
#End Region

#Region "Property"
    Protected Property ArgenteaCOMObject As ARGLIB.argpay
        Get
            Return _ArgenteaCOMObject
        End Get
        Set(value As ARGLIB.argpay)
            _ArgenteaCOMObject = value
        End Set
    End Property
    Protected Property TheModCntr As TPDotnet.Pos.ModCntr
        Get
            Return _TheModCntr
        End Get
        Set(value As TPDotnet.Pos.ModCntr)
            _TheModCntr = value
        End Set
    End Property
    Protected Property taobj As TPDotnet.Pos.TA
        Get
            Return _taobj
        End Get
        Set(value As TPDotnet.Pos.TA)
            _taobj = value
        End Set
    End Property
    Protected Property MyCurrentRecord As TPDotnet.Pos.TaBaseRec
        Get
            Return _MyCurrentRecord
        End Get
        Set(value As TPDotnet.Pos.TaBaseRec)
            _MyCurrentRecord = value
        End Set
    End Property
    Protected Property MyCurrentDetailRec As TPDotnet.Pos.TaBaseRec
        Get
            Return _MyCurrentDetailRec
        End Get
        Set(value As TPDotnet.Pos.TaBaseRec)
            _MyCurrentDetailRec = value
        End Set
    End Property
    Protected ReadOnly Property ExternalID As String
        Get
            Select Case MyCurrentRecord.sid
                Case TPDotnet.Pos.TARecTypes.iTA_MEDIA
                    Return CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).PAYMENTinMedia.szExternalID
                Case TPDotnet.Pos.TARecTypes.iTA_ART_SALE
                    Return CType(CType(MyCurrentRecord, TPDotnet.IT.Common.Pos.TaArtSaleRec).ARTinArtSale, TPDotnet.IT.Common.Pos.ART).szITSpecialItemType
                Case TPDotnet.Pos.TARecTypes.iTA_ART_RETURN
                    Return CType(CType(MyCurrentRecord, TPDotnet.IT.Common.Pos.TaArtReturnRec).ARTinArtReturn, TPDotnet.IT.Common.Pos.ART).szITSpecialItemType
                Case TPDotnet.Pos.TARecTypes.iTA_SIGN_OFF
                    Return "SIGNOFF"
                Case Else
                    Return String.Empty
            End Select
        End Get
    End Property
    Protected Property paramArg As ArgenteaParameters
        Get
            Return _paramArg
        End Get
        Set(value As ArgenteaParameters)
            _paramArg = value
        End Set
    End Property

#End Region

#Region "Variabiles"
    Protected _ArgenteaCOMObject As ARGLIB.argpay
    Protected _TheModCntr As TPDotnet.Pos.ModCntr
    Protected _taobj As TPDotnet.Pos.TA
    Protected _MyCurrentRecord As TPDotnet.Pos.TaBaseRec
    Protected _MyCurrentDetailRec As TPDotnet.Pos.TaBaseRec
    Protected _paramArg As ArgenteaParameters
    Protected paramCommon As TPDotnet.IT.Common.Pos.CommonParameters
    Protected NameClass As String
    Protected eController As Controller
#End Region

#Region "Protected Function"
    Protected Overridable Function Execute(ByRef eMethod As Method, ByRef Parameters As Dictionary(Of String, Object)) As Object
        Dim funcName As String = "Execute"
        LOG_Debug(getLocationString(funcName), "Init")
        Dim argenteaFunctionReturnObject(0) As ArgenteaFunctionReturnObject

        Dim response As New ArgenteaResponse
        Execute = 1
        Try

            Init(Parameters)

            If GetStatus(eMethod) = TaExternalServiceRec.ExternalServiceStatus.PreChecked Then
                If Not PrecheckHandlerBeforeInvoke(eMethod, eController) Then
                    Execute = 1
                    Return Execute
                End If
            End If

            TPDotnet.IT.Common.Pos.Common.ShowScreen(TheModCntr, True, paramCommon.WaitScreenName)

            LOG_Debug(getLocationString(funcName), "Create handler for this operation")

            Dim typeClass As Type = Type.GetType(NameClass)

            Dim objInstance As Object = Activator.CreateInstance(typeClass)

            Dim executorParameters As New List(Of Object)
            OperationParameters(eMethod).ToList.ForEach(Sub(ByVal opParam As MethodParameter)
                                                            executorParameters.Add(Me.GetType.GetProperty(opParam.ToString, BindingFlags.Instance Or BindingFlags.GetProperty Or BindingFlags.NonPublic Or BindingFlags.Public).GetValue(Me, Nothing))
                                                        End Sub)

            response = typeClass.InvokeMember(eMethod.ToString, BindingFlags.Default Or BindingFlags.InvokeMethod, Nothing, objInstance, executorParameters.ToArray)

            paramCommon.Copies = paramArg.Copies
            paramCommon.PrintWithinTA = paramArg.PrintWithinTA
            paramCommon.SuppressLogo = paramArg.SuppressLogo
            paramCommon.SaveExternalTa = paramArg.SaveExternalTa

            If response.ReturnCode = ArgenteaFunctionsReturnCode.OK Then
                argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
                If (Not CSVHelper.ParseReturnString(response.MessageOut, response.FunctionType, argenteaFunctionReturnObject, response.CharSeparator)) Then
                    Execute = 1
                    Return Execute
                End If
            Else
                Execute = 1
                Return Execute
            End If

            If GetStatus(eMethod) = TaExternalServiceRec.ExternalServiceStatus.PreChecked Then
                If Not PrecheckHandlerAfterInvoke(eMethod, eController, argenteaFunctionReturnObject, response) Then
                    Execute = 1
                    Return Execute
                End If
            End If

            FillExternalServiceRecord(argenteaFunctionReturnObject, response, eMethod)

            Execute = 0

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex)
        Finally
            TPDotnet.IT.Common.Pos.Common.ShowScreen(TheModCntr, False, paramCommon.WaitScreenName)
            If response.ReturnCode = ArgenteaFunctionsReturnCode.KO AndAlso Not String.IsNullOrEmpty(response.GetProperty("szErrorMessage")) Then
                Dim cmd As Common = New Common
                cmd.ShowError(TheModCntr, response.GetProperty("szErrorMessage"), String.Empty)
            End If
            LOG_FuncExit(getLocationString(funcName), "returns ")
        End Try
        Return Execute
    End Function

    Private Function PrecheckHandlerBeforeInvoke(eMethod As Method, eController As Controller) As Boolean
        Dim cmd As New TPDotnet.IT.Common.Pos.Common
        Select Case eController
            Case Controller.ExternalGiftCardController
                PrecheckHandlerBeforeInvoke = ExtGiftCardFormHandler()
            Case Else
                PrecheckHandlerBeforeInvoke = True
        End Select
        Return PrecheckHandlerBeforeInvoke
    End Function

    Private Function ExtGiftCardFormHandler() As Boolean
        Dim szBarcode As String = String.Empty
        Dim szCaptionDescription As String = String.Empty
        If TypeOf (MyCurrentRecord) Is TPDotnet.Pos.TaArtSaleRec Then
            szBarcode = GetBarcodeFromTemplate(taobj, TheModCntr, MyCurrentRecord.GetPropertybyName("szInputString"))
            szCaptionDescription = CType(MyCurrentRecord, TPDotnet.Pos.TaArtSaleRec).ARTinArtSale.szDesc
        ElseIf TypeOf (MyCurrentRecord) Is TPDotnet.Pos.TaArtReturnRec Then
            szBarcode = MyCurrentRecord.GetPropertybyName("szITExtGiftCardEAN")
            szCaptionDescription = CType(MyCurrentRecord, TPDotnet.Pos.TaArtReturnRec).ARTinArtReturn.szDesc
        End If
        If String.IsNullOrEmpty(szBarcode) Then
            szBarcode = CallForm(szCaptionDescription)
            If String.IsNullOrEmpty(szBarcode) Then
                ExtGiftCardFormHandler = False
            Else
                If Not MyCurrentRecord.ExistField("szITExtGiftCardEAN") Then MyCurrentRecord.AddField("szITExtGiftCardEAN", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                MyCurrentRecord.setPropertybyName("szITExtGiftCardEAN", szBarcode.ToString)
                ExtGiftCardFormHandler = True
            End If
        Else
            ExtGiftCardFormHandler = True
        End If
        Return ExtGiftCardFormHandler
    End Function

    Protected Overridable Function GetBarcodeFromTemplate(taobj As TPDotnet.Pos.TA, TheModCntr As ModCntr, ByVal szInputString As String) As String
        Dim funcName As String = "GetBarcodeFromTemplate"

        GetBarcodeFromTemplate = ""
        Try
            TheModCntr.InputString_Renamed = szInputString
            TheModCntr.theBarcodeCls.SetNormalBuffer = TheModCntr.InputString_Renamed
            If ScanTemplateWithCheck("all", BARCODE_TEMPLATE, TheModCntr, taobj, False) = True Then

                If TheModCntr.theBarcodeCls.Hits > 0 Then
                    ' ok , we have hits for Barcode
                    For i As Integer = 1 To TheModCntr.theBarcodeCls.Hits
                        'we identified the barcode
                        'it is a voucher barcode, which is not serialized
                        'the MediaMemberNo is stored as ExternalID
                        Dim szString As String = TheModCntr.theBarcodeCls.GetMatchByName(i)
                        Select Case szString
                            Case TD_SERIAL_NUMBER
                                GetBarcodeFromTemplate = TheModCntr.theBarcodeCls.GetMatch(i)
                        End Select
                    Next i
                End If
            End If
        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            TheModCntr.theBarcodeCls.SetNormalBuffer = String.Empty
        End Try


        Return GetBarcodeFromTemplate
    End Function

    Private Function CallForm(ByRef szDescription As String) As String
        Dim giftForm As FormArgenteaItemInput = TheModCntr.GetCustomizedForm(GetType(FormArgenteaItemInput), STRETCH_TO_SMALL_WINDOW)
        giftForm.ArticleDescription = szDescription
        CallForm = giftForm.DisplayMe(taobj, TheModCntr, FormRoot.DialogActive.d1_DialogActive)
        giftForm.Close()
        Return CallForm
    End Function

    Private Function PrecheckHandlerAfterInvoke(ByRef eMethod As Method, ByRef eController As Controller, ByRef argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject, ByRef response As ArgenteaResponse) As Boolean
        Dim cmd As New TPDotnet.IT.Common.Pos.Common
        Select Case eController
            Case Controller.BPCeliacController
                If Not argenteaFunctionReturnObject(0).Successfull Then
                    If argenteaFunctionReturnObject(0).CodeResult = CodeResult.UnderFunded Then
                        TPDotnet.IT.Common.Pos.Common.ShowScreen(TheModCntr, False, paramCommon.WaitScreenName)
                        If MsgBoxResult.Ok = cmd.ShowQuestion(TheModCntr, "LevelITCommonArgenteaUnderFunded", CDec((CInt(argenteaFunctionReturnObject(0).Amount) / 100)).ToString) Then
                            PrecheckHandlerAfterInvoke = True
                        Else
                            PrecheckHandlerAfterInvoke = False
                            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                        End If
                    Else
                        PrecheckHandlerAfterInvoke = True
                        eMethod = Method.Payment
                    End If
                End If
            Case Else
                PrecheckHandlerAfterInvoke = True
        End Select
        Return PrecheckHandlerAfterInvoke
    End Function

    Private Sub FillExternalServiceRecord(argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject, response As ArgenteaResponse, eMethod As Method)
        For i As Integer = 0 To argenteaFunctionReturnObject.GetUpperBound(0)
            If String.IsNullOrEmpty(response.TransactionID) Then
                response.TransactionID = argenteaFunctionReturnObject(i).TerminalID
            End If

            Dim lAmount As Integer = 0
            If response.ExistProperty("lAmount") Then
                lAmount = response.GetProperty("lAmount")
            End If
            If (Integer.TryParse(argenteaFunctionReturnObject(i).Amount, 0)) Then
                lAmount = CInt(argenteaFunctionReturnObject(i).Amount)
            End If

            Dim szReceipt As String = String.Empty
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).Receipt) Then
                szReceipt = argenteaFunctionReturnObject(i).Receipt
            End If

            Dim TaExternalServiceRec As TaExternalServiceRec = taobj.CreateTaObject(Of TaExternalServiceRec)(Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE)
            If Not TaExternalServiceRec.ExistField("szCardType") Then TaExternalServiceRec.AddField("szCardType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            If Not TaExternalServiceRec.ExistField("szOperationResult") Then TaExternalServiceRec.AddField("szOperationResult", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            If Not TaExternalServiceRec.ExistField("lAmount") Then TaExternalServiceRec.AddField("lAmount", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            If Not TaExternalServiceRec.ExistField("szTransactionID") Then TaExternalServiceRec.AddField("szTransactionID", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

            If response.ExistProperty("lPinCounter") Then
                If Not TaExternalServiceRec.ExistField("lPinCounter") Then TaExternalServiceRec.AddField("lPinCounter", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                TaExternalServiceRec.setPropertybyName("lPinCounter", response.GetProperty("lPinCounter"))
            End If
            If response.ExistProperty("szBarcode") Then
                If Not TaExternalServiceRec.ExistField("szBarcode") Then TaExternalServiceRec.AddField("szBarcode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szBarcode", response.GetProperty("szBarcode"))
            End If

            With TaExternalServiceRec
                .theHdr.lTaRefToCreateNmbr = MyCurrentRecord.theHdr.lTaCreateNmbr
                If Not String.IsNullOrEmpty(szReceipt) Then .szReceipt = szReceipt
                .szServiceType = GetOperationType(ExternalID)
                .lCopies = paramArg.Copies
                .szStatus = GetStatus(eMethod).ToString()
                .bPrintReceipt = paramArg.PrintWithinTA
                .setPropertybyName("lAmount", lAmount)
                .setPropertybyName("szTransactionID", response.TransactionID)
                .setPropertybyName("szCardType", argenteaFunctionReturnObject(i).Acquirer.ToString())
                .setPropertybyName("szOperationResult", response.ReturnCode.ToString())
            End With

            If MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA Then
                Dim MyTaMediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
                With MyTaMediaRec
                    .dTaPaidTotal = CDec(lAmount / 100)
                    .dTaPaid = CDec(lAmount / 100)
                    .dPaidForeignCurr = CDec(lAmount / 100)
                    .dPaidForeignCurrTotal = CDec(lAmount / 100)
                End With
            End If

            taobj.Add(TaExternalServiceRec)
        Next

    End Sub

    Protected Sub Init(ByRef Parameters As Dictionary(Of String, Object))
        paramArg = New ArgenteaParameters()
        ArgenteaCOMObject = New ARGLIB.argpay()
        TheModCntr = Parameters("Controller")
        taobj = Parameters("Transaction")
        paramCommon = Parameters("Parameters")
        MyCurrentRecord = Parameters("CurrentRecord")
        MyCurrentDetailRec = Parameters("CurrentDetailRecord")

        paramArg.LoadParametersByReflection(TheModCntr, "Argentea")
        paramCommon.WaitScreenName = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "PROCESS_NAME").Trim
        Dim szNameSpace As String = GetType(Controller).Namespace
        If eController = Controller.None Then
            eController = GetController(GetOperationType(ExternalID))
        End If
        NameClass = szNameSpace + "." + eController.ToString
    End Sub
    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
#End Region

#Region "Private Function"
    Private Function GetOperationType(ByVal szExternalID As String) As String
        Dim GetService As String = String.Empty
        Dim funcName As String = "GetService"
        Dim availableServices As String() = {JiffyMedia, SatispayMedia, BitCoinMedia, ElectronicFundsTransferMedia, ElectronicMealVoucherCeliacMedia, PayFastMedia, PhoneRechargeItem, ExternalGiftCardItem, GiftCardItem}

        Try
            ' get & check the current payment service
            GetService = availableServices.Where(Function(s) szExternalID.StartsWith(s.ToUpper)).FirstOrDefault

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "returns ")
        End Try
        Return GetService
    End Function
#End Region

#Region "Public Function"
    Public Function ElectronicFundsTransferPay_Pay(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferPay.Pay
        Return Execute(Method.Payment, Parameters)
    End Function
    Public Function ElectronicFundsTransferVoid_Void(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferVoid.Void
        Return Execute(Method.Void, Parameters)
    End Function
    Public Function ElectronicFundsTransferClosure_Closure(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferClosure.Closure
        eController = Controller.EFTController
        Return Execute(Method.Closure, Parameters)
    End Function
    Public Function ElectronicFundsTransferTotals_Totals(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferTotals.Totals
        eController = Controller.EFTController
        Return Execute(Method.Totals, Parameters)
    End Function
    Public Function IElectronicMealVoucherPay_Pay(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherPay.Pay
        Return Execute(Method.Payment, Parameters)
    End Function
    Public Function IElectronicMealVoucherVoid_Void(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherVoid.Void
        Return Execute(Method.Void, Parameters)
    End Function
    Public Function CheckPhoneRecharge(ByRef Parameters As Dictionary(Of String, Object)) As IPhoneRechargeReturnCode Implements IPhoneRechargeActivationPreCheck.CheckPhoneRecharge
        eController = Controller.PhoneRechargeController
        Return Execute(Method.CheckPhoneRecharge, Parameters)
    End Function
    Public Function ActivatePhoneRecharge(ByRef Parameters As Dictionary(Of String, Object)) As IPhoneRechargeReturnCode Implements IPhoneRechargeActivation.ActivatePhoneRecharge
        eController = Controller.PhoneRechargeController
        Return Execute(Method.ActivatePhoneRecharge, Parameters)
    End Function
    Public Function ActivationExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardActivation.ActivationExternalGiftCard
        eController = Controller.ExternalGiftCardController
        Return Execute(Method.ActivationExternalGiftCard, Parameters)
    End Function

#Region "Check"
    Public Function ElectronicFundsTransferPay_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferPay.Check
        Return IElectronicFundsTransferReturnCode.OK
    End Function
    Private Function ElectronicFundsTransferVoid_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferVoid.Check
        Return IElectronicFundsTransferReturnCode.OK
    End Function
    Public Function ElectronicFundsTransferCheck_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferClosure.Check
        Return IElectronicFundsTransferReturnCode.OK
    End Function
    Public Function ElectronicFundsTransferTotals_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferTotals.Check
        Return IElectronicFundsTransferReturnCode.OK
    End Function
    Public Function IElectronicMealVoucherPay_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherPay.Check
        Return Execute(Method.Check, Parameters)
    End Function
    Private Function IElectronicMealVoucherVoid_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherVoid.Check
        Return IElectronicFundsTransferReturnCode.OK
    End Function

    Public Function DeActivationExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardDeActivation.DeActivationExternalGiftCard
        eController = Controller.ExternalGiftCardController
        Return Execute(Method.DeActivationExternalGiftCard, Parameters)
    End Function

    Public Function ConfirmExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardConfirm.ConfirmExternalGiftCard
        eController = Controller.ExternalGiftCardController
        Return Execute(Method.ConfirmExternalGiftCard, Parameters)
    End Function
#End Region
#End Region

End Class

Public Class CodeResult
    Public Const UnderFunded As String = "0300"
End Class
