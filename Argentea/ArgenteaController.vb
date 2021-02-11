Imports System
Imports System.Collections.Generic
Imports TPDotnet.IT.Common.Pos
Imports Microsoft.VisualBasic
Imports System.Reflection
Imports ARGLIB = PAGAMENTOLib
Imports System.Linq
Imports TPDotnet.Pos
Imports System.Xml.Linq
Imports System.Xml.XPath

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
    Implements IGiftCardActivationPreCheck
    Implements IGiftCardActivation
    Implements IGiftCardBalanceInquiry
    Implements IGiftCardRedeem
    Implements IGiftCardCancellationPayment
    Implements IElectronicMealVoucherBalance
    Implements IElectronicMealVoucherClosure
    Implements IElectronicMealVoucherTotals
    Implements IValassisCouponValidation
    Implements IValassisCouponNotification

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
        ValassisCouponController
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
        CheckGiftCard
        ActivateGiftCard
        GiftCardBalanceInquiry
        RedeemGiftCard
        GiftCardCancellation
        Balance
        ValidationValassis
        NotificationValassis
    End Enum

#End Region

#Region "Dictionary"

    Protected GetController As New Dictionary(Of String, Controller) From {
        {"SATISPAY", Controller.ADVController},
        {"JIFFY", Controller.ADVController},
        {"BITCOIN", Controller.ADVController},
        {"EFT", Controller.EFTController},
        {"PAYFAST", Controller.DTPController},
        {"EMEALCELIAC", Controller.BPCeliacController},
        {"GIFTCARD", Controller.GiftCardController},
        {"EXTERNALGIFTCARD", Controller.ExternalGiftCardController},
        {"PHONERECHARGE", Controller.PhoneRechargeController},
        {"SIGNOFF", Controller.EFTController},
        {"EMEAL", Controller.BPEController}
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
        {Method.ConfirmExternalGiftCard, TaExternalServiceRec.ExternalServiceStatus.Activated},
        {Method.CheckGiftCard, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.ActivateGiftCard, TaExternalServiceRec.ExternalServiceStatus.Activated},
        {Method.GiftCardBalanceInquiry, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.RedeemGiftCard, TaExternalServiceRec.ExternalServiceStatus.Activated},
        {Method.GiftCardCancellation, TaExternalServiceRec.ExternalServiceStatus.Deleted},
        {Method.Balance, TaExternalServiceRec.ExternalServiceStatus.PreChecked},
        {Method.ValidationValassis, TaExternalServiceRec.ExternalServiceStatus.Activated},
        {Method.NotificationValassis, TaExternalServiceRec.ExternalServiceStatus.Activated}
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
        {Method.ConfirmExternalGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.CheckGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.ActivateGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.GiftCardBalanceInquiry, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.RedeemGiftCard, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.GiftCardCancellation, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.MyCurrentDetailRec, MethodParameter.paramArg}},
        {Method.Balance, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.ValidationValassis, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}},
        {Method.NotificationValassis, {MethodParameter.ArgenteaCOMObject, MethodParameter.taobj, MethodParameter.TheModCntr, MethodParameter.MyCurrentRecord, MethodParameter.paramArg}}
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
                Case TPDotnet.Pos.TARecTypes.iTA_CUSTOMER
                    Return IIf(MyCurrentRecord.ExistField("szServiceType"), MyCurrentRecord.GetPropertybyName("szServiceType"), String.Empty)
                Case TPDotnet.IT.Common.Pos.TARecTypes.iTA_EXTERNAL_SERVICE
                    Return CType(MyCurrentRecord, TPDotnet.IT.Common.Pos.TaExternalServiceRec).szServiceType
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

            If Not HandlerBeforeInvoke(eMethod, eController) Then
                Execute = 1
                Return Execute
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
            paramCommon.SaveExternalTa = paramArg.SaveExternalTa

            'If response.ReturnCode = ArgenteaFunctionsReturnCode.OK Then
            argenteaFunctionReturnObject(0) = New ArgenteaFunctionReturnObject
            If (Not CSVHelper.ParseReturnString(response.MessageOut, response.FunctionType, argenteaFunctionReturnObject, response.CharSeparator, 1, paramArg.EftChiusuraLegacy)) Then
                Execute = 1
                Return Execute
            End If
            'Else
            '    Execute = 1
            '    Return Execute
            'End If

            If Not HandlerAfterInvoke(eMethod, eController, argenteaFunctionReturnObject, response) Then
                Execute = 1
                Return Execute
            End If


            FillExternalServiceRecord(argenteaFunctionReturnObject, response, eMethod)

            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                Execute = 1
            Else
                If argenteaFunctionReturnObject(0).Successfull Then
                    Execute = 0
                Else
                    Execute = 1
                End If
            End If


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


    Protected Sub Init(ByRef Parameters As Dictionary(Of String, Object))
        paramArg = New ArgenteaParameters()
        ArgenteaCOMObject = New ARGLIB.argpay()
        TheModCntr = Parameters("Controller")
        taobj = Parameters("Transaction")
        paramCommon = Parameters("Parameters")
        MyCurrentRecord = Parameters("CurrentRecord")
        MyCurrentDetailRec = Parameters("CurrentDetailRecord")

        paramArg.LoadParametersByReflection(TheModCntr, "Argentea")
        paramArg.Copies = paramCommon.Copies
        paramArg.PrintWithinTA = paramCommon.PrintWithinTA
        paramCommon.WaitScreenName = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "PROCESS_NAME").Trim
        paramCommon.SuppressLogo = IIf(TheModCntr.getParam(PARAMETER_DLL_NAME + ".Argentea." + "PRINT_LOGO_ON_EXTERNAL_RECEIPTS").Trim.ToUpper.Equals("N"), True, False)
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
        Dim availableServices As String() = {JiffyMedia, SatispayMedia, BitCoinMedia, ElectronicFundsTransferMedia,
            ElectronicMealVoucherCeliacMedia, ElectronicMealVoucherMedia,
            PayFastMedia, PhoneRechargeItem, ExternalGiftCardItem, GiftCardItem, "SIGNOFF", ValassisCoupon, ValassisCouponMedia}

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

    Private Function HandlerBeforeInvoke(eMethod As Method, eController As Controller) As Boolean
        Dim cmd As New TPDotnet.IT.Common.Pos.Common
        Select Case eController
            Case Controller.ExternalGiftCardController
                If GetStatus(eMethod) = TaExternalServiceRec.ExternalServiceStatus.PreChecked Then
                    HandlerBeforeInvoke = ExtGiftCardFormHandler()
                Else
                    HandlerBeforeInvoke = True
                End If
            Case Controller.GiftCardController
                If eMethod = Method.CheckGiftCard Then
                    HandlerBeforeInvoke = GiftCardFormHandler(eMethod)
                ElseIf eMethod = Method.GiftCardBalanceInquiry Then
                    HandlerBeforeInvoke = GiftCardFormHandler(eMethod)
                ElseIf eMethod = Method.RedeemGiftCard Then
                    HandlerBeforeInvoke = GiftCardFormHandler(eMethod)
                Else
                    HandlerBeforeInvoke = True
                End If
            Case Else
                HandlerBeforeInvoke = True
        End Select
        Return HandlerBeforeInvoke
    End Function

    Private Function GiftCardFormHandler(ByRef eMethod As Method) As Boolean
        Dim szCaptionDescription As String = String.Empty
        Dim szBarcode As String = String.Empty

        If eMethod = Method.CheckGiftCard Then
            If TypeOf (MyCurrentRecord) Is TPDotnet.Pos.TaArtSaleRec Then
                szCaptionDescription = CType(MyCurrentRecord, TPDotnet.Pos.TaArtSaleRec).ARTinArtSale.szDesc
            End If
            If String.IsNullOrEmpty(szBarcode) Then
                szBarcode = CallForm(GetType(FormItemInput), szCaptionDescription)
            Else
                GiftCardFormHandler = True
            End If
        ElseIf eMethod = Method.GiftCardBalanceInquiry Then
            If MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA Then
                If String.IsNullOrEmpty(CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).szBarcode) Then
                    szCaptionDescription = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).PAYMENTinMedia.szDesc
                    Dim szValue As String = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).dTaPaid.ToString(TheModCntr.getFormatString4Price)
                    szBarcode = CallForm(GetType(FormItemValueInput), szCaptionDescription, szValue)
                    CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).dTaPaid = szValue
                Else
                    szBarcode = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).szBarcode
                End If
            Else
                szBarcode = CallForm(GetType(FormItemInput), szCaptionDescription)
            End If
        ElseIf eMethod = Method.RedeemGiftCard Then
            If Not String.IsNullOrEmpty(CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).szBarcode) Then
                If TheModCntr.bExternalDialog = False AndAlso TheModCntr.bCalledFromWebService = False Then
                    szCaptionDescription = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).PAYMENTinMedia.szDesc
                    Dim szValue As String = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).dTaPaid.ToString(TheModCntr.getFormatString4Price)
                    If String.IsNullOrEmpty(CallForm(GetType(FormItemValueInput), szCaptionDescription, szValue, CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).szBarcode)) Then
                        Return False
                    End If
                    CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).dTaPaid = szValue
                End If
                szBarcode = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).szBarcode
            Else
                Return True
            End If
        Else
            GiftCardFormHandler = True
        End If

        If String.IsNullOrEmpty(szBarcode) Then
            GiftCardFormHandler = False
        Else
            If Not MyCurrentRecord.ExistField("szITGiftCardEAN") Then MyCurrentRecord.AddField("szITGiftCardEAN", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            MyCurrentRecord.setPropertybyName("szITGiftCardEAN", szBarcode.ToString)
            GiftCardFormHandler = True
        End If

        Return GiftCardFormHandler
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
            szBarcode = CallForm(GetType(FormItemInput), szCaptionDescription)
            If String.IsNullOrEmpty(szBarcode) Then
                ExtGiftCardFormHandler = False
            Else
                If Not MyCurrentRecord.ExistField("szITExtGiftCardEAN") Then MyCurrentRecord.AddField("szITExtGiftCardEAN", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                MyCurrentRecord.setPropertybyName("szITExtGiftCardEAN", szBarcode.ToString)
                ExtGiftCardFormHandler = True
            End If
        Else
            If Not MyCurrentRecord.ExistField("szITExtGiftCardEAN") Then MyCurrentRecord.AddField("szITExtGiftCardEAN", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            MyCurrentRecord.setPropertybyName("szITExtGiftCardEAN", szBarcode.ToString)
            ExtGiftCardFormHandler = True
        End If
        Return ExtGiftCardFormHandler
    End Function

    Private Function CallForm(ByVal type As Type, ByRef szDescription As String, Optional ByRef szValue As String = "", Optional szfrmBarcode As String = "") As String
        Dim frmItemInput As Object = Nothing
        Dim szBarcode As String = String.Empty
        If type Is GetType(FormItemInput) Then
            frmItemInput = TheModCntr.GetCustomizedForm(GetType(FormItemInput), STRETCH_TO_SMALL_WINDOW)
        ElseIf type Is GetType(FormItemValueInput) Then
            frmItemInput = TheModCntr.GetCustomizedForm(GetType(FormItemValueInput), STRETCH_TO_SMALL_WINDOW)
            frmItemInput.Value = szValue
            frmItemInput.Barcode = szfrmBarcode
        End If

        frmItemInput.ArticleDescription = szDescription
        szBarcode = frmItemInput.DisplayMe(taobj, TheModCntr, FormRoot.DialogActive.d1_DialogActive)

        If type Is GetType(FormItemValueInput) Then szValue = frmItemInput.Value

        frmItemInput.Close()
        Return szBarcode
    End Function

    Private Function HandlerAfterInvoke(ByRef eMethod As Method, ByRef eController As Controller, ByRef argenteaFunctionReturnObject() As ArgenteaFunctionReturnObject, ByRef response As ArgenteaResponse) As Boolean
        Dim cmd As New TPDotnet.IT.Common.Pos.Common
        Select Case eController
            Case Controller.BPCeliacController
                If GetStatus(eMethod) = TaExternalServiceRec.ExternalServiceStatus.PreChecked Then
                    If Not argenteaFunctionReturnObject(0).Successfull Then
                        If argenteaFunctionReturnObject(0).CodeResult = CodeResult.UnderFunded Then
                            TPDotnet.IT.Common.Pos.Common.ShowScreen(TheModCntr, False, paramCommon.WaitScreenName)
                            If MsgBoxResult.Ok = cmd.ShowQuestion(TheModCntr, "LevelITCommonArgenteaUnderFunded", CDec((CInt(argenteaFunctionReturnObject(0).Amount) / 100)).ToString) Then
                                HandlerAfterInvoke = False
                                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                            Else
                                HandlerAfterInvoke = False
                                response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                            End If
                        Else
                            HandlerAfterInvoke = True
                            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                        End If
                    Else
                        HandlerAfterInvoke = True
                        eMethod = Method.Payment
                    End If
                Else
                    HandlerAfterInvoke = True
                End If
            Case Controller.GiftCardController
                If eMethod = Method.GiftCardBalanceInquiry AndAlso MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA Then
                    Dim lAmount As Decimal = CDec(response.GetProperty("lAmount"))
                    response.SetProperty("lAmount", CInt(Math.Min(CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec).dTaPaid, lAmount) * 100))
                    HandlerAfterInvoke = True
                    If MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA Then
                        If lAmount = 0 Then
                            cmd.ShowError(TheModCntr, ERR_GIFTCARD_HAS_NO_MONEY.ToString, "UserMessage", ERR_GIFTCARD_HAS_NO_MONEY)
                            HandlerAfterInvoke = False
                            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
                        End If
                    End If
                Else
                    HandlerAfterInvoke = True
                End If
            Case Controller.BPEController
                If eMethod = Method.Payment Then
                    If argenteaFunctionReturnObject(0).Successfull Then
                        Dim mediaRec As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
                        Dim list = From x In argenteaFunctionReturnObject(0).ListBPsEvaluated
                                   Group By value = x.Value / 100
                                   Into gp = Group, sum = Sum(x.Value / 100), count = Count(x.Value)
                                   Select New With {.v = value, .a = sum, .q = count}

                        Dim index As Integer = 1
                        For Each bp As Object In list
                            Dim qty As String = "lBP_QUANTITY_" + index.ToString
                            Dim amount As String = "lBP_AMOUNT_" + index.ToString
                            Dim value As String = "lBP_VALUE_" + index.ToString
                            If Not mediaRec.ExistField(qty) Then mediaRec.AddField(qty, DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                            If Not mediaRec.ExistField(amount) Then mediaRec.AddField(amount, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
                            If Not mediaRec.ExistField(value) Then mediaRec.AddField(value, DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)

                            mediaRec.setPropertybyName(qty, bp.q)
                            mediaRec.setPropertybyName(amount, bp.a)
                            mediaRec.setPropertybyName(value, bp.v)
                            index = index + 1
                        Next
                    End If
                End If
                HandlerAfterInvoke = True

            Case Controller.ValassisCouponController

                If eMethod = Method.ValidationValassis Then

                    If Not argenteaFunctionReturnObject(0).Successfull AndAlso
                    Not String.IsNullOrEmpty(argenteaFunctionReturnObject(0).CodeResult.ToString) AndAlso
                        MyCurrentRecord.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_MEDIA Then
                        showValassisErrorMessage(argenteaFunctionReturnObject(0), cmd)
                        Return False
                    End If

                    If Not argenteaFunctionReturnObject(0).Successfull AndAlso
                        MyCurrentRecord.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_CUSTOMER Then
                        Return False
                    End If

                    Dim szXMLNode As String = argenteaFunctionReturnObject(0).NodeXML.ToString()
                    If String.IsNullOrEmpty(szXMLNode) Then Return True

                    Dim szTerminalID As String = argenteaFunctionReturnObject(0).TerminalID
                    Dim szClientCode As String = argenteaFunctionReturnObject(0).ClientCode

                    Dim xdoc As XDocument = XDocument.Parse(szXMLNode, LoadOptions.None)

                    Dim szPosData As String = String.Empty
                    Try
                        szPosData = xdoc.Element("VCA_AUTH_RESP").Element("POS_DATA").Value.ToString()
                    Catch ex As Exception
                        szPosData = String.Empty
                    End Try
                    If xdoc.Elements.Descendants("COUPON").Count > 1 Then ReDim Preserve argenteaFunctionReturnObject(xdoc.Elements.Descendants("COUPON").Count - 1)
                    Dim dTaTotal As Decimal = taobj.GetTotal
                    For Each xEl As XElement In xdoc.Elements.Descendants("COUPON").ToList()
                        Dim index As Integer = xdoc.Elements.Descendants("COUPON").ToList().IndexOf(xEl)
                        If index <> 0 Then argenteaFunctionReturnObject(index) = New ArgenteaFunctionReturnObject

                        'Common properties
                        argenteaFunctionReturnObject(index).ArgenteaFunction = InternalArgenteaFunctionTypes.ValidationValassis
                        argenteaFunctionReturnObject(index).TerminalID = szTerminalID
                        argenteaFunctionReturnObject(index).ClientCode = szClientCode
                        argenteaFunctionReturnObject(index).PosData = szPosData

                        argenteaFunctionReturnObject(index).CouponCode = xEl.Element("COUPON_CODE").Value.ToString()
                        argenteaFunctionReturnObject(index).CodeResult = xEl.Element("RESULT_CODE").Value.ToString()
                        argenteaFunctionReturnObject(index).Amount = xEl.Element("VALUE").Value.ToString()
                        argenteaFunctionReturnObject(index).CouponTransID = xEl.Element("TRANS_ID").Value.ToString()
                        argenteaFunctionReturnObject(index).CouponCancelReason = String.Empty
                        argenteaFunctionReturnObject(index).SkuSaleNum = xEl.Element("SKU_SALE_NUM").Value.ToString()
                        argenteaFunctionReturnObject(index).szSkuList = xEl.Element("SKU_LIST").Value.ToString()
                        argenteaFunctionReturnObject(index).szMinRecpAmt = xEl.Element("MIN_RECP_AMT").Value.ToString()
                        argenteaFunctionReturnObject(index).szSkuSaleNum = xEl.Element("SKU_SALE_NUM").Value.ToString()
                        argenteaFunctionReturnObject(index).szSkuSaleMode = xEl.Element("SKU_SALE_MODE").Value.ToString()
                        argenteaFunctionReturnObject(index).szCouponType = xEl.Element("COUPON_TYPE").Value.ToString()

                        Dim szSkuList As String = argenteaFunctionReturnObject(index).szSkuList
                        Dim lMinRecpAmt As Integer = CInt(argenteaFunctionReturnObject(index).szMinRecpAmt)
                        Dim lSkuSaleNum As Integer = CInt(argenteaFunctionReturnObject(index).szSkuSaleNum)
                        Dim lSkuSaleMode As Integer = CInt(argenteaFunctionReturnObject(index).szSkuSaleMode)
                        Dim lCouponType As Integer = CInt(argenteaFunctionReturnObject(index).szCouponType)
                        If argenteaFunctionReturnObject(index).CodeResult = IValassisValidationCouponResultCode.OK AndAlso MyCurrentRecord.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_MEDIA Then
                            CheckConditionValassis(lMinRecpAmt, lCouponType, dTaTotal, szSkuList, lSkuSaleNum, lSkuSaleMode, argenteaFunctionReturnObject(index), response)
                        End If

                        If MyCurrentRecord.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_MEDIA Then
                            If Not showValassisErrorMessage(argenteaFunctionReturnObject(index), cmd) Then Return False
                        End If
                    Next
                End If
                If eMethod = Method.NotificationValassis Then
                    If MyCurrentRecord.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_MEDIA Then
                        Dim myRecMedia As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
                        If myRecMedia.dTaPaidTotal < 0 Then
                            For Each base As TPDotnet.Pos.TaBaseRec In taobj.taCollection
                                If base.sid = TPDotnet.IT.Common.Pos.Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE AndAlso
                                base.theHdr.lTaRefToCreateNmbr = MyCurrentRecord.theHdr.lTaCreateNmbr AndAlso
                                   (base.ExistField("szCouponCancelReason") AndAlso
                                   base.GetPropertybyName("szCouponCancelReason") = CInt(IValassisNotificationCancelReasonCode.TRXSUSPEND).ToString) Then
                                    If argenteaFunctionReturnObject(0).Successfull Then
                                        Dim myBaseExtRec As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(base, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                                        myBaseExtRec.szStatus = TaExternalServiceRec.ExternalServiceStatus.Deleted.ToString
                                        Dim szSkuSold As String = base.GetPropertybyName("szSkuSold")
                                        Dim lSkuSaleNum As Integer = CInt(base.GetPropertybyName("szSkuSaleNum"))
                                        Dim lArtCount As Integer = 0

                                        For Each szSku As String In szSkuSold.Split(",")
                                            Dim xElList As List(Of XElement) = New List(Of XElement)
                                            If szSku.Length = 0 Then Continue For
                                            If szSku.EndsWith("*") Then
                                                xElList = taobj.TAtoXDocument(False, 0, False).XPathSelectElements("/TAS/NEW_TA/ART_SALE[(szItemLookupCode[starts-with(.," + szSku.ToString.Replace("*", "") + ")]) and (bIsInValassisSkuSold=1 or lValassisSkuSoldCount>0)] ").ToList
                                            Else
                                                xElList = taobj.TAtoXDocument(False, 0, False).XPathSelectElements("/TAS/NEW_TA/ART_SALE[(szItemLookupCode=" + szSku.ToString + ") and (bIsInValassisSkuSold=1 or lValassisSkuSoldCount>0)]").ToList
                                            End If

                                            For Each xEl As XElement In xElList
                                                If lArtCount = lSkuSaleNum Then
                                                    Exit For
                                                End If
                                                Dim lSkuSoldCount As Integer = 0
                                                Dim lSkuSold As Integer = 0
                                                Dim lTaCreateNmbr As Integer = CInt(xEl.XPathSelectElement("Hdr/lTaCreateNmbr").Value)
                                                Dim myBaseRec As TPDotnet.Pos.TaBaseRec = taobj.GetTALine(taobj.GetPositionFromCreationNmbr(lTaCreateNmbr))
                                                If Not myBaseRec.ExistField("bIsInValassisSkuSold") Then myBaseRec.AddField("bIsInValassisSkuSold", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                                                If Not myBaseRec.ExistField("lValassisSkuSoldCount") Then myBaseRec.AddField("lValassisSkuSoldCount", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                                                Integer.TryParse(myBaseRec.GetPropertybyName("lValassisSkuSoldCount"), lSkuSoldCount)
                                                lSkuSold = Math.Min(lSkuSaleNum, lSkuSoldCount)
                                                Dim skuConditionSold As Integer = (lSkuSoldCount - lSkuSold)
                                                myBaseRec.setPropertybyName("lValassisSkuSoldCount", skuConditionSold.ToString)
                                                myBaseRec.setPropertybyName("bIsInValassisSkuSold", False)

                                                lArtCount += lSkuSold
                                            Next

                                            'Dim xEl As XElement = xElList.FirstOrDefault()
                                            'Dim lTaCreateNmbr As Integer = CInt(xEl.XPathSelectElement("Hdr/lTaCreateNmbr").Value)
                                            'Dim myBaseRec As TPDotnet.Pos.TaBaseRec = taobj.GetTALine(taobj.GetPositionFromCreationNmbr(lTaCreateNmbr))
                                            'If Not myBaseRec.ExistField("bIsInSkuSold") Then myBaseRec.AddField("bIsInSkuSold", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                                            'myBaseRec.setPropertybyName("bIsInSkuSold", False)
                                        Next



                                        HandlerAfterInvoke = True
                                    Else
                                        base.setPropertybyName("szCouponCancelReason", String.Empty)
                                        HandlerAfterInvoke = False
                                    End If
                                End If

                            Next
                        End If
                        Return HandlerAfterInvoke
                    ElseIf (MyCurrentDetailRec IsNot Nothing AndAlso MyCurrentDetailRec.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_CUSTOMER) Then
                        Dim szXMLNode As String = argenteaFunctionReturnObject(0).NodeXML.ToString()
                        If String.IsNullOrEmpty(szXMLNode) Then Return True
                        Dim xdoc As XDocument = XDocument.Parse(szXMLNode, LoadOptions.None)
                        Dim szTransID As String = xdoc.XPathSelectElement("VCA_CONF_RESP/COUPON_RESULT/TRANS_ID").Value.ToString()
                        For Each base As TPDotnet.Pos.TaBaseRec In taobj.taCollection
                            If base.sid = TPDotnet.IT.Common.Pos.Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE Then
                                Dim szCouponTransID As String = IIf(base.ExistField("szCouponTransID"), base.GetPropertybyName("szCouponTransID"), "")
                                Dim szCouponCancelReason As String = IIf(base.ExistField("szCouponCancelReason"), base.GetPropertybyName("szCouponCancelReason"), "0")

                                If (szCouponTransID = szTransID) AndAlso szCouponCancelReason <> "0" Then
                                    If argenteaFunctionReturnObject(0).Successfull Then
                                        Dim myBaseExtRec As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(base, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                                        myBaseExtRec.szStatus = TaExternalServiceRec.ExternalServiceStatus.Deleted.ToString
                                    End If
                                End If
                            End If
                        Next
                        Return True
                    End If
                End If
                HandlerAfterInvoke = True

            Case Else
                'If Not argenteaFunctionReturnObject(0).Successfull Then
                '    HandlerAfterInvoke = False
                'Else
                HandlerAfterInvoke = True
                'End If
        End Select
        Return HandlerAfterInvoke
    End Function

#Region "Valassis Handler"
    Private Function showValassisErrorMessage(argenteaFunctionReturnObject As ArgenteaFunctionReturnObject, cmd As TPDotnet.IT.Common.Pos.Common) As Boolean
        Select Case argenteaFunctionReturnObject.CodeResult
            Case IValassisCouponReturnCode.OK
                If Not String.IsNullOrEmpty(argenteaFunctionReturnObject.CouponCancelReason) Then
                    Dim defaultError As String = CType([Enum].Parse(GetType(IValassisNotificationCancelReasonCode), argenteaFunctionReturnObject.CouponCancelReason), IValassisNotificationCancelReasonCode).ToString()
                    cmd.ShowError(TheModCntr, defaultError, "LevelITCommonArgenteaCouponValassisCancelError_" + CInt(argenteaFunctionReturnObject.CouponCancelReason).ToString)
                    Return True
                End If
            Case ""
                cmd.ShowError(TheModCntr, argenteaFunctionReturnObject.Description, "LevelITCommonArgenteaValassisError")
                Return False
            Case Else
                Dim defaultError As String = CType([Enum].Parse(GetType(IValassisValidationCouponResultCode), argenteaFunctionReturnObject.CodeResult), IValassisValidationCouponResultCode).ToString()
                cmd.ShowError(TheModCntr, defaultError, "LevelITCommonArgenteaValassisError_" + argenteaFunctionReturnObject.CodeResult)
                Return False
        End Select
        Return True
    End Function
    Private Sub CheckConditionValassis(lMinRecpAmt As Integer, lCouponType As Integer, ByRef dTaTotal As Decimal, szSkuList As String, lSkuSaleNum As Integer, lSkuSaleMode As Integer, argenteaFunctionReturnObject As ArgenteaFunctionReturnObject, response As ArgenteaResponse)

        Dim szSkuSold As String = String.Empty
        Dim szSkuNotSold As String = String.Empty
        If Not Common.CheckConditionCouponTypeValassis(lCouponType) Then
            argenteaFunctionReturnObject.CouponCancelReason = CInt(IValassisNotificationCancelReasonCode.COUPONTYPENOTMANAGED).ToString
            argenteaFunctionReturnObject.CodeResult = CInt(IValassisValidationCouponResultCode.OK).ToString
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            Exit Sub
        End If

        If Not Common.CheckConditionTypeValassis(lMinRecpAmt) Then
            argenteaFunctionReturnObject.CouponCancelReason = CInt(IValassisNotificationCancelReasonCode.COUPONTYPENOTMANAGED).ToString
            argenteaFunctionReturnObject.CodeResult = CInt(IValassisValidationCouponResultCode.OK).ToString
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            Exit Sub
        End If

        If Not Common.CheckConditionSkuValassis(szSkuList, lSkuSaleNum, lSkuSaleMode, taobj, szSkuSold, szSkuNotSold) Then
            argenteaFunctionReturnObject.CouponCancelReason = CInt(IValassisNotificationCancelReasonCode.SKUNOTSOLD).ToString
            argenteaFunctionReturnObject.CodeResult = CInt(IValassisValidationCouponResultCode.OK).ToString
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            Exit Sub
        Else
            argenteaFunctionReturnObject.SkuSold = szSkuSold
            argenteaFunctionReturnObject.SkuList = szSkuNotSold
        End If

        If Not Common.CheckConditionAmountValassis(dTaTotal, MyCurrentRecord, argenteaFunctionReturnObject.Amount) Then
            argenteaFunctionReturnObject.CouponCancelReason = CInt(IValassisNotificationCancelReasonCode.COUPONGREATHERTHENTOTAMOUNT).ToString
            argenteaFunctionReturnObject.CodeResult = CInt(IValassisValidationCouponResultCode.OK).ToString
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
            Exit Sub
        End If
    End Sub

#End Region
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
                If response.ReturnCode <> ArgenteaFunctionsReturnCode.KO Then
                    szReceipt = argenteaFunctionReturnObject(i).Receipt
                End If
            End If


            Dim szBarcode As String = String.Empty
            If response.ExistProperty("szBarcode") Then
                szBarcode = response.GetProperty("szBarcode")
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).CouponCode) Then
                szBarcode = argenteaFunctionReturnObject(i).CouponCode.ToString
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
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).CodeResult) Then
                If Not TaExternalServiceRec.ExistField("szResultCode") Then TaExternalServiceRec.AddField("szResultCode", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                TaExternalServiceRec.setPropertybyName("szResultCode", argenteaFunctionReturnObject(i).CodeResult.ToString)
            End If
            If Not String.IsNullOrEmpty(szBarcode) Then
                If Not TaExternalServiceRec.ExistField("szBarcode") Then TaExternalServiceRec.AddField("szBarcode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szBarcode", szBarcode.ToString)
            End If

            'Ony for Coupon Valassis
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).CouponCancelReason) Then
                If Not TaExternalServiceRec.ExistField("szCouponCancelReason") Then TaExternalServiceRec.AddField("szCouponCancelReason", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                TaExternalServiceRec.setPropertybyName("szCouponCancelReason", argenteaFunctionReturnObject(i).CouponCancelReason.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).CouponTransID) Then
                If Not TaExternalServiceRec.ExistField("szCouponTransID") Then TaExternalServiceRec.AddField("szCouponTransID", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                TaExternalServiceRec.setPropertybyName("szCouponTransID", argenteaFunctionReturnObject(i).CouponTransID.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).SkuSold) Then
                If Not TaExternalServiceRec.ExistField("szSkuSold") Then TaExternalServiceRec.AddField("szSkuSold", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szSkuSold", argenteaFunctionReturnObject(i).SkuSold.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).SkuList) Then
                If Not TaExternalServiceRec.ExistField("szSkuList") Then TaExternalServiceRec.AddField("szSkuList", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szSkuList", argenteaFunctionReturnObject(i).SkuList.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).SkuSaleNum) Then
                If Not TaExternalServiceRec.ExistField("szSkuSaleNum") Then TaExternalServiceRec.AddField("szSkuSaleNum", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szSkuSaleNum", argenteaFunctionReturnObject(i).SkuSaleNum.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).ClientCode) Then
                If Not TaExternalServiceRec.ExistField("szClientCode") Then TaExternalServiceRec.AddField("szClientCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szClientCode", argenteaFunctionReturnObject(i).ClientCode.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).PosData) Then
                If Not TaExternalServiceRec.ExistField("szPosData") Then TaExternalServiceRec.AddField("szPosData", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szPosData", argenteaFunctionReturnObject(i).PosData.ToString)
            End If

            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).szSkuList) Then
                If Not TaExternalServiceRec.ExistField("szSkuList") Then TaExternalServiceRec.AddField("szSkuList", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szSkuList", argenteaFunctionReturnObject(i).szSkuList.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).szMinRecpAmt) Then
                If Not TaExternalServiceRec.ExistField("szMinRecpAmt") Then TaExternalServiceRec.AddField("szMinRecpAmt", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szMinRecpAmt", argenteaFunctionReturnObject(i).szMinRecpAmt.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).szSkuSaleNum) Then
                If Not TaExternalServiceRec.ExistField("szSkuSaleNum") Then TaExternalServiceRec.AddField("szSkuSaleNum", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szSkuSaleNum", argenteaFunctionReturnObject(i).szSkuSaleNum.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).szSkuSaleMode) Then
                If Not TaExternalServiceRec.ExistField("szSkuSaleMode") Then TaExternalServiceRec.AddField("szSkuSaleMode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szSkuSaleMode", argenteaFunctionReturnObject(i).szSkuSaleMode.ToString)
            End If
            If Not String.IsNullOrEmpty(argenteaFunctionReturnObject(i).szCouponType) Then
                If Not TaExternalServiceRec.ExistField("szCouponType") Then TaExternalServiceRec.AddField("szCouponType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                TaExternalServiceRec.setPropertybyName("szCouponType", argenteaFunctionReturnObject(i).szCouponType.ToString)
            End If

            If Not TaExternalServiceRec.ExistField("szFunctionType") Then TaExternalServiceRec.AddField("szFunctionType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            TaExternalServiceRec.setPropertybyName("szFunctionType", response.FunctionType.ToString)

            With TaExternalServiceRec
                ' only for coupon valassis. I need to stored the external service even if the coupon is not used
                If String.IsNullOrEmpty(argenteaFunctionReturnObject(i).CouponCancelReason) Then .theHdr.lTaRefToCreateNmbr = MyCurrentRecord.theHdr.lTaCreateNmbr
                If Not String.IsNullOrEmpty(szReceipt) Then .szReceipt = szReceipt
                .szServiceType = GetOperationType(ExternalID)
                .lCopies = paramArg.Copies
                .szStatus = GetStatus(eMethod).ToString()
                .bSuppressLogo = paramCommon.SuppressLogo
                .bPrintReceipt = paramArg.PrintWithinTA
                .setPropertybyName("lAmount", lAmount.ToString())
                .setPropertybyName("szTransactionID", response.TransactionID.ToString())
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

#End Region

#Region "Public Function"
    Public Function ElectronicFundsTransferPay_Pay(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferPay.Pay
        Return Execute(Method.Payment, Parameters)
    End Function
    Public Function ElectronicFundsTransferVoid_Void(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferVoid.Void
        Return Execute(Method.Void, Parameters)
    End Function
    Public Function ElectronicFundsTransferClosure_Closure(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferClosure.Closure
        Return Execute(Method.Closure, Parameters)
    End Function
    Public Function ElectronicFundsTransferTotals_Totals(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode Implements IElectronicFundsTransferTotals.Totals
        Return Execute(Method.Totals, Parameters)
    End Function
    Public Function IElectronicMealVoucherPay_Pay(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherPay.Pay
        Return Execute(Method.Payment, Parameters)
    End Function
    Public Function IElectronicMealVoucherVoid_Void(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherVoid.Void
        Return Execute(Method.Void, Parameters)
    End Function
    Public Function CheckPhoneRecharge(ByRef Parameters As Dictionary(Of String, Object)) As IPhoneRechargeReturnCode Implements IPhoneRechargeActivationPreCheck.CheckPhoneRecharge
        Return Execute(Method.CheckPhoneRecharge, Parameters)
    End Function
    Public Function ActivatePhoneRecharge(ByRef Parameters As Dictionary(Of String, Object)) As IPhoneRechargeReturnCode Implements IPhoneRechargeActivation.ActivatePhoneRecharge
        Return Execute(Method.ActivatePhoneRecharge, Parameters)
    End Function
    Public Function ActivationExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardActivation.ActivationExternalGiftCard
        Return Execute(Method.ActivationExternalGiftCard, Parameters)
    End Function
    Public Function DeActivationExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardDeActivation.DeActivationExternalGiftCard
        Return Execute(Method.DeActivationExternalGiftCard, Parameters)
    End Function
    Public Function ConfirmExternalGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IExternalGiftCardReturnCode Implements IExternalGiftCardConfirm.ConfirmExternalGiftCard
        Return Execute(Method.ConfirmExternalGiftCard, Parameters)
    End Function
    Public Function CheckGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardActivationPreCheck.CheckGiftCard
        Return Execute(Method.CheckGiftCard, Parameters)
    End Function
    Public Function ActivateGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardActivation.ActivateGiftCard
        Return Execute(Method.ActivateGiftCard, Parameters)
    End Function
    Public Function GiftCardBalanceInquiry(ByRef Parameters As Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardBalanceInquiry.GiftCardBalanceInquiry
        eController = Controller.GiftCardController
        Return Execute(Method.GiftCardBalanceInquiry, Parameters)
    End Function
    Public Function RedeemGiftCard(ByRef Parameters As Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardRedeem.RedeemGiftCard
        Return Execute(Method.RedeemGiftCard, Parameters)
    End Function
    Public Function GiftCardCancellation(ByRef Parameters As Dictionary(Of String, Object)) As IGiftCardReturnCode Implements IGiftCardCancellationPayment.GiftCardCancellation
        Return Execute(Method.GiftCardCancellation, Parameters)
    End Function
    Public Function Balance(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherBalance.Balance
        eController = Controller.BPEController
        Return Execute(Method.Balance, Parameters)
    End Function
    Public Function IElectronicMealVoucherClosure_Closure(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherClosure.Closure
        eController = Controller.BPEController
        Return Execute(Method.Closure, Parameters)
    End Function
    Public Function IElectronicMealVoucherTotals_Totals(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherTotals.Totals
        Return Execute(Method.Totals, Parameters)
    End Function
    Public Function ValidationValassis(ByRef Parameters As Dictionary(Of String, Object)) As IValassisCouponReturnCode Implements IValassisCouponValidation.ValidationValassis
        eController = Controller.ValassisCouponController
        Return Execute(Method.ValidationValassis, Parameters)
    End Function
    Public Function NotificationValassis(ByRef Parameters As Dictionary(Of String, Object)) As IValassisCouponReturnCode Implements IValassisCouponNotification.NotificationValassis
        eController = Controller.ValassisCouponController
        Return Execute(Method.NotificationValassis, Parameters)
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
        Return IElectronicMealVoucherReturnCode.OK
    End Function
    Public Function IElectronicMealVoucherBalance_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherBalance.Check
        Return IElectronicMealVoucherReturnCode.OK
    End Function
    Public Function IElectronicMealVoucherClosure_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherClosure.Check
        Return IElectronicMealVoucherReturnCode.OK

    End Function
    Private Function IElectronicMealVoucherTotals_Check(ByRef Parameters As Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode Implements IElectronicMealVoucherTotals.Check
        Return IElectronicMealVoucherReturnCode.OK
    End Function
#End Region
#End Region

End Class

Public Class CodeResult
    Public Const UnderFunded As String = "0202"
End Class
