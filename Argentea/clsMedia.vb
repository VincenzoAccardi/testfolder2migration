Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.IT.Common.Pos
Imports System.Collections.Generic
Imports TPDotnet.Services.Rounding

Public Class clsMedia
    Inherits TPDotnet.IT.Common.Pos.clsMedia

#Region "Documentation"
    ' ********** ********** ********** **********
    ' Internal media class for Argnetea EMV handling
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Basiglio, 2014, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Overridden methods"
    Protected Overrides Function DoSpecialHandling4CreditCardsOnline(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec, ByRef MyTaMediaMemberDetailRec As TPDotnet.Pos.TaMediaMemberDetailRec) As Boolean

        Dim funcName As String = "DoSpecialHandling4CreditCardsOnline"
        Dim Ret As Integer
        Dim MyTaMediaSwap As TaMediaMemberSwap
        Dim bIsVoidReceipt As Boolean
        Dim bIsMediaCorrect As Boolean

        DoSpecialHandling4CreditCardsOnline = False

        Try
            LOG_FuncStart(getLocationString(funcName))

            GetStati(taobj, TheModCntr, MyTaMediaRec)
            If Me.m_bIsSwapMedia Then
                LOG_Debug(getLocationString(funcName), "We are in a swap media receipt. Just return true.")
                DoSpecialHandling4CreditCardsOnline = True
                Exit Function
            End If

            If MyTaMediaRec.dTaPaidTotal < 0.0 OrElse MyTaMediaRec.dReturn < 0.0 Then
                'for performance reasons 
                For Ret = 1 To taobj.GetNmbrofRecs
                    If taobj.GetTALine(Ret).sid = PosDef.TARecTypes.iTA_VOID_RECEIPT Then
                        bIsVoidReceipt = True
                        LOG_Debug(getLocationString(funcName), "We are in a void receipt")
                        Exit For
                    End If
                    If taobj.GetTALine(Ret).sid = PosDef.TARecTypes.iTA_MEDIAMEMBER_SWAP Then
                        MyTaMediaSwap = taobj.GetTALine(Ret)
                        If MyTaMediaSwap.sFunction = PosDef.TATaMediaMemberSwap.iCORRECT Then
                            bIsMediaCorrect = True
                            LOG_Debug(getLocationString(funcName), "We are in a media correction")
                            MyTaMediaSwap = Nothing
                        End If
                        Exit For
                    End If
                Next Ret
            End If

            If MyTaMediaRec.dTaPaidTotal > 0.0# And MyTaMediaRec.dReturn = 0.0# Then

                ' payment
                LOG_Debug(getLocationString(funcName), "before calling payment")
                If MyTaMediaRec.PAYMENTinMedia.szExternalID.Equals(TPDotnet.IT.Common.Pos.EFT.ADVController.PaymentAdv.JIFFY.ToString()) OrElse
                    MyTaMediaRec.PAYMENTinMedia.szExternalID.Equals(TPDotnet.IT.Common.Pos.EFT.ADVController.PaymentAdv.SATISPAY.ToString()) OrElse
                    MyTaMediaRec.PAYMENTinMedia.szExternalID.Equals(TPDotnet.IT.Common.Pos.EFT.ADVController.PaymentAdv.BITCOIN.ToString()) Then
                    DoSpecialHandling4CreditCardsOnline = TPDotnet.IT.Common.Pos.EFT.ADVController.Instance.Payment(taobj, TheModCntr, MyTaMediaRec, MyTaMediaMemberDetailRec)
                Else
                    DoSpecialHandling4CreditCardsOnline = TPDotnet.IT.Common.Pos.EFT.EFTController.Instance.Payment(taobj, TheModCntr, MyTaMediaRec, MyTaMediaMemberDetailRec)
                End If
            ElseIf bIsVoidReceipt OrElse bIsMediaCorrect _
            OrElse MyTaMediaRec.dTaPaidTotal < 0 Then

                If MyTaMediaRec.PAYMENTinMedia.szSpecialUsageVoidReceipt = MM_VOIDRECEIPT_SHOULD_BE_VOIDED Then

                    LOG_Debug(getLocationString(funcName), "Don't call the Void on the EFT device because the media is configured as should be voided.")
                    DoSpecialHandling4CreditCardsOnline = True
                    Exit Function

                End If
                LOG_Debug(getLocationString(funcName), "before calling Void")

                If MyTaMediaRec.PAYMENTinMedia.szExternalID.Equals(TPDotnet.IT.Common.Pos.EFT.ADVController.PaymentAdv.JIFFY.ToString()) OrElse
                    MyTaMediaRec.PAYMENTinMedia.szExternalID.Equals(TPDotnet.IT.Common.Pos.EFT.ADVController.PaymentAdv.SATISPAY.ToString()) OrElse
                    MyTaMediaRec.PAYMENTinMedia.szExternalID.Equals(TPDotnet.IT.Common.Pos.EFT.ADVController.PaymentAdv.BITCOIN.ToString()) Then
                    DoSpecialHandling4CreditCardsOnline = TPDotnet.IT.Common.Pos.EFT.ADVController.Instance.Void(taobj, TheModCntr)
                Else
                    ' line/immediate void
                    DoSpecialHandling4CreditCardsOnline = TPDotnet.IT.Common.Pos.EFT.EFTController.Instance.Void(taobj, TheModCntr)
                End If
            End If

            If Not DoSpecialHandling4CreditCardsOnline Then
                TPMsgBox(TPDotnet.Pos.PosDef.TARMessageTypes.TPERROR, getPosTxtNew(TheModCntr.contxt, "UserMessage", TXT_EFT_PAYMENT_ABORT), TXT_EFT_PAYMENT_ABORT, TheModCntr, "UserMessage")
            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString(funcName), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString(funcName), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function returns " & DoSpecialHandling4CreditCardsOnline.ToString)
        End Try

    End Function

    Protected Overrides Function DoSpecialHandling4Vouchers1(ByRef taobj As TPDotnet.Pos.TA, ByRef theModCntr As TPDotnet.Pos.ModCntr, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec, ByRef MyTaMediaMemberDetailRec As TPDotnet.Pos.TaMediaMemberDetailRec) As Boolean
        DoSpecialHandling4Vouchers1 = False

        Dim funcName As String = "DoSpecialHandling4Vouchers1"
        Dim MyCust As CUST
        Dim ret As IGiftCardReturnCode
        Dim dBalance As System.Decimal = New System.Decimal(0)
        Dim szReceipt As String = String.Empty

        Try
            LOG_FuncStart(getLocationString(funcName))

            If String.IsNullOrEmpty(MyTaMediaRec.PAYMENTinMedia.szExternalID) OrElse
                Not String.Equals(MyTaMediaRec.PAYMENTinMedia.szExternalID, "GIFTCARDARGENTEA", StringComparison.InvariantCultureIgnoreCase) Then
                LOG_Error(getLocationString(funcName), "MediaMember " & MyTaMediaRec.PAYMENTinMedia.lMediaMember & " is not configured as GIFTCARDARGENTEA")
                Exit Function
            End If

            ' check the currently status (void correct swap e.g)
            ' ==================================================
            GetStati(taobj, theModCntr, MyTaMediaRec)

            If MyTaMediaRec.dTaPaidTotal > 0.0# AndAlso MyTaMediaRec.dReturn = 0.0# Then

                ' payment
                LOG_Debug(getLocationString(funcName), "before calling payment")

                If String.IsNullOrEmpty(MyTaMediaRec.szBarcode) Then

                    DoSpecialHandling4Vouchers1 = RedeemVoucher(taobj, theModCntr, MyTaMediaRec, MyTaMediaMemberDetailRec)
                    If DoSpecialHandling4Vouchers1 = False Then
                        Exit Function
                    End If

                    MyTaMediaRec.szBarcode = m_szSerialNmbr

                End If

                If theModCntr.bCalledFromWebService OrElse theModCntr.bExternalDialog Then
                    If String.IsNullOrEmpty(MyTaMediaRec.szBarcode) Then
                        TPMsgBox(PosDef.TARMessageTypes.TPINFORMATION,
                     getPosTxtNew(theModCntr.contxt, "Message", TEXT_INVALID_BARCODE),
                     TEXT_INVALID_BARCODE, theModCntr, "Message")
                        DoSpecialHandling4Vouchers1 = False
                        Exit Function
                    End If
                End If

                If Not String.IsNullOrEmpty(MyTaMediaRec.szBarcode) Then

                    ret = InquiryGiftCard(theModCntr, taobj, MyTaMediaRec.szBarcode, dBalance, szReceipt)
                    Select Case ret

                        Case IGiftCardReturnCode.KO,
                             IGiftCardReturnCode.KO_SKIP_STANDARD,
                             IGiftCardReturnCode.OK_SKIP_STANDARD

                            DoSpecialHandling4Vouchers1 = False
                            Exit Function

                    End Select

                    If Not dBalance > 0 Then
                        ' Gift card does not contains money!
                        TPMsgBox(PosDef.TARMessageTypes.TPERROR, getPosTxtNew((theModCntr.contxt), "UserMessage", ERR_GIFTCARD_HAS_NO_MONEY), ERR_GIFTCARD_HAS_NO_MONEY, theModCntr, "UserMessage")
                        DoSpecialHandling4Vouchers1 = False
                        Exit Function
                    End If

                    MyTaMediaRec.dTaPaidTotal = Math.Min(MyTaMediaRec.dTaPaidTotal, dBalance)
                    MyTaMediaRec.dTaPaid = MyTaMediaRec.dTaPaidTotal / MyTaMediaRec.dTaQty
                    MyTaMediaRec.dPaidForeignCurrTotal = MyTaMediaRec.dTaPaidTotal
                    MyTaMediaRec.dPaidForeignCurr = MyTaMediaRec.dTaPaidTotal

                End If

                ' we use the standard RedeemVoucher function because it requires the input for:
                '   - barcode
                '   - amount
                DoSpecialHandling4Vouchers1 = RedeemVoucher(taobj, theModCntr, MyTaMediaRec, MyTaMediaMemberDetailRec)
                If DoSpecialHandling4Vouchers1 = False Then
                    Exit Function
                End If

                If Not MyTaMediaRec.PAYMENTinMedia.bPayOverpaid AndAlso Math.Round(taobj.GetReturnValue, theModCntr.sLocCurrDecNmbr) > 0 Then
                    TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                         String.Concat(getPosTxtNew(theModCntr.contxt, "Message", ERR_NO_OVERPAY), " ", MyTaMediaRec.dTaPaidTotal.ToString),
                         ERR_NO_OVERPAY, theModCntr, "Message")
                    LOG_Info(getLocationString("DoPostChecks"), "overpayment for that payment not allowed")
                    DoSpecialHandling4Vouchers1 = False
                    Exit Function
                End If

                If MyTaMediaRec.dTaPaidTotal > dBalance Then
                    ' Gift card does not contains enough money!
                    DoSpecialHandling4Vouchers1 = False
                    Exit Function
                End If

                ret = CheckRedeemGiftCard(theModCntr, taobj, MyTaMediaRec)
                Select Case ret

                    Case IGiftCardReturnCode.KO,
                         IGiftCardReturnCode.KO_SKIP_STANDARD,
                         IGiftCardReturnCode.OK_SKIP_STANDARD

                        DoSpecialHandling4Vouchers1 = False
                        Exit Function

                End Select

            ElseIf m_bIsMediaCorrect _
                OrElse m_bIsSwapMedia _
                OrElse m_bIsVoidReceipt Then

                ' we can do nothing in this case, only report to chashier 
                TPMsgBox(TPDotnet.Pos.PosDef.TARMessageTypes.TPWARNING, "We cannot really void this media member!", 0, theModCntr, "Message")

                DoSpecialHandling4Vouchers1 = True
                Exit Function

            ElseIf MyTaMediaRec.dTaPaidTotal < 0 Then

                ' line/immediate void
                LOG_Debug(getLocationString(funcName), "Media line Void")

                ret = GiftCardCancellation(theModCntr, taobj, MyTaMediaRec)
                Select Case ret

                    Case IGiftCardReturnCode.KO,
                         IGiftCardReturnCode.KO_SKIP_STANDARD,
                         IGiftCardReturnCode.OK_SKIP_STANDARD

                        DoSpecialHandling4Vouchers1 = False
                        Exit Function

                End Select

                DoSpecialHandling4Vouchers1 = True
                Exit Function

            Else

                ' line/immediate void
                LOG_Debug(getLocationString(funcName), "before calling Void")

                ' we can do nothing in this case, only report to chashier 
                TPMsgBox(TPDotnet.Pos.PosDef.TARMessageTypes.TPWARNING, "We cannot really void this media member!", 0, theModCntr, "Message")

                DoSpecialHandling4Vouchers1 = True
                Exit Function

            End If

            If m_MyTaSerialized Is Nothing Then
                m_MyTaSerialized = taobj.CreateTaObject(PosDef.TARecTypes.iTA_SERIALIZED)
                MyCust = taobj.CreateTaSubObject("KEY_CUSTOMER")
                taobj.getCustInfos(MyCust)
                If Not MyCust Is Nothing Then
                    m_MyTaSerialized.szCustomerID = MyCust.szCustomerID
                End If
                MyCust = Nothing
                m_MyTaSerialized.szPosItemID = ""
                m_MyTaSerialized.szComment = ""
                m_MyTaSerialized.lMediaMember = MyTaMediaRec.PAYMENTinMedia.lMediaMember
                m_MyTaSerialized.lMediaNmbr = MyTaMediaRec.PAYMENTinMedia.lMediaNmbr
                m_MyTaSerialized.szSerialNmbr = m_szSerialNmbr
                m_MyTaSerialized.dAmount = MyTaMediaRec.dTaPaidTotal
                m_MyTaSerialized.dPayExchRate = MyTaMediaRec.PAYMENTinMedia.dPayExchRate
                m_MyTaSerialized.dAmountForeign = MyTaMediaRec.dPaidForeignCurr * MyTaMediaRec.dTaQty
                m_MyTaSerialized.szSerializeTypeCode = MyTaMediaRec.PAYMENTinMedia.szSerializeTypeCode
                m_MyTaSerialized.szTypeCode = SERIALNMBR_MEDIAMEMBER
                m_MyTaSerialized.szStatus = SERIAL_REDEEMED
                m_MyTaSerialized.lOperatorID = theModCntr.lActOperatorID

                If m_MyTaSerialized.szSerialNmbr.Length > 0 Then
                    m_MyTaSerialized.szTypeCode = SERIALNMBR_MEDIAMEMBER
                    m_MyTaSerialized.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr

                    taobj.Add(m_MyTaSerialized)
                    If Not MyTaMediaMemberDetailRec Is Nothing Then
                        MyTaMediaMemberDetailRec.szSerialNmbr = m_MyTaSerialized.szSerialNmbr
                        MyTaMediaMemberDetailRec.szExternalID = MyTaMediaRec.PAYMENTinMedia.szExternalID
                    End If
                End If
            End If

            m_MyTaSerialized = Nothing

            'we have to call TARefresh because MyTaMediaRec were changed
            'while it already was added to TA
            taobj.TARefresh()
            DoSpecialHandling4Vouchers1 = True

            Exit Function

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DoSpecialHandling4Vouchers1"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DoSpecialHandling4Vouchers1"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DoSpecialHandling4Vouchers1"), String.Concat("Function DoSpecialHandling4Vouchers1 returns ", DoSpecialHandling4Vouchers1.ToString))
        End Try
    End Function

    Protected Overrides Function GetDialogInputs4Redemption(ByVal taobj As TPDotnet.Pos.TA, ByVal TheModCntr As TPDotnet.Pos.ModCntr,
                                                         ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec) As Boolean

        Dim bRet As Boolean

        GetDialogInputs4Redemption = False

        Try

            LOG_Debug(getLocationString("GetDialogInputs4Redemption"), "before starting internal dialog")

            m_szSerialNmbr = MyTaMediaRec.szSerialNmbr
            m_szSerializeTypeCode = MyTaMediaRec.PAYMENTinMedia.szSerializeTypeCode

            m_szAmount = MyTaMediaRec.dPaidForeignCurr.ToString(GetFormatString4Media(TheModCntr, MyTaMediaRec))
            If m_bIsMediaCorrect = True Then
                'the amount is not allowed to be changed
                m_szFixAmount = m_szAmount
            Else
                m_szFixAmount = MyTaMediaRec.PAYMENTinMedia.dFixAmount.ToString(GetFormatString4Media(TheModCntr, MyTaMediaRec))
            End If
            m_objPayment = MyTaMediaRec.PAYMENTinMedia
            m_szMediaDesc = MyTaMediaRec.PAYMENTinMedia.szDesc

            LoadDlgForm(taobj, TheModCntr)
            If MyTaMediaRec.PAYMENTinMedia.lPayType = PayTypes.iMEDIA_DISCOUNT Then
                m_FormVoucher.txt_VoucherAmount.MaxValue = taobj.GetDiscountableTotal()
                m_FormVoucher.txt_VoucherAmount.IsNumeric = True
                m_FormVoucher.txt_VoucherAmount.UseSigns = False
            End If
            m_FormVoucher.txt_VoucherAmount.NumberDecimalDigits = MyTaMediaRec.PAYMENTinMedia.lPayDecNmbr
            m_FormVoucher.lbl_VoucherDesc.Text = MyTaMediaRec.PAYMENTinMedia.szDesc
            m_FormVoucher.bDialogActive = True
            If Not String.IsNullOrEmpty(m_FormVoucher.txt_VoucherNmbr.Text) Then
                m_FormVoucher.txt_VoucherAmount.SelectionLength = m_FormVoucher.txt_VoucherAmount.Text.Length
                m_FormVoucher.txt_VoucherAmount.Focus()
            End If
            ' dialog is running, we will wait until Dialog is completed
            'Wait until dialog is completed
            Do While m_FormVoucher.bDialogActive = True
                System.Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()
            Loop

            LOG_Debug(getLocationString("GetDialogInputs4Redemption"), "after Dialog")

            'Load dialog return values
            If CBool(m_FormVoucher.Tag) = True Then
                bRet = True
            Else
                bRet = False
            End If
            m_szSerialNmbr = m_FormVoucher.txt_VoucherNmbr.Text

            m_szAmount = m_FormVoucher.txt_VoucherAmount.Text
            If bRet = False Then
                Return False
            End If

            'we take over the values from the Dialog
            MyTaMediaRec.dTaQty = 1
            MyTaMediaRec.dPaidForeignCurr = Convert.ToDecimal(m_szAmount)
            MyTaMediaRec.dPaidForeignCurrTotal = MyTaMediaRec.dPaidForeignCurr
            MyTaMediaRec.dTaPaid = Rounding.dRounding(System.Math.Abs(MyTaMediaRec.dPaidForeignCurr) / MyTaMediaRec.PAYMENTinMedia.dPayExchRate,
                                                TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_UP,
                                                MyTaMediaRec.PAYMENTinMedia.lPaySDOC,
                                                MyTaMediaRec.PAYMENTinMedia.lPayDecNmbr)
            MyTaMediaRec.dTaPaidTotal = MyTaMediaRec.dTaPaid
            MyTaMediaRec.szSerialNmbr = m_szSerialNmbr
            If Not m_FormVoucher.szBarcode Is Nothing Then
                MyTaMediaRec.szBarcode = m_FormVoucher.szBarcode
            End If

            Return bRet

        Catch ex As Exception
            LOG_Error(getLocationString("GetDialogInputs4Redemption"), ex)
        Finally
            Try
                LOG_FuncExit(getLocationString("GetDialogInputs4Redemption"), String.Concat("Function GetDialogInputs4Redemption returns ", GetDialogInputs4Redemption.ToString))
                If Not m_FormVoucher Is Nothing Then
                    TheModCntr.DialogActiv = False
                    TheModCntr.DialogFormName = ""
                    TheModCntr.EndForm()
                    TheModCntr.SetFuncKeys((True))
                    m_FormVoucher.Close()
                    m_FormVoucher = Nothing
                End If
            Catch ix As Exception
                LOG_Error(getLocationString("GetDialogInputs4Redemption"), ix)
            End Try
        End Try

    End Function

    ''' <summary>
    '''     Gestione spcializzata al pagamento  dietro
    '''     i tipi di Buono Pasto in gestione Argentea
    '''     
    '''       BPC ->  Buono Pasto Cartaceo (Tasto su pagamenti in Tab dedicata con Cartaceo)
    '''       BPE ->  Buono Pasto Elettronico (Tasto su pagamenti in Tab dedicata con Elettronico)
    '''
    ''' </summary>
    ''' <param name="taobj">La Transaction in corso</param>
    ''' <param name="theModCntr">Il Controller a riporto</param>
    ''' <param name="MyTaMediaRec">La MediaRec iniziale come Pagato (rimossa e aggioranta con i nuovi mediarec all'ok del pagato da qui)</param>
    ''' <param name="MyTaMediaMemberDetailRec">La MediaRec di riporto alla funzione sulla classe base (non utilizzata qui)</param>
    ''' <returns></returns>
    Protected Overrides Function DoSpecialHandling4Vouchers2(ByRef taobj As TPDotnet.Pos.TA, ByRef theModCntr As ModCntr, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec, ByRef MyTaMediaMemberDetailRec As TaMediaMemberDetailRec) As Boolean
        Dim funcName As String = "DoSpecialHandling4Vouchers2"

        'Dim dPayableAmount As Decimal = 0
        Dim m_TAMediaRec As TPDotnet.Pos.TaMediaRec = MyTaMediaRec

        Try
            'DoSpecialHandling4Vouchers2 = False
            LOG_FuncStart(getLocationString(funcName))

            ' Sulla Base opero per il flow di default
            ' su tutti gli altri casi di gestione.
            '
            '   ->  Voucher2 tutti i tipi

            If String.IsNullOrEmpty(MyTaMediaRec.PAYMENTinMedia.szExternalID.Trim) Then
                Return MyBase.DoSpecialHandling4Vouchers2(taobj, theModCntr, MyTaMediaRec, MyTaMediaMemberDetailRec)
            End If

            'Dim m_PayableAmout As Decimal = Math.Min(GetPayableAmount(m_Transaction, m_ModCntr) - (m_Transaction.GetPayedValue - m_TAMediaRec.dTaPaidTotal), m_Transaction.GetTotal)

            ' Altrimenti opero per questo flow specifico
            ' rispetto  alla  gestione  di cui preposto.
            '
            '   -> In questo caso gestiamo.:
            '       BPC ->  Buono Pasto Cartaceo (Tasto su pagamenti in Tab dedicata con Cartaceo)
            '       BPE ->  Buono Pasto Elettronico (Tasto su pagamenti in Tab dedicata con Elettronico)
            Dim _DoSpecialHandling4Vouchers2 As IBPReturnCode = IBPReturnCode.KO

            If MyTaMediaRec.PAYMENTinMedia.szExternalID.Trim.ToUpper = "BPC" Then
                ' Ticket Cartaceo Basato su Service Argentea
                _DoSpecialHandling4Vouchers2 = ProcessBPCartaceo(theModCntr, taobj, MyTaMediaRec, MyTaMediaMemberDetailRec)
            Else ' MyTaMediaRec.PAYMENTinMedia.szExternalID.Trim.ToUpper = "BPE" Then
                'Ticket Elettronico Basato su POS Argentea
                _DoSpecialHandling4Vouchers2 = ProcessBCElettronico(MyTaMediaRec, MyTaMediaMemberDetailRec)
            End If

            ' Se il risultato del Processo è stato
            ' concluso in modo corretto faccio  il
            ' committ dell' operazione  altrimenti
            ' esco senza committ ( OK = True KO = False)
            If _DoSpecialHandling4Vouchers2 = IBPReturnCode.OK Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            LOG_ErrorInTry(getLocationString(funcName), ex)
        Finally
            LOG_FuncExit(getLocationString(funcName), "Function " + funcName + " returns " & DoSpecialHandling4Vouchers2.ToString)
        End Try


    End Function

#End Region

#Region "IGiftCardRedeemPreCheck"

    Public Function CheckRedeemGiftCard(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec) As IGiftCardReturnCode
        CheckRedeemGiftCard = IGiftCardReturnCode.OK
        Dim funcName As String = "CheckRedeemGiftCard"
        Dim handler As IGiftCardRedeemPreCheck

        Try

            handler = createPosModelObject(Of IGiftCardRedeemPreCheck)(TheModCntr, "GiftCardController", 0, False)
            If handler Is Nothing Then
                ' gift card handler is not defined into the database
                CheckRedeemGiftCard = IGiftCardReturnCode.KO
                Exit Function
            End If

            CheckRedeemGiftCard = handler.CheckRedeemGiftCard(New Dictionary(Of String, Object) From {
                                                              {"Controller", TheModCntr},
                                                              {"Transaction", taobj},
                                                              {"MediaRecord", MyTaMediaRec}
                                                          })

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

#Region "IGiftCardBalanceInquiry"

    Public Function InquiryGiftCard(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA, ByRef szBarcode As String, ByRef dBalance As System.Decimal, ByRef szReceipt As String) As IGiftCardReturnCode
        InquiryGiftCard = IGiftCardReturnCode.OK
        Dim funcName As String = "CheckRedeemGiftCard"
        Dim handler As IGiftCardBalanceInquiry
        Dim parameters As Dictionary(Of String, Object)

        Try
            parameters = New Dictionary(Of String, Object) From {
                                                             {"Controller", TheModCntr},
                                                             {"Transaction", taobj},
                                                             {"Barcode", szBarcode},
                                                             {"Value", dBalance},
                                                             {"Receipt", szReceipt},
                                                             {"GiftCardBalanceInternalInquiry", True}
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

#Region "IGiftCardCancellationPayment"

    Public Function GiftCardCancellation(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec) As IGiftCardReturnCode
        GiftCardCancellation = IGiftCardReturnCode.OK
        Dim funcName As String = "GiftCardCancellation"
        Dim handler As IGiftCardCancellationPayment
        Dim parameters As Dictionary(Of String, Object)

        Try
            parameters = New Dictionary(Of String, Object) From {
                                                              {"Controller", TheModCntr},
                                                              {"Transaction", taobj},
                                                              {"MediaRecord", MyTaMediaRec}
                                                          }

            handler = createPosModelObject(Of IGiftCardCancellationPayment)(TheModCntr, "GiftCardController", 0, False)
            If handler Is Nothing Then
                ' gift card handler is not defined into the database
                GiftCardCancellation = IGiftCardReturnCode.KO
                Exit Function
            End If

            GiftCardCancellation = handler.GiftCardCancellation(parameters)

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

#Region "IBuoniPasto"

    ''' <summary>
    '''     Implementazione  del  Metodo  dell'interfaccia  che espone
    '''     la funzione princiapale  atta ad eseguire la presentazione
    '''     del form per l'immissione dei buoni pasto, e negli handler
    '''     alla gestione dell'inserimento dei vari EAN di BP da validare
    '''     ad Argentea.
    ''' </summary>
    ''' <param name="TheModCntr">Il Controller di riporto dalla Transazione di base</param>
    ''' <param name="taobj">La Transazione di base per completo</param>
    ''' <param name="MyTaMediaRec">Il Media iniziale di riporto da aggiornare una volta eseguito il procecco corrente</param>
    ''' <param name="MyTaMediaMemberDetailRec">Il detaglio del Mediarec di riporto</param>
    ''' <returns></returns>
    Public Function ProcessBPCartaceo(ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef taobj As TPDotnet.Pos.TA, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec, ByRef MyTaMediaMemberDetailRec As TaMediaMemberDetailRec) As IBPReturnCode

        ProcessBPCartaceo = IBPReturnCode.OK
        Dim funcName As String = "ProcessBPCartaceo"
        Dim handler As IBPCDematerialize

        Try

            ' Ricreiamo il Form e l'Handler per la Gestione solo istanza
            ' l'handler per l'oggetto PosModel da passare al Form di istanza.

            handler = createPosModelObject(Of IBPCDematerialize)(TheModCntr, "BPCController", 0, False)
            If handler Is Nothing Then
                ' gift card handler is not defined into the database
                ProcessBPCartaceo = IBPReturnCode.KO
                Exit Function
            End If

            ' Richiama il metodo dell'interfaccia --> "BPCController.Dematerialize"
            ' che avvia il form e rimane sulla gestione tramite eventi sul form per
            ' la scanzione e validazione dei BP

            ProcessBPCartaceo = handler.Dematerialize(New Dictionary(Of String, Object) From {
                                                              {"Controller", TheModCntr},
                                                              {"Transaction", taobj},
                                                              {"MediaRecord", MyTaMediaRec},
                                                              {"MediaMemberDetailRecord", MyTaMediaMemberDetailRec}
                                                          })
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


    Public Function ProcessBCElettronico()



    End Function

#End Region

End Class
