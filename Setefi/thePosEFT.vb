Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.Win32
Imports TPDotnet.Pos
Imports TPDotnet.Services.Rounding
Imports TPDotnet.IT.Common.Pos

Public Class thePosEFT
    Implements TPDotnet.Pos.IEFT

#Region "Documentation"
    ' ********** ********** ********** **********
    ' E F T - Setefi protocol implementation
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region
#Region "Global member variables"

    Protected m_theModCntr As ModCntr
    Protected m_taobj As TA
    Protected m_colObjects As Collections.Hashtable

    Protected WithEvents TheSetefiEft As Setefi = New Setefi
    Protected MyOperationNumber As Short = 1 ' numOperazione start from 1. It is saved into registry

    ' special handling parameters
    Protected IsGift As Boolean = False
    Protected GiftPan As String = "00000000000"
    Protected IsRCard As Boolean = False
    Protected RCardPan As String = "0000"
#End Region

#Region "Parameters"
    Protected PRINT_KO_RECEIPT_TOO As Boolean = False
    Protected NUMBER_OF_COPY_FOR_OK_RECEIPT As Integer = 1
    Protected ASK_FOR_MANUAL_PAYMENT As Boolean = False
    Protected SHOW_MSG_TRANSACTION_EXECUTED As Boolean = False
    'Protected PAN_GIFT_CARD_LISTA_NOZZE As String = "00000000000"
    Protected ClosureTotalValueFirstLineTokens As New System.Collections.Generic.List(Of String)
    Protected ClosureTotalValueSecondLineTokens As New System.Collections.Generic.List(Of String)
#End Region

#Region "Form handling"
    Dim WithEvents myForm As FormEFTOperation = Nothing
    'Dim myForm As FormEFTHandling = Nothing

    ' open the eft form and disable the others function
    Protected Overridable Sub OpenEftForm()
        myForm = Me.TheModCntr.GetCustomizedForm(GetType(FormEFTOperation), STRETCH_TO_SMALL_WINDOW)
        myForm.theModCntr = Me.TheModCntr
        TheModCntr.DialogFormName = myForm.Text
        TheModCntr.SetFuncKeys(False)
        TheModCntr.DialogActiv = True
        myForm.Waiter1.Visible = True
        myForm.Waiter1.Start()
        myForm.Show()
    End Sub

    Protected Overridable Sub CloseEftForm()
        TheModCntr.DialogActiv = False
        TheModCntr.DialogFormName = ""
        TheModCntr.SetFuncKeys(True)
        myForm.Waiter1.Visible = False
        myForm.Waiter1.Stop()
        myForm.Close()
        TheModCntr.EndForm()
    End Sub
#End Region

#Region "Interface methods"
    ' eft payment
    ' return with 0 = OK or <> 0 = BAD
    Public Overridable Function PostEFTDevice(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec) As Short Implements TPDotnet.Pos.IEFT.PostEFTDevice
        Return EFTHandler(taobj, MyTaMediaRec, False, True, True)
    End Function

    ' void the last eft transaction
    ' return with 0 = OK or <> 0 = BAD
    Public Overridable Function VoidEFTDevice(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec) As Short Implements TPDotnet.Pos.IEFT.VoidEftDevice
        Return EFTHandler(taobj, MyTaMediaRec, True, False, True)
    End Function

#End Region

#Region "Internal functions"

    ' ask if the transaction was closed directly on the eft
    Protected Overridable Function AskForManualPayment() As Short
        Dim iRet As Short
        iRet = TPMsgBoxRet(PosDef.TARMessageTypes.TPQUESTION, getPosTxtNew((TheModCntr.contxt), "UserMessage", TXT_EFT_MANUAL_PAYMENT), MsgBoxStyle.YesNo, TheModCntr, 777, "Eseguito pagamento manuale?")
        If iRet = MsgBoxResult.No Then
            Return 110
        Else
            Return 0
        End If
    End Function

#Region "Swap Media"

    Protected Overridable Sub SwapElectronicMedia(ByRef taobj As TA, ByRef TheMediaRec As TaMediaRec, ByRef CardType As String)

        Dim MySelectMediaClass As clsSelectMedia
        'Dim MyTaMediaRec As TaMediaRec
        'Dim i As Integer
        Dim m_lMediaMember As Integer = 0 ' just for internal test all media will be AMEX
        Dim bRet As Boolean
        Dim szFileName As String = ""

        Try

            LOG_FuncStart(getLocationString("SwapElectronicMedia"))

            'get an instance of the class with reads the DB-table for the Media
            MySelectMediaClass = createPosModelObject(Of clsSelectMedia)(TheModCntr, "clsSelectMedia", 0, True)
            If MySelectMediaClass Is Nothing Then
                ' no media Record for this extension in database present
                LOG_Error(getLocationString("SwapElectronicMedia"), "configuration error: module not found: clsSelectMedia")
                Exit Sub
            End If

            'read the swap table
            'format of each swap file line is Type=MediaMember (eg. 00=400)
            'Type is defined by Setefi
            'MediaMember is defined in TP.NET
            szFileName = getPosConfigurationPath() + "\" + "EftSwap.txt"
            If File.Exists(szFileName) Then
                Dim lines() As String = File.ReadAllLines(szFileName)
                For Each line As String In lines
                    Dim Type As String = "", MediaMember As String = ""
                    Type = line.Split("=")(0)
                    MediaMember = line.Split("=")(1)
                    If Not Type Is Nothing AndAlso Type <> "" AndAlso Not MediaMember Is Nothing AndAlso MediaMember <> "" Then
                        If Type = CardType Then
                            m_lMediaMember = Integer.Parse(MediaMember)
                            Exit For
                        End If
                    End If
                Next
            End If

            If m_lMediaMember > 0 AndAlso m_lMediaMember <> TheMediaRec.PAYMENTinMedia.lMediaMember Then
                If Not TheMediaRec Is Nothing Then
                    bRet = MySelectMediaClass.FillPaymentDataFromID(TheModCntr, TheMediaRec.PAYMENTinMedia, _
                                                             m_lMediaMember, taobj, taobj.colObjects)
                    If bRet = False Then
                        ' no media Record for this extension in database present
                        LOG_Error(getLocationString("SwapElectronicMedia"), "MediaMember " & m_lMediaMember.ToString & " not found in database")
                        MySelectMediaClass = Nothing
                        'TheMediaRec = Nothing
                        Exit Sub
                    Else
                        AddDetailToMedia(taobj, TheMediaRec)
                    End If
                End If
            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("SwapElectronicMedia"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("SwapElectronicMedia"), InnerEx)
            End Try
        Finally
            MySelectMediaClass = Nothing
            LOG_FuncExit(getLocationString("SwapElectronicMedia"), "Function SwapElectronicMedia")
        End Try
    End Sub

    Protected Overridable Sub AddDetailToMedia(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec)

        Dim MyTaMediaMemberDetailRec As TaMediaMemberDetailRec = Nothing

        Try
            LOG_FuncStart(getLocationString("AddDetailToMedia"))

            If MyTaMediaRec.PAYMENTinMedia.bIsDocRM <> False Then
                If MyTaMediaRec.dReturn = 0.0# Then
                    'it is a payment
                    MyTaMediaMemberDetailRec = taobj.CreateTaObject(PosDef.TARecTypes.iTA_MEDIA_MEMBER_DETAIL)
                    MyTaMediaMemberDetailRec.lAppearanceWorkstationNmbr = taobj.lWorkStationNmbr
                    MyTaMediaMemberDetailRec.lAppearanceTaNmbr = taobj.lactTaNmbr
                    MyTaMediaMemberDetailRec.lAppearanceTaCreateNmbr = taobj.lTaCreateNmbr
                    MyTaMediaMemberDetailRec.lAppearanceSeqNmbr = 1
                    MyTaMediaMemberDetailRec.szAppearanceDate = Format(Now, "yyyyMMddHHmmss")
                    MyTaMediaMemberDetailRec.lRetailStoreID = taobj.lRetailStoreID
                    MyTaMediaMemberDetailRec.lOperatorID = taobj.lActOperatorID
                    MyTaMediaMemberDetailRec.lMediaNmbr = MyTaMediaRec.PAYMENTinMedia.lMediaNmbr
                    MyTaMediaMemberDetailRec.lMediaMember = MyTaMediaRec.PAYMENTinMedia.lMediaMember
                    MyTaMediaMemberDetailRec.szMediaDesc = MyTaMediaRec.PAYMENTinMedia.szDesc
                    MyTaMediaMemberDetailRec.dQty = MyTaMediaRec.dTaQty 'Quantity
                    MyTaMediaMemberDetailRec.dAmount = MyTaMediaRec.dTaPaid 'Single amount in local curr
                    MyTaMediaMemberDetailRec.dAmountForeign = MyTaMediaRec.dPaidForeignCurr 'Single amount in foreign curr
                    MyTaMediaMemberDetailRec.dTotalAmount = MyTaMediaRec.dTaPaidTotal 'Total amount (dQty * dTaPaid) in local curr
                    MyTaMediaMemberDetailRec.dTotalAmountForeign = Rounding.dRounding(MyTaMediaRec.dPaidForeignCurr * MyTaMediaRec.dTaQty, _
                                                                                TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, _
                                                                                MyTaMediaRec.PAYMENTinMedia.lPaySDOC, _
                                                                                MyTaMediaRec.PAYMENTinMedia.lPayDecNmbr) 'Total amount (dPaidForeignCurr * dTaQty) in foreign curr
                    MyTaMediaMemberDetailRec.szPostingType = "POSIN" 'this type will not get serialized!
                    MyTaMediaMemberDetailRec.szCountingType = MyTaMediaRec.PAYMENTinMedia.szPOSCountingType
                    MyTaMediaMemberDetailRec.szDeclarationStatus = ""
                    MyTaMediaMemberDetailRec.lPrintOnReceipt = 0
                    MyTaMediaMemberDetailRec.szStatusText = ""
                    MyTaMediaMemberDetailRec.szBarCode = MyTaMediaRec.szBarcode
                    MyTaMediaMemberDetailRec.szSerialNmbr = MyTaMediaRec.szSerialNmbr

                    MyTaMediaMemberDetailRec.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr
                    taobj.Add(MyTaMediaMemberDetailRec)
                End If
            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("AddDetailToMedia"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("AddDetailToMedia"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("AddDetailToMedia"), "Function AddDetailToMedia")
        End Try

        Exit Sub
    End Sub
#End Region

    ' Load user defined parameters if any.
    Protected Overridable Sub GetParameters()

        Dim TMP_PRINT_KO_RECEIPT_TOO As String = TheModCntr.getParam("StPosMod" + "." + TypeName(Me) + "." + "PRINT_KO_RECEIPT_TOO")
        If TMP_PRINT_KO_RECEIPT_TOO <> "" AndAlso TMP_PRINT_KO_RECEIPT_TOO <> " " AndAlso TMP_PRINT_KO_RECEIPT_TOO <> "N" Then
            PRINT_KO_RECEIPT_TOO = True
        End If

        Dim TMP_NUMBER_OF_COPY_FOR_OK_RECEIPT As String = TheModCntr.getParam("StPosMod" + "." + TypeName(Me) + "." + "NUMBER_OF_COPY_FOR_OK_RECEIPT")
        If TMP_NUMBER_OF_COPY_FOR_OK_RECEIPT <> "" AndAlso TMP_NUMBER_OF_COPY_FOR_OK_RECEIPT <> " " Then
            If IsNumeric(TMP_NUMBER_OF_COPY_FOR_OK_RECEIPT) Then
                NUMBER_OF_COPY_FOR_OK_RECEIPT = Integer.Parse(TMP_NUMBER_OF_COPY_FOR_OK_RECEIPT)
            End If
        End If

        Dim TMP_ASK_FOR_MANUAL_PAYMENT As String = TheModCntr.getParam("StPosMod" + "." + TypeName(Me) + "." + "ASK_FOR_MANUAL_PAYMENT")
        If TMP_ASK_FOR_MANUAL_PAYMENT <> "" AndAlso TMP_ASK_FOR_MANUAL_PAYMENT <> " " AndAlso TMP_ASK_FOR_MANUAL_PAYMENT <> "N" Then
            ASK_FOR_MANUAL_PAYMENT = True
        End If

        Dim TMP_SHOW_MSG_TRANSACTION_EXECUTED As String = TheModCntr.getParam("StPosMod" + "." + TypeName(Me) + "." + "SHOW_MSG_TRANSACTION_EXECUTED")
        If TMP_SHOW_MSG_TRANSACTION_EXECUTED <> "" AndAlso TMP_SHOW_MSG_TRANSACTION_EXECUTED <> " " AndAlso TMP_SHOW_MSG_TRANSACTION_EXECUTED <> "N" Then
            SHOW_MSG_TRANSACTION_EXECUTED = True
        End If

        ClosureTotalValueFirstLineTokens.Clear()
        Dim TMP_CLOSURE_LINE_TOKENS As String = theModCntr.getParam("StPosMod" + "." + TypeName(Me) + "." + "CLOSURE_FIRST_LINE_TOKENS")
        If TMP_CLOSURE_LINE_TOKENS <> "" AndAlso TMP_CLOSURE_LINE_TOKENS <> " " Then
            Dim split As String() = TMP_CLOSURE_LINE_TOKENS.Split(";")
            For Each s As String In split
                ClosureTotalValueFirstLineTokens.Add(s.ToUpper)
            Next
        End If

        ClosureTotalValueSecondLineTokens.Clear()
        TMP_CLOSURE_LINE_TOKENS = theModCntr.getParam("StPosMod" + "." + TypeName(Me) + "." + "CLOSURE_SECOND_LINE_TOKENS")
        If TMP_CLOSURE_LINE_TOKENS <> "" AndAlso TMP_CLOSURE_LINE_TOKENS <> " " Then
            Dim split As String() = TMP_CLOSURE_LINE_TOKENS.Split(";")
            For Each s As String In split
                ClosureTotalValueSecondLineTokens.Add(s.ToUpper)
            Next
            ClosureTotalValueSecondLineTokens.Reverse()
        End If

    End Sub

    Public Overridable Sub SetSpecialHandlingParameters(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec, ByVal IsVoid As Boolean, Optional ByVal AddEftInfoToCurrentTa As Boolean = True, Optional ByVal CreateAndPrintAnEftTa As Boolean = True)

        ' no special handling as default
        IsGift = False
        GiftPan = "00000000000"
        IsRCard = False
        RCardPan = "0000"

    End Sub

    ' the eft transaction handler
    Protected TheAmount As Double = 0.0#
    Public Overridable Function EFTHandler(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec, ByVal IsVoid As Boolean, Optional ByVal AddEftInfoToCurrentTa As Boolean = True, Optional ByVal CreateAndPrintAnEftTa As Boolean = True) As Short

        Dim szRequestType As String = ""

        EFTHandler = 110 ' set a bad return code as default
        LOG_FuncStart(getLocationString("EFTHandler"))

        TheAmount = 0.0#

        Try
            OpenEftForm()

            GetParameters()
            SetSpecialHandlingParameters(taobj, MyTaMediaRec, IsVoid, AddEftInfoToCurrentTa, CreateAndPrintAnEftTa)

            ' init variables depending on the kind of operation (payment/void)
            If Not IsVoid Then
                TheAmount = MyTaMediaRec.dTaPaidTotal
            Else
                TheAmount = GetLastPaymentTransactionAmount() / (10 ^ m_theModCntr.iEXACTNESS_IN_DIGITS)
            End If

            LOG_Error(getLocationString("EFTHandler"), "Start payment thread")

            ' call the worker
            If TheSetefiEft.PaySetefi(TheAmount, Me.TheModCntr, IsGift, GiftPan, IsVoid, IsRCard, RCardPan) = 0 Then
                Do While TheSetefiEft.DialogActiv = True
                    Sleep(100)
                    System.Windows.Forms.Application.DoEvents()
                Loop
            End If


            LOG_Error(getLocationString("EFTHandler"), "End payment thread (ESITO = <" & TheSetefiEft.DatiAutorizzazioneECarta.ESITO_TRANSAZIONE & _
                                                        "> STATO = <" & TheSetefiEft.Stato.ToString & ")")

            If String.Equals(TheSetefiEft.DatiAutorizzazioneECarta.ESITO_TRANSAZIONE, "00") Then
                ' Esito reports OK
                Dim TxSaved As Boolean = EFTResultHandler(taobj, MyTaMediaRec, TheSetefiEft.GetTicket(), AddEftInfoToCurrentTa, CreateAndPrintAnEftTa)
                If Not IsVoid Then
                    SwapElectronicMedia(taobj, MyTaMediaRec, TheSetefiEft.DatiAutorizzazioneECarta.CODICE_SOCIETA_EMETTITRICE)
                    SetLastPaymentTransactionAmount(MyTaMediaRec.dTaPaidTotal * (10 ^ m_theModCntr.iEXACTNESS_IN_DIGITS))
                    SetLastPaymentTransactionNumber(taobj.lactTaNmbr - 1)
                End If
                If TxSaved Then
                    EFTHandler = 0
                End If
            Else
                ' Esito reports BAD
                TPMsgBox(PosDef.TARMessageTypes.TPERROR, TheSetefiEft.ErrorMessage, 0, theModCntr, TheSetefiEft.ErrorMessage)
                LOG_Error(getLocationString("EFTHandler"), TheSetefiEft.ErrorMessage)
            End If

            'If String.Equals(TheSetefiEft.DatiAutorizzazioneECarta.ESITO_TRANSAZIONE, "  ") Then
            If String.IsNullOrEmpty(TheSetefiEft.DatiAutorizzazioneECarta.ESITO_TRANSAZIONE) _
            OrElse String.IsNullOrEmpty(TheSetefiEft.DatiAutorizzazioneECarta.ESITO_TRANSAZIONE.Trim) _
            Then
                ' we don't know what is really happened
                If ASK_FOR_MANUAL_PAYMENT AndAlso Not IsGift Then
                    EFTHandler = AskForManualPayment()
                End If
            End If

            ' abort the thread if it is alive
            TheSetefiEft.AbortPayment()

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("EFTHandler"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("EFTHandler"), InnerEx)
            End Try
        Finally
            CloseEftForm()
            LOG_FuncExit(getLocationString("EFTHandler"), "Function EFTHandler return " & EFTHandler.ToString)
        End Try

    End Function

    Public Overridable Function EFTClosureHandler(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec, ByVal IsVoid As Boolean, Optional ByVal AddEftInfoToCurrentTa As Boolean = True, Optional ByVal CreateAndPrintAnEftTa As Boolean = True) As Short

        Dim szRequestType As String = ""
        Dim szTicket As String = ""

        EFTClosureHandler = 110 ' set a bad return code as default
        LOG_FuncStart(getLocationString("EFTClosureHandler"))

        Try
            'OpenEftForm()

            GetParameters()
            SetSpecialHandlingParameters(taobj, MyTaMediaRec, IsVoid, AddEftInfoToCurrentTa, CreateAndPrintAnEftTa)

            LOG_Error(getLocationString("EFTClosureHandler"), "Start payment thread")

            ' call the worker
            If TheSetefiEft.CloseSetefi(Me.theModCntr) = 0 Then
                Do While TheSetefiEft.DialogActiv = True
                    Sleep(100)
                    System.Windows.Forms.Application.DoEvents()
                Loop
            End If

            szTicket = TheSetefiEft.GetTicket()
            LOG_Error(getLocationString("EFTClosureHandler"), "End closure thread (TextLen = <" & szTicket.Length & _
                                                        "> STATO = <" & TheSetefiEft.Stato.ToString & ")")

            If TheSetefiEft.Stato = Setefi.States.SETEFI_SUCCESS AndAlso Not String.IsNullOrEmpty(szTicket) Then
                ' Esito reports OK
                Dim TxSaved As Boolean = EFTResultHandler(taobj, MyTaMediaRec, TheSetefiEft.GetTicket(), AddEftInfoToCurrentTa, CreateAndPrintAnEftTa, True)
                EFTClosureHandler = 0
            Else
                ' Esito reports BAD
                TPMsgBox(PosDef.TARMessageTypes.TPERROR, TheSetefiEft.ErrorMessage, 0, theModCntr, TheSetefiEft.ErrorMessage)
                LOG_Error(getLocationString("EFTClosureHandler"), TheSetefiEft.ErrorMessage)
            End If

            ' abort the thread if it is alive
            TheSetefiEft.AbortPayment()

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("EFTClosureHandler"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("EFTClosureHandler"), InnerEx)
            End Try
        Finally
            'CloseEftForm()
            LOG_FuncExit(getLocationString("EFTHandler"), "Function EFTClosureHandler return " & EFTClosureHandler.ToString)
        End Try

    End Function

    Protected Class ClosureTotalLineIdentifiers
#Region "Variables"
        Protected m_FirstLine As String = ""
        Protected m_SecondLine As String = ""
#End Region

#Region "Properties"
        Public Property FirstLine() As String
            Get
                Return m_FirstLine
            End Get
            Set(ByVal value As String)
                m_FirstLine = value
            End Set
        End Property

        Public Property SecondLine() As String
            Get
                Return m_SecondLine
            End Get
            Set(ByVal value As String)
                m_SecondLine = value
            End Set
        End Property
#End Region
#Region "Constructor"
        Public Sub New(ByVal first As String, ByVal second As String)
            m_FirstLine = first
            m_SecondLine = second
        End Sub
#End Region
    End Class

    ' add the EFT_INFO record to the current transaction.
    ' Call also the EriteEFTTransactionToServer.
    Public Function EFTResultHandler(ByRef taobj As TA, ByRef MyTaMediaRec As TaMediaRec, ByRef ReceiptText As String, Optional ByVal AddEftInfoToCurrentTa As Boolean = True, Optional ByVal CreateAndPrintAnEftTa As Boolean = True, Optional ByVal IsClosure As Boolean = False) As Boolean

        Dim m_ReceiptText As String = ReceiptText
        Dim bLineCouldBeTheTotal As Boolean = False
        Dim bTotalLineFound As Boolean = False

        EFTResultHandler = False
        LOG_FuncStart(getLocationString("EFTResultHandler"))

        Try
            ' we use the following linees to print the eft receipt on the pos receipt
            Dim MyEftInfo As TaEFTInfo = Nothing
            MyEftInfo = taobj.CreateTaObject(PosDef.TARecTypes.iTA_EFTINFO)
            MyEftInfo.lShowInfoLineOnPosReceipt = 0
            MyEftInfo.lShowReceiptLineOnPosReceipt = 1

            Dim tc(3) As String
            tc(0) = vbCrLf
            tc(1) = vbCr
            tc(2) = vbLf

            ' this part is valid only for payment/void transactions
            If Not m_ReceiptText.Length > 0 AndAlso Not IsClosure Then
                Dim msg As String = getPosTxtNew((theModCntr.contxt), "UserMessage", WNG_EFT_PAYMENT_OK_BUT_NOT_RECEIPT)
                TPMsgBox(PosDef.TARMessageTypes.TPWARNING, msg, 0, theModCntr, "Message")
                LOG_Error(getLocationString("EFTResultHandler"), "The transaction has been executed but the text is not present.")
                m_ReceiptText += "DATA " & Format(Now, "dd/MM/yyyy hh:mm") & vbCrLf
                m_ReceiptText += "Nome Acquirer: " & TheSetefiEft.DatiAutorizzazioneECarta.NOME_ACQUIRER & vbCrLf
                m_ReceiptText += "Cod. Acquirer: " & TheSetefiEft.DatiAutorizzazioneECarta.CODICE_ACQUIRER_PAGOBANCOMAT & vbCrLf
                m_ReceiptText += "PAN: " & TheSetefiEft.DatiAutorizzazioneECarta.PAN & vbCrLf
                m_ReceiptText += "Scadenza: " & TheSetefiEft.DatiAutorizzazioneECarta.SCADENZA_CARTA_MMAA & vbCrLf
                m_ReceiptText += "N.Op: " & TheSetefiEft.DatiAutorizzazioneECarta.NUMERO_PROGRESSIVO_OPERAZIONE & vbCrLf
                m_ReceiptText += "Cod. Aut.: " & TheSetefiEft.DatiAutorizzazioneECarta.CODICE_AUTORIZZAZIONE & vbCrLf
                m_ReceiptText += "EURO " & Format(TheAmount, m_theModCntr.getFormatString4Price()) & vbCrLf
                If TheSetefiEft.DatiAutorizzazioneECarta.TIPO_CARTA = "4" Then
                    m_ReceiptText += "*                      *" & vbCrLf & _
                                     " ______________________ " & vbCrLf & _
                                     " ( Firma - Signature )"
                End If
            End If

            ' the presentation handling does not allow a line completely composed by ".", so lets replace it with "_"
            m_ReceiptText = m_ReceiptText.Replace("........................", "________________________")

            Dim PrintArray() As String = m_ReceiptText.Split(tc, StringSplitOptions.None)
            For i As Integer = 0 To PrintArray.Length - 1
                'Dim sztmp As String = PrintArray(i)

                'sztmp = sztmp.Trim(tc(0))
                'sztmp = sztmp.Trim(tc(1))
                'sztmp = sztmp.Trim(tc(2))
                MyEftInfo.AddField("EFTReceiptLine" & Format(i + 1, "00"), DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            Next

            If IsClosure Then
                MyEftInfo.AddField("dGeneralClosureAmount", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            End If

            For i As Integer = 0 To PrintArray.Length - 1
                Dim sztmp As String = PrintArray(i)

                sztmp = sztmp.Trim(tc(0))
                sztmp = sztmp.Trim(tc(1))
                sztmp = sztmp.Trim(tc(2))
                MyEftInfo.setPropertybyName("EFTReceiptLine" & Format(i + 1, "00"), sztmp)

                If IsClosure AndAlso Not bTotalLineFound Then
                    If bLineCouldBeTheTotal Then
                        ' EUR                 1.23
                        For Each s As String In ClosureTotalValueSecondLineTokens
                            If sztmp.Trim.ToUpper.StartsWith(s) Then
                                Dim sVal As String = sztmp.Trim.ToUpper.Replace(s, "")
                                Dim dVal As Double = 0
                                If Double.TryParse(sVal, dVal) Then
                                    MyEftInfo.setPropertybyName("dGeneralClosureAmount", dVal)
                                    bTotalLineFound = True
                                    Exit For
                                End If
                            End If
                        Next
                        bLineCouldBeTheTotal = False

                    ElseIf ClosureTotalValueFirstLineTokens.Contains(sztmp.Trim.ToUpper) Then
                        bLineCouldBeTheTotal = True
                    End If
                End If

            Next

            ' set reference if eft info came from a media record
            If Not MyTaMediaRec Is Nothing Then
                MyEftInfo.theHdr.lTaRefToCreateNmbr = MyTaMediaRec.theHdr.lTaCreateNmbr
            End If

            If CreateAndPrintAnEftTa Then
                If Not WriteEFTTransactionToServer(MyEftInfo) Then
                    LOG_Error(getLocationString("EFTResultHandler"), "Could not write eft transaction to server.")
                    Exit Function
                End If
            End If

            ' re-allign the ta number, may be changed by WriteEFTTransactionToServer
            taobj.AssignRegValues()

            Dim MyTaHdrRec As TaHdrRec = taobj.GetTALine(1)
            MyTaHdrRec.lTaNmbr = taobj.lactTaNmbr
            MyTaHdrRec.lReceiptNmbr = taobj.lactTaNmbr

            If AddEftInfoToCurrentTa Then
                ' Ema 20130325 : reset the lTaCreateNmbr in order to be assigned properly by the following statement
                MyEftInfo.theHdr.lTaCreateNmbr = 0
                MyEftInfo.theHdr.lTaSeqNmbr = 0
                taobj.Add(MyEftInfo)
            End If

            EFTResultHandler = True

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("EFTResultHandler"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("EFTResultHandler"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("EFTResultHandler"), "Function EFTResultHandler return " & EFTResultHandler.ToString)
        End Try

    End Function

    ' create a transaction with the eft transaction receipt, print it, and write to the store server
    Public Function WriteEFTTransactionToServer(ByRef MyTaEFTInfoRec As TPDotnet.Pos.TaEFTInfo) As Boolean

        Dim MyNewTa As TA
        Dim MyNewTaHashTable As New System.Collections.Hashtable()
        Dim MyTaFtrRec As TaFtrRec

        WriteEFTTransactionToServer = False

        Try
            MyNewTa = createTheTA(TheModCntr, MyNewTaHashTable)
            Dim szMode As String = GetOnlyMode(TheModCntr)
            'initialize the new transaction with the values used to show on display
            '====================================================================== 
            If MyNewTa.iTaStatus = 0 Then
                'Initialize the new Transaction
                MyNewTa.szSignOnName = m_theModCntr.szSignOnName
                MyNewTa.szPrintCode = m_theModCntr.szPrintCode
                MyNewTa.szActEmployeeName = m_theModCntr.szActEmployeeName
                MyNewTa.lActOperatorID = m_theModCntr.lActOperatorID
                MyNewTa.lRetailStoreID = m_theModCntr.lRetailStoreID
                MyNewTa.szWorkstationGroupID = m_theModCntr.szWorkstationGroupID
                MyNewTa.szWorkstationID = m_theModCntr.szWorkstationID
                MyNewTa.lWorkStationNmbr = m_theModCntr.lWorkstationNmbr
                MyNewTa.szTaxCalculationAtTotal = TheModCntr.getParam(PARAMETER_MOD_CNTR & ".ModCntr.TAX_CALC_AT_TOTAL")
                MyNewTa.bTrainingMode = TheModCntr.bTrainingMode
                MyNewTa.szMode = szMode
                MyNewTa.iExactness = m_theModCntr.iEXACTNESS_IN_DIGITS

                MyNewTa.TAStart(theModCntr.contxt, theModCntr.con)

                ' here we reload the parameters from table parameter
                'TheModCntr.ReadParameter()

                ' fill the headerlines from globstoreval
                fillHeaderLines(TheModCntr, MyNewTa)

                MyNewTa.TARefresh()
            End If

            ' todo : set bPrintReceipt to false and handle directly the ticket prints to be sure that the ticket is printed out
            MyNewTa.bPrintReceipt = True
            MyNewTa.bTAtoFile = True
            MyNewTa.bDelete = True ' ok , we will delete this TA

            ' add out eft record
            MyNewTa.Add(MyTaEFTInfoRec)

            'MyNewTa.TAEnd(fillFooterLines((TheModCntr.con), MyNewTa, TheModCntr))
            MyTaFtrRec = fillFooterLines((theModCntr.con), MyNewTa, theModCntr)
            MyTaFtrRec.bPrintBarcode = 0
            MyTaFtrRec.szTaType = TA_TYPE_TMP_NO_BARCODE
            MyNewTa.TAEnd(MyTaFtrRec)

            Dim myclsEndTaHandling As clsEndTAHandling
            If MyNewTa.bPrintReceipt OrElse MyNewTa.bTAtoFile Then
                Dim bRet As Boolean = False
                myclsEndTaHandling = createPosObject(Of clsEndTAHandling)(TheModCntr, "clsEndTAHandling", 0)
                If Not myclsEndTaHandling Is Nothing Then
                    bRet = myclsEndTaHandling.EndTA(MyNewTa, TheModCntr)
                    myclsEndTaHandling = Nothing
                End If

                ' abort in case of fiscal printer errors only
                If bRet = False Then
                    ' Exit Function
                Else
                    WriteEFTTransactionToServer = True
                End If
            End If
        Catch ex As Exception
            Try
                LOG_Error(getLocationString("WriteEFTTransactionToServer"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("WriteEFTTransactionToServer"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("WriteEFTTransactionToServer"), "Function Finished")
        End Try

    End Function

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

#Region "Registry EFT Functions"

    Protected Sub SetLastPaymentTransactionAmount(ByVal LastPaymentTransactionAmount As Integer)
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

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function

    Protected Sub SetLastPaymentTransactionNumber(ByVal LastPaymentTransactionNumber As Integer)
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            REG.SetValue("LastPaymentTransactionNumber", LastPaymentTransactionNumber.ToString)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Sub

    Public Shared Function GetLastPaymentTransactionNumber() As Integer
        Dim REG As RegistryKey = Nothing
        Try
            REG = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wincor Nixdorf\TPDotnet\pos\EFT", True)
            GetLastPaymentTransactionNumber = REG.GetValue("LastPaymentTransactionNumber", 0)
        Catch ex As Exception

        Finally
            If Not REG Is Nothing Then
                REG.Close()
            End If
        End Try
    End Function

#End Region

#Region "Italy"

    Private Delegate Sub InfoUpdateEventHandler(ByVal message As String)
    Public Sub WriteInfo(ByVal value As String)

        Try
            If myForm Is Nothing Then
                Exit Sub
            End If

            If (myForm.ListboxLog.InvokeRequired) Then
                myForm.ListboxLog.Invoke(New InfoUpdateEventHandler(AddressOf WriteInfo), value)
            Else
                myForm.ListboxLog.Items.Add(value)
                myForm.ListboxLog.ScrollDown()
                myForm.ListboxLog.Refresh()
            End If

            'If (myForm.lbl_Info.InvokeRequired) Then
            '    myForm.lbl_Info.Invoke(New InfoUpdateEventHandler(AddressOf WriteInfo), value)
            'Else
            '    myForm.lbl_Info.Text += value
            'End If
        Catch ex As ObjectDisposedException
            Console.WriteLine(ex.Message)
        End Try

    End Sub


    Public Sub DaSistemiStatusChanged(ByVal state As Integer, ByVal message As String) Handles TheSetefiEft.DaSistemiStatusChanged

        If state = Setefi.States.SETEFI_ERROR Then
            WriteInfo(String.Format("{0}", message))

        ElseIf state = Setefi.States.SETEFI_IN_PROGRESS Then
            WriteInfo(String.Format("{0}", message))

        ElseIf state = Setefi.States.SETEFI_SUCCESS Then
            WriteInfo(String.Format("{0}", message))
        End If

    End Sub

    Private Delegate Sub RemainingSecondsUpdateEventHandler(ByVal seconds As Integer)
    Public Sub WriteRemainingSeconds(ByVal s As Integer)
        Try
            If myForm Is Nothing Then
                Exit Sub
            End If

            If myForm.timeout_lbl.InvokeRequired Then
                myForm.timeout_lbl.Invoke(New RemainingSecondsUpdateEventHandler(AddressOf WriteRemainingSeconds), s)
            Else
                myForm.timeout_lbl.Text = String.Format("Tempo rimanente per l'operazione in corso : {0}", s)
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Public Sub DaSistemiRemainindSecondsChanged(ByVal s As Integer) Handles TheSetefiEft.DaSistemiRemainingSecond
        WriteRemainingSeconds(s)
    End Sub

    Public Sub CancelOperationIsAvailable(ByVal available As Boolean)
        Try
            If myForm Is Nothing Then
                Exit Sub
            End If

            If myForm.cmdCancel.InvokeRequired Then
                myForm.cmdCancel.Invoke(New DaSistemiCancelOperationIsAvailableChangedEventHandler(AddressOf DaSistemiCancelOperationIsAvailableChanged), available)
            Else
                myForm.cmdCancel.Visible = available
                myForm.cmdCancel.Enabled = available
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Delegate Sub DaSistemiCancelOperationIsAvailableChangedEventHandler(ByVal available As Boolean)
    Public Sub DaSistemiCancelOperationIsAvailableChanged(ByVal available As Boolean) Handles TheSetefiEft.DaSistemiCancelOperationIsAvailableChanged
        CancelOperationIsAvailable(available)
    End Sub

    Public Sub EFTDeviceMessage(ByRef szMessage As String) Handles myForm.EFTDeviceMessage

        If szMessage = "Abort" Then
            Dim msg As String = getPosTxtNew((theModCntr.contxt), "UserMessage", TXT_EFT_PAYMENT_ABORT)
            TheSetefiEft.AbortPayment(msg)
        End If

    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property theModCntr() As TPDotnet.Pos.IModCntr Implements TPDotnet.Pos.IEFT.theModCntr
        Get
            Return m_theModCntr
        End Get
    End Property

    Public ReadOnly Property taobj() As TPDotnet.Pos.TA Implements TPDotnet.Pos.IEFT.taobj
        Get
            Return m_taobj
        End Get
    End Property

#End Region

    Public Sub New()

    End Sub

    Public ReadOnly Property CardPAN() As String Implements TPDotnet.Pos.IEFT.CardPAN
        Get
            Return ""
        End Get
    End Property

    Public Property colObjects() As System.Collections.Hashtable Implements TPDotnet.Pos.IEFT.colObjects
        Get
            Return m_colObjects
        End Get
        Set(ByVal value As System.Collections.Hashtable)
            m_colObjects = value
        End Set
    End Property

    Public Function CreditEftDevice(ByRef taobj As TPDotnet.Pos.TA, ByRef MyTaMediaRec As TPDotnet.Pos.TaMediaRec) As Short Implements TPDotnet.Pos.IEFT.CreditEftDevice

    End Function

    Public ReadOnly Property DoEOD() As Boolean Implements TPDotnet.Pos.IEFT.DoEOD
        Get

        End Get
    End Property

    Public Function DoServiceRequest(ByVal taobj As TPDotnet.Pos.TA, ByVal szTheRequest As String) As Integer Implements TPDotnet.Pos.IEFT.DoServiceRequest

        ' set the default return code
        DoServiceRequest = -1

        Try
            LOG_FuncStart(getLocationString("DoServiceRequest"))

            Select Case szTheRequest

                Case "GetLastPaymentTransactionAmount"
                    GetLastPaymentTransactionAmount()
                    Exit Select

                Case "GetLastPaymentTransactionNumber"
                    GetLastPaymentTransactionNumber()
                    Exit Select

                Case "GetOperationNumber"

                    Exit Select

                Case Else
                    LOG_Warning(PosDef.TARMessageTypes.TPWARNING, getLocationString("DoServiceRequest"), "Unknow request : (" & szTheRequest & ")")
                    Exit Select

            End Select

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DoServiceRequest"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DoServiceRequest"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DoServiceRequest"), "Function returns : " & DoServiceRequest.ToString)
        End Try

    End Function

    Public Function EndOfDay(ByVal taobj As TPDotnet.Pos.TA) As Short Implements TPDotnet.Pos.IEFT.EndOfDay

    End Function

    Public Sub Initialize(ByRef taobj As TPDotnet.Pos.TA, ByVal theModCntr As TPDotnet.Pos.IModCntr) Implements TPDotnet.Pos.IEFT.Initialize

        Try
            LOG_FuncStart(getLocationString("Initialize"))

            m_taobj = taobj
            m_theModCntr = theModCntr

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("Initialize"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("Initialize"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("Initialize"), "Function returns ")
        End Try
    End Sub

    Public Function Login() As Short Implements TPDotnet.Pos.IEFT.Login

    End Function

    Public Function Logoff(ByRef taobj As TPDotnet.Pos.TA) As Short Implements TPDotnet.Pos.IEFT.Logoff
        EFTClosureHandler(taobj, Nothing, False, True, True)
    End Function

    Public Function OnlineAgent(ByRef taobj As TPDotnet.Pos.TA, ByRef MyOnlineAgent As TPDotnet.Pos.TaOnlineAgent) As Short Implements TPDotnet.Pos.IEFT.OnlineAgent

    End Function

    Public ReadOnly Property PaddingChar() As String Implements TPDotnet.Pos.IEFT.PaddingChar
        Get
            Return " "
        End Get
    End Property


    Public ReadOnly Property SeparatorChar() As String Implements TPDotnet.Pos.IEFT.SeparatorChar
        Get
            Return vbCrLf
        End Get
    End Property

    Public ReadOnly Property Serialize() As Boolean Implements TPDotnet.Pos.IEFT.Serialize
        Get

        End Get
    End Property

    Public Function SilentEndOfDay(ByVal taobj As TPDotnet.Pos.TA) As Short Implements TPDotnet.Pos.IEFT.SilentEndOfDay
        'EFTClosureHandler(taobj, Nothing, False, True, False)
    End Function

    Public Function SilentLogoff(ByRef taobj As TPDotnet.Pos.TA) As Short Implements TPDotnet.Pos.IEFT.SilentLogoff
        'EFTClosureHandler(taobj, Nothing, False, True, False)
    End Function

    Public Sub Terminate() Implements TPDotnet.Pos.IEFT.Terminate
        ' called exiting from application
        'EFTClosureHandler(taobj, Nothing, False, True, False)
    End Sub


    Public Function TicketReprint(ByVal taobj As TPDotnet.Pos.TA) As Short Implements TPDotnet.Pos.IEFT.TicketReprint

    End Function


End Class