Imports System
Imports System.IO.Ports
Imports System.Threading
Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Runtime.InteropServices
Imports TPDotnet.Pos

Public Class Setefi

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

#Region "Events"
    Public Event DaSistemiStatusChanged(ByVal state As Integer, ByVal message As String)
    Public Event DaSistemiRemainingSecond(ByVal seconds As Integer)
    Public Event DaSistemiCancelOperationIsAvailableChanged(ByVal available As Boolean)
#End Region

#Region "Thread Variables"
    Private AbortP As Boolean = False
    Private thread As Thread = Nothing
#End Region

#Region "Serial Variables"
    Private Shared EFT As SerialPort = Nothing
#End Region

#Region "State Variables"
    Protected RxState As RxStates
    Protected TxState As TxStates
    Protected m_State As States
#End Region

#Region "Debug messages"
    Protected m_DebugRxMessages As Boolean = False
#End Region

#Region "Useful Internal Variables"
    Protected m_IsGift As Boolean
    Protected m_TheGiftPAN As String
    Protected m_IsRCard As Boolean
    Protected m_TheRCardPAN As String
    Protected m_IsVoid As Boolean
    Protected m_Amount As Double
    Protected m_TheModCntr As ModCntr
    Protected m_DialogActiv As Boolean
    Protected m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI As ESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI
    Protected m_TheSCONTRINO_DA_TERMINALE As SCONTRINO_DA_TERMINALE
    Protected m_TheRICHIESTA_IMPORTO As RICHIESTA_IMPORTO
    Protected m_PaymentMessageToUse As PaymentMessages

    Protected Const lWorkstationNmbr As Integer = 0
#End Region

#Region "Useful properties availables to the caller"
    Public ReadOnly Property DatiAutorizzazioneECarta() As ESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI
        Get
            Return m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI
        End Get
    End Property

    Public ReadOnly Property ScontrinoEFT() As SCONTRINO_DA_TERMINALE
        Get
            Return m_TheSCONTRINO_DA_TERMINALE
        End Get
    End Property

    Public ReadOnly Property Stato() As States
        Get
            Return m_State
        End Get
    End Property

    Public Property PaymentMessagesToUse() As PaymentMessages
        Get
            Return m_PaymentMessageToUse
        End Get
        Set(ByVal value As PaymentMessages)
            m_PaymentMessageToUse = value
        End Set
    End Property

    Public Property DialogActiv() As Boolean
        Get
            Return m_DialogActiv
        End Get
        Set(ByVal value As Boolean)
            m_DialogActiv = value
        End Set
    End Property
#End Region

#Region "ReadOnly Variables"
    ReadOnly MAX_MSG_SIZE As Integer = 1024
    ReadOnly STATUS_MESSAGE_LEN As Integer = 22
    ReadOnly SOH As Byte = &H1 ' token for start of status message
    ReadOnly EOT As Byte = &H4 ' token for end of status message
    ReadOnly STX As Byte = &H2 ' token for start of application message
    ReadOnly ETX As Byte = &H3 ' token for end of application message
    ReadOnly ACK As Byte = &H6 ' token for start of ACK message
    ReadOnly NAK As Byte = &H15 ' token for start of ACK message
    ReadOnly ESC As Byte = &H1B ' token for end of ticket message
    ReadOnly ACK_CMD() As Byte = {&H6, &H3, &H7A} ' the ACK message within ETX and LRC
    ReadOnly NAK_CMD() As Byte = {&H15, &H3, &H69} ' the NAK message within ETX and LRC
    ReadOnly END_TICKET() As Byte = {&H7D, &H1B} ' the NAK message within ETX and LRC
#End Region

#Region "Enum"
    Enum States
        SETEFI_ERROR = -1
        SETEFI_SUCCESS = 0
        SETEFI_IN_PROGRESS = 1
    End Enum

    Enum RxStates
        RX_START_TRANSACTION_REQUEST
        RX_ACK
        RX_TRANSACTION_RESULT
        RX_STATUS_MESSAGE
        RX_EFT_TICKET
    End Enum

    Enum PaymentMessages
        TX_ACTIVATE_PAYMENT = TxStates.TX_ACTIVATE_PAYMENT
        TX_ACTIVATE_PAYMENT_MSG_O = TxStates.TX_ACTIVATE_PAYMENT_MSG_O
    End Enum

    Enum TxStates
        TX_ACK
        TX_TRANSACTION_AMOUNT
        TX_ENABLE_PRINT_ON_ECR
        TX_DISABLE_PRINT_ON_ECR
        TX_ACTIVATE_PAYMENT
        TX_ACTIVATE_PAYMENT_MSG_O
        TX_ACTIVATE_VOID
        TX_NAK
        TX_CLOSURE
        TX_TOTALS_AND_CLOSURE
    End Enum

    Enum KindOfException
        MSG_IS_INCORRECT ' send nak and read again
        NAK_RECEIVED     ' write and read again
        ABORT_OPERATION
        TRANSACTION_FAILED
    End Enum
#End Region

#Region "RX Structure"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Public Structure ESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI
        Implements COMMON_RX_STRUCTURE

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=2)> Public ESITO_TRANSAZIONE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=19)> Public PAN As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public FISSO_1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=2)> Public CODICE_SOCIETA_EMETTITRICE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=4)> Public SCADENZA_CARTA_MMAA As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=15)> Public CODICE_ACQUIRER_PAGOBANCOMAT As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=9)> Public FISSO_2 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=6)> Public FISSO_3 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=7)> Public NUMERO_PROGRESSIVO_OPERAZIONE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=9)> Public CODICE_AUTORIZZAZIONE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=16)> Public NOME_ACQUIRER As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public TIPO_CARTA As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public FISSO_4 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=20)> Public FISSO_5 As String

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure SCONTRINO_DA_TERMINALE
        Implements COMMON_RX_STRUCTURE

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=200)> Public LINEE_DA_STAMPARE As String

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure RICHIESTA_IMPORTO
        Implements COMMON_RX_STRUCTURE

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String

    End Structure

    Protected Interface COMMON_RX_STRUCTURE ' more dummy then common!
    End Interface
#End Region

#Region "TX Structure"

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure AVVIO_PAGAMENTO
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=3)> Public RISERVATO2 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CAMPI_AGGIUNTIVI_1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO3 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CAMPI_AGGIUNTIVI_CARTA_PAGAMENTO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=3)> Public RISERVATO4 As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure AVVIO_PAGAMENTO_MSG_O

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=7)> Public CODICE_COMMERCIANTE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=2)> Public NUMERO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=3)> Public CODICE_PRODOTTO_FIDELITY As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CAMPI_AGGIUNTIVI_1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public TIPO_CARTA_SCELTA As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=11)> Public GIFT_PAN As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CAMPI_AGGIUNTIVI_CARTA_PAGAMENTO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CAMPO_PAN_MASCHERATO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=2)> Public FILLER As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CONTROLLO_CAMPI_CARTA As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=4)> Public CAMPI_CARTA As String

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure AVVIO_STORNO

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO2 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO3 As String

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure TOTALI_CHIUSURA

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure CHIUSURA

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public IDENTIFICATIVO_TERMINALE As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public RISERVATO1 As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=1)> Public CODICE_MESSAGGIO As String

    End Structure

#End Region

#Region "Structure Functions"
    Protected Function PutRxDataIntoStruct(ByVal RxString As String, ByRef RxStruct As COMMON_RX_STRUCTURE) As Boolean

        Dim MyPtr As IntPtr
        PutRxDataIntoStruct = False
        Try
            MyPtr = Marshal.StringToBSTR(RxString)
            RxStruct = Marshal.PtrToStructure(MyPtr, RxStruct.GetType)
            PutRxDataIntoStruct = True
        Catch ex As Exception
            PutRxDataIntoStruct = False
        Finally
            Marshal.FreeBSTR(MyPtr)
        End Try

    End Function

    Protected Sub ClearStructures()
        Dim ClearString As String
        Try
            ClearString = New String(" ", Marshal.SizeOf(m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI))
            PutRxDataIntoStruct(ClearString, m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI)
            ClearString = New String(" ", Marshal.SizeOf(m_TheRICHIESTA_IMPORTO))
            PutRxDataIntoStruct(ClearString, m_TheRICHIESTA_IMPORTO)
            ClearString = New String(" ", Marshal.SizeOf(m_TheSCONTRINO_DA_TERMINALE))
            PutRxDataIntoStruct(ClearString, m_TheSCONTRINO_DA_TERMINALE)
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Serial Functions"
    Protected Sub ListSerial()
        Dim ports As String() = SerialPort.GetPortNames()

        'Console.WriteLine("The following serial ports were found:")

        Dim port As String
        For Each port In ports
            'Console.WriteLine(port)
        Next port

    End Sub

    Private Function CloseEFTSerial() As Boolean

        If EFT Is Nothing Then
            Return True
        End If

        For i As Integer = 0 To 3
            Try
                If EFT.IsOpen Then
                    EFT.Dispose()
                    EFT.Close()
                    System.Threading.Thread.Sleep(500)
                Else
                    Exit For
                End If
            Catch ex As Exception
                Throw New EFTException("Chiusura porta : " + ex.Message)
                If i = 3 Then
                    Return False
                End If
                Continue For
            End Try
        Next

        EFT = Nothing

        Return True

    End Function

    Protected Sub ReadIniFile()

        Dim document As XmlDocument
        Dim modelList As XmlNodeList
        Dim modelNode As XmlNode

        Try
            document = New XmlDocument()
            document.Load(getPosConfigurationPath() + "\\Eft.xml")

            modelList = document.SelectNodes("/Eft/Setting")
            For Each modelNode In modelList

                Dim Protocol As String = modelNode.Attributes.GetNamedItem("Protocol").Value
                If Protocol.Equals("Protocol21") Then

                    ' Default
                    m_PaymentMessageToUse = PaymentMessages.TX_ACTIVATE_PAYMENT


                    Dim paramList As XmlNodeList = modelNode.ChildNodes
                    Dim paramNode As XmlNode
                    For Each paramNode In paramList

                        If paramNode.Name.Equals("Com") Then

                            EFT.PortName = paramNode.InnerText

                        ElseIf paramNode.Name.Equals("Baud") Then

                            EFT.BaudRate = paramNode.InnerText

                        ElseIf paramNode.Name.Equals("Parity") Then

                            Select Case paramNode.InnerText

                                Case "Even"
                                    EFT.Parity = System.IO.Ports.Parity.Even
                                Case "Mark"
                                    EFT.Parity = System.IO.Ports.Parity.Mark
                                Case "None"
                                    EFT.Parity = System.IO.Ports.Parity.None
                                Case "Odd"
                                    EFT.Parity = System.IO.Ports.Parity.Odd
                                Case "Space"
                                    EFT.Parity = System.IO.Ports.Parity.Space

                            End Select

                        ElseIf paramNode.Name.Equals("DataBits") Then

                            EFT.DataBits = paramNode.InnerText

                        ElseIf paramNode.Name.Equals("StopBits") Then

                            EFT.StopBits = paramNode.InnerText

                        ElseIf paramNode.Name.Equals("DebugTxRxMessages") Then

                            m_DebugRxMessages = CBool(paramNode.InnerText)

                        ElseIf paramNode.Name.Equals("PaymentMessageToUse") Then

                            m_PaymentMessageToUse = PaymentMessages.TX_ACTIVATE_PAYMENT
                            If (paramNode.InnerText.ToUpper = "O") Then
                                m_PaymentMessageToUse = PaymentMessages.TX_ACTIVATE_PAYMENT_MSG_O
                            End If

                        End If

                    Next

                End If

            Next
        Catch errorVariable As Exception
            'Console.Write(errorVariable.ToString())
        End Try

    End Sub

    Protected Function OpenEFTSerial() As Boolean

        If EFT.IsOpen Then
            EFT.DiscardInBuffer()
            EFT.DiscardOutBuffer()
            Return True
        End If

        For i As Integer = 0 To 3
            Try
                EFT.Open()
                Exit For
            Catch ex As Exception
                Throw New EFTException("Apertura porta: " + ex.Message)
                If i > 3 Then
                    Return False
                End If
                System.Threading.Thread.Sleep(500)
                Continue For
            End Try
        Next

        EFT.DiscardInBuffer()
        EFT.DiscardOutBuffer()

        Return True

    End Function

    Protected Function IsConnected() As Boolean

        'If EFT.IsOpen Then
        'Return (EFT.CDHolding Or EFT.CtsHolding Or EFT.DsrHolding)
        'Else
        'Return False
        'End If
        Return EFT.IsOpen

    End Function
#End Region

#Region "Read From EFT"
    Private indexSOH As Integer = -1
    Private indexEOT As Integer = -1
    Private theStatusMsgIsSplitted As Boolean = False
    Dim theStatusMsg(1024) As Byte
    Private Function ReadFromEFT(ByVal what As Integer, ByVal timeout As Integer) As Boolean

        Dim currentBytesNo As Integer = 0
        Dim currentBytes(1024) As Byte
        Dim totalBytesNo As Integer = 0
        Dim totalBytes(1024) As Byte
        Dim EndOfPacketAlreadyFound As Boolean
        Dim RetryNo As Integer = 1
        Dim EndOfPacket As Byte = ETX

        Me.RxState = what

        If Not IsConnected() Then
            Throw New EFTException("EFT non presente", KindOfException.ABORT_OPERATION)
        End If

        Dim startDateTime As DateTime = DateTime.Now

        EFT.ReadTimeout = 1000
        EndOfPacketAlreadyFound = False
        System.Array.Clear(totalBytes, 0, totalBytes.Length - 1)
        totalBytesNo = 0

        If m_DebugRxMessages Then
            LOG_Error(getLocationString("ReadFromEFT"), "Waiting for message : " & Me.RxState.ToString)
        End If

        While True

            Try
                System.Array.Clear(currentBytes, 0, currentBytes.Length - 1)
                currentBytesNo = EFT.Read(currentBytes, 0, currentBytes.Length - 1)

                If m_DebugRxMessages Then
                    Try
                        Dim dbg As String = BitConverter.ToString(currentBytes, 0, currentBytesNo)
                        LOG_Error(getLocationString("ReadFromEFT"), "<" & dbg & ">")
                    Catch ex As Exception
                        LOG_Error(getLocationString("ReadFromEFT"), "Exception debugging the received message")
                    End Try
                End If

                ' Checking the received buffer
                If currentBytesNo < 0 Then
                    Throw New EFTException("Errore lettura dati dalla porta specificata!", KindOfException.ABORT_OPERATION)
                ElseIf ((totalBytesNo + currentBytesNo) > MAX_MSG_SIZE) Then
                    Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(Too many long message)", KindOfException.MSG_IS_INCORRECT)
                End If

                System.Array.ConstrainedCopy(currentBytes, 0, totalBytes, totalBytesNo, currentBytesNo)
                totalBytesNo += currentBytesNo

                indexEOT = System.Array.IndexOf(totalBytes, EOT)
                If indexEOT <> -1 Then ' the current message contains the EOT

                    Dim tmpIndexSOH As Integer = System.Array.IndexOf(totalBytes, SOH)
                    If tmpIndexSOH <> -1 Then
                        ' the current message contains both SOH & EOT
                        indexSOH = tmpIndexSOH
                        System.Array.Clear(theStatusMsg, 0, theStatusMsg.Length - 1)
                        System.Array.ConstrainedCopy(totalBytes, indexSOH, theStatusMsg, 0, indexEOT)
                        Dim msg As String = ByteArrayToStr(theStatusMsg)
                        UpdateState(msg, GetStatus()) ' we don't change the status
                        System.Array.Clear(theStatusMsg, 0, theStatusMsg.Length - 1)
                        indexSOH = -1
                        indexEOT = -1
                        theStatusMsgIsSplitted = False
                    Else
                        ' the current message doen't contanins the SOH, it should be splitted
                        If theStatusMsgIsSplitted Then
                            ' the message is really splitted
                            Dim msg As String = ByteArrayToStr(theStatusMsg)
                            System.Array.Clear(theStatusMsg, 0, theStatusMsg.Length - 1)
                            System.Array.ConstrainedCopy(totalBytes, 0, theStatusMsg, 0, indexEOT)
                            Dim ind As Integer = msg.IndexOf(Nothing)
                            Dim msg2 As String = msg.Substring(0, ind) & ByteArrayToStr(theStatusMsg)
                            UpdateState(msg2, GetStatus()) ' we don't change the status
                            System.Array.Clear(theStatusMsg, 0, theStatusMsg.Length - 1)
                            indexSOH = -1
                            indexEOT = -1
                            theStatusMsgIsSplitted = False
                        Else
                            ' error the message is not splitted
                            System.Array.Clear(theStatusMsg, 0, theStatusMsg.Length - 1)
                            indexSOH = -1
                            indexEOT = -1
                            theStatusMsgIsSplitted = False
                        End If
                    End If

                    ' check if the message is splitted
                    'If theStatusMsgIsSplitted Then
                    '    theStatusMsgIsSplitted = False
                    'Else
                    'End If
                Else
                    indexSOH = System.Array.IndexOf(totalBytes, SOH)
                    If indexSOH <> -1 Then
                        ' found SOH without EOT: save the message and set as splitted
                        theStatusMsgIsSplitted = True
                        System.Array.Clear(theStatusMsg, 0, theStatusMsg.Length - 1)
                        System.Array.ConstrainedCopy(totalBytes, indexSOH, theStatusMsg, 0, totalBytesNo)
                    Else
                        ' this message does not contains any status message: do nothing
                    End If
                End If

                Dim indexETX As Integer = System.Array.IndexOf(totalBytes, ETX)
                If indexETX = -1 Then 'And Not EndOfPacketAlreadyFound Then ' Did I receive the EndOfPacket ?
                    Continue While ' No, I didn't
                Else ' Yes, I did

                    ' check if we have to loop again in order to wait the crc
                    If ((indexETX + 1) = totalBytesNo) Then
                        Continue While ' we received the etx but we miss the last bytes : the crc
                    End If

                    ' check the appropriate start character
                    Dim startToken As Byte
                    If what = RxStates.RX_ACK Then
                        startToken = ACK
                    Else
                        startToken = STX
                    End If

                    Dim startTokenIndex As Integer = -1

                    ' check if the start character is present
                    startTokenIndex = System.Array.IndexOf(totalBytes, startToken)
                    If startTokenIndex = -1 Then
                        Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(No start token)", KindOfException.MSG_IS_INCORRECT)
                    End If

                    ' get only the message contained between the start and the end character
                    Dim theMsg(1024) As Byte
                    System.Array.Clear(theMsg, 0, theMsg.Length - 1)
                    System.Array.ConstrainedCopy(totalBytes, startTokenIndex, theMsg, 0, indexETX + 2)
                    Dim crc As Byte = Me.CalculateCRC(totalBytes, startTokenIndex, indexETX)
                    If crc = totalBytes(indexETX + 1) Then
                        System.Array.Clear(totalBytes, 0, totalBytes.Length - 1)
                        System.Array.ConstrainedCopy(theMsg, 0, totalBytes, 0, totalBytes.Length)
                        totalBytesNo = indexETX + 2
                        Exit While
                    Else
                        Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(CRC)", KindOfException.MSG_IS_INCORRECT)
                    End If

                End If

            Catch ex As TimeoutException
                Dim ts As TimeSpan = DateTime.Now.Subtract(startDateTime)
                If ts.TotalSeconds < timeout Then
                    RaiseEvent DaSistemiRemainingSecond((timeout - ts.TotalSeconds))
                Else
                    Throw New EFTException("Tempo scaduto!", KindOfException.ABORT_OPERATION)
                End If
                If AbortP Then
                    Throw New EFTException("Interrotto dall'operatore!", KindOfException.ABORT_OPERATION)
                End If
            Catch ex As Exception
                Throw New EFTException("Lettura dati : " + ex.Message, KindOfException.ABORT_OPERATION)
            End Try

        End While

        If totalBytesNo = 0 Then
            Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(Read 0 bytes)", KindOfException.ABORT_OPERATION)
        End If

        If (String.Compare(Me.ByteArrayToStr(totalBytes), Me.ByteArrayToStr(Me.NAK_CMD)) = 0) Then
            Throw New EFTException("NAK!", KindOfException.NAK_RECEIVED)
        End If

        Select Case what

            Case RxStates.RX_ACK
                If Not (String.Compare(Me.ByteArrayToStr(totalBytes), Me.ByteArrayToStr(Me.ACK_CMD)) = 0) Then
                    Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(Expected ACK)", KindOfException.MSG_IS_INCORRECT)
                End If
                Exit Select

            Case RxStates.RX_START_TRANSACTION_REQUEST
                Dim RxMsgAsString As String = ""
                Dim tmp(1024) As Byte
                Array.ConstrainedCopy(totalBytes, 1, tmp, 0, Marshal.SizeOf(m_TheRICHIESTA_IMPORTO))
                RxMsgAsString = Me.ByteArrayToStr(tmp)
                PutRxDataIntoStruct(RxMsgAsString, m_TheRICHIESTA_IMPORTO)
                If m_TheRICHIESTA_IMPORTO.CODICE_MESSAGGIO <> "I" Then
                    Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(Expected I message)", KindOfException.MSG_IS_INCORRECT)
                End If
                Exit Select

            Case RxStates.RX_TRANSACTION_RESULT
                Dim RxMsgAsString As String = ""
                Dim tmp(1024) As Byte
                Array.ConstrainedCopy(totalBytes, 1, tmp, 0, Marshal.SizeOf(m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI))
                RxMsgAsString = Me.ByteArrayToStr(tmp)
                PutRxDataIntoStruct(RxMsgAsString, m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI)
                If m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI.CODICE_MESSAGGIO <> "E" Then
                    Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(Expected E message)", KindOfException.MSG_IS_INCORRECT)
                End If
                If Not m_TheESITO_TRANSAZIONE_DATI_AUTORIZZAZIONE_CARTA_AGGIUNTIVI.ESITO_TRANSAZIONE = "00" Then
                    Throw New EFTException("Transazione non eseguita!", KindOfException.TRANSACTION_FAILED)
                End If
                Exit Select

            Case RxStates.RX_EFT_TICKET
                Dim RxMsgAsString As String = ""
                Dim tmp(1024) As Byte
                Array.ConstrainedCopy(totalBytes, 1, tmp, 0, Marshal.SizeOf(m_TheSCONTRINO_DA_TERMINALE))
                RxMsgAsString = Me.ByteArrayToStr(tmp)
                PutRxDataIntoStruct(RxMsgAsString, m_TheSCONTRINO_DA_TERMINALE)
                If m_TheSCONTRINO_DA_TERMINALE.CODICE_MESSAGGIO <> "S" Then
                    Throw New EFTException("Messaggio non corretto!" & vbNewLine & "(Expected S message)", KindOfException.MSG_IS_INCORRECT)
                End If
                Exit Select

        End Select

        Return True

    End Function
#End Region

#Region "Write to EFT"
    Private Function WriteToEFT(ByVal what As Integer) As Boolean

        Dim txBuffer(1024) As Byte
        Dim count As Integer = 0
        TxState = what
        'Console.WriteLine("writeToDaSistemi : {0}", TxState)

        If Not IsConnected() Then
            Throw New EFTException("Scrittura dati : EFT non presente", KindOfException.ABORT_OPERATION)
            Return False
        End If

        EFT.WriteTimeout = 3000
        System.Array.Clear(txBuffer, 0, txBuffer.Length - 1)

        Select Case TxState

            Case TxStates.TX_ACK
                count = BuildTxACKMessage(txBuffer)

            Case TxStates.TX_TRANSACTION_AMOUNT
                count = BuildTxTransactionAmountMessage(txBuffer)

            Case TxStates.TX_ENABLE_PRINT_ON_ECR
                count = BuildEnablePrintOnECRMessage(txBuffer)

            Case TxStates.TX_DISABLE_PRINT_ON_ECR
                count = BuildDisablePrintOnECRMessage(txBuffer)

            Case TxStates.TX_ACTIVATE_PAYMENT
                count = BuildActivatePaymentMessage(txBuffer)

            Case TxStates.TX_ACTIVATE_PAYMENT_MSG_O
                count = BuildActivatePaymentMsgOMessage(txBuffer)

            Case TxStates.TX_ACTIVATE_VOID
                count = BuildActivateVoidMessage(txBuffer)

            Case TxStates.TX_NAK
                count = BuildTxNAKMessage(txBuffer)

            Case TxStates.TX_CLOSURE
                count = BuildTxClosureMessage(txBuffer)

            Case TxStates.TX_TOTALS_AND_CLOSURE
                count = BuildTxTotalsAndClosureMessage(txBuffer)

        End Select

        Try

            If m_DebugRxMessages Then
                LOG_Error(getLocationString("WriteToEFT"), "Sending message : " & Me.TxState.ToString)
            End If

            EFT.Write(txBuffer, 0, count)

            If m_DebugRxMessages Then
                Try
                    Dim dbg As String = BitConverter.ToString(txBuffer, 0, count)
                    LOG_Error(getLocationString("WriteToEFT"), "<" & dbg & ">")
                Catch ex As Exception
                    LOG_Error(getLocationString("WriteToEFT"), "Exception debugging the transmitted message")
                End Try
            End If

            'Console.WriteLine("Sent message : " + ToHexString(txBuffer, count))
            'Console.WriteLine("Sent message length : {0}", count)
        Catch ex As Exception
            Throw New EFTException(" Trasmissione importo : " + ex.Message, KindOfException.ABORT_OPERATION)
            Return False
        End Try

        Return True

    End Function
#End Region

#Region "CRC"
    Private Function CalculateCRC(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Byte

        Dim i As Integer
        Dim crc As Byte = &H7F
        For i = offset To count
            crc = (crc Xor buffer(i))
        Next i

        'Console.WriteLine("CRC = " + Hex(crc))
        Return crc

    End Function
#End Region

#Region "ACK/NAK Trasmission"
    Private Function TxACK() As Boolean

        Try
            EFT.Write(ACK_CMD, 0, ACK_CMD.Length)
        Catch ex As Exception
            Throw New EFTException("Trasmissione ACK : " + ex.Message)
            Return False
        End Try

        Return True

    End Function

    Private Function TxNAK() As Boolean

        Try
            EFT.Write(NAK_CMD, 0, NAK_CMD.Length)
        Catch ex As Exception
            Throw New EFTException("Trasmissione NAK : " + ex.Message)
            Return False
        End Try

        Return True

    End Function
#End Region

#Region "Useful string functions"
    Public Shared Function ToHexString(ByVal bytes() As Byte, ByVal count As Integer) As String

        Dim hexStr As String = ""
        Dim i As Integer
        For i = 0 To count - 1
            hexStr = hexStr + "<" + Hex(bytes(i)) + ">"
        Next i
        Return hexStr

    End Function

    Private Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.ASCIIEncoding()
        Return encoding.GetBytes(str)
    End Function

    Private Function ByteArrayToStr(ByVal bytes() As Byte) As String
        Dim encoding As New System.Text.ASCIIEncoding()
        ByteArrayToStr = encoding.GetString(bytes)
    End Function

#End Region

#Region "Build trasmission buffer"
    Private Function BuildTxACKMessage(ByRef b() As Byte) As Integer

        System.Array.Copy(Me.ACK_CMD, b, Me.ACK_CMD.Length)
        Return Me.ACK_CMD.Length

    End Function

    Private Function BuildTxNAKMessage(ByRef b() As Byte) As Integer

        System.Array.Copy(Me.NAK_CMD, b, Me.NAK_CMD.Length)
        Return Me.NAK_CMD.Length

    End Function

    Private Function BuildTxTotalsAndClosureMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        Dim MyCHIUSURA As New CHIUSURA

        With MyCHIUSURA
            .IDENTIFICATIVO_TERMINALE = Format(lWorkstationNmbr, "00000000")
            .RISERVATO1 = "0"
            .CODICE_MESSAGGIO = "T"
        End With

        s = Microsoft.VisualBasic.Chr(STX)
        With MyCHIUSURA
            s += .IDENTIFICATIVO_TERMINALE & _
                .RISERVATO1 & _
                .CODICE_MESSAGGIO
        End With
        s += Microsoft.VisualBasic.Chr(ETX)

        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function

    Private Function BuildTxClosureMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        Dim MyCHIUSURA As New CHIUSURA

        With MyCHIUSURA
            .IDENTIFICATIVO_TERMINALE = Format(lWorkstationNmbr, "00000000")
            .RISERVATO1 = "0"
            .CODICE_MESSAGGIO = "C"
        End With

        s = Microsoft.VisualBasic.Chr(STX)
        With MyCHIUSURA
            s += .IDENTIFICATIVO_TERMINALE & _
                .RISERVATO1 & _
                .CODICE_MESSAGGIO
        End With
        s += Microsoft.VisualBasic.Chr(ETX)

        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function

    Private Function BuildActivateVoidMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        Dim MyAVVIO_STORNO As New AVVIO_STORNO

        With MyAVVIO_STORNO
            .IDENTIFICATIVO_TERMINALE = Format(lWorkstationNmbr, "00000000")
            .RISERVATO1 = "0"
            .CODICE_MESSAGGIO = "S"
            .RISERVATO2 = "0"
            .RISERVATO3 = "0"
        End With

        s = Microsoft.VisualBasic.Chr(STX)
        With MyAVVIO_STORNO
            s += .IDENTIFICATIVO_TERMINALE & _
                .RISERVATO1 & _
                .CODICE_MESSAGGIO & _
                .RISERVATO2 & _
                .RISERVATO3
        End With
        s += Microsoft.VisualBasic.Chr(ETX)

        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function


    Private Function BuildActivatePaymentMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        Dim MyAVVIO_PAGAMENTO As New AVVIO_PAGAMENTO

        With MyAVVIO_PAGAMENTO
            .IDENTIFICATIVO_TERMINALE = Format(lWorkstationNmbr, "00000000")
            .RISERVATO1 = "0"
            .CODICE_MESSAGGIO = "P"
            .RISERVATO2 = "000"
            .CAMPI_AGGIUNTIVI_1 = "2"
            .RISERVATO3 = "0"
            .CAMPI_AGGIUNTIVI_CARTA_PAGAMENTO = "2"
            .RISERVATO4 = "000"
        End With

        s = Microsoft.VisualBasic.Chr(STX)
        With MyAVVIO_PAGAMENTO
            s += .IDENTIFICATIVO_TERMINALE & _
            .RISERVATO1 & _
            .CODICE_MESSAGGIO & _
            .RISERVATO2 & _
            .CAMPI_AGGIUNTIVI_1 & _
            .RISERVATO3 & _
            .CAMPI_AGGIUNTIVI_CARTA_PAGAMENTO & _
            .RISERVATO4
        End With
        s += Microsoft.VisualBasic.Chr(ETX)

        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function

    Private Function BuildActivatePaymentMsgOMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        Dim MyAVVIO_PAGAMENTO_MSG_O As New AVVIO_PAGAMENTO_MSG_O

        ' notes from setefi documentation
        ' All payments
        '   - TIPO_CARTA_SCELTA = "9"
        '   - GIFT_PAN = "<Prm = ExternalID from Media>"
        '   - CONTROLLO_CAMPI_CARTA = "0"
        '   - CAMPI_CARTA = "0000"
        ' Gift Card (Lista Nozze)
        '   - TIPO_CARTA_SCELTA = "8"
        '   - GIFT_PAN = "<ExternalID from Media>"
        '   - CONTROLLO_CAMPI_CARTA = "0"
        '   - CAMPI_CARTA = "0000"
        ' R. Card
        '   - TIPO_CARTA_SCELTA = "9"
        '   - GIFT_PAN = "Prm = ExternalID from Media"
        '   - CONTROLLO_CAMPI_CARTA = "1"
        '   - CAMPI_CARTA = "<Last 4 chars from customer code>"
        With MyAVVIO_PAGAMENTO_MSG_O
            .CODICE_COMMERCIANTE = New String("0", 7)
            .NUMERO_TERMINALE = Format(lWorkstationNmbr, "00")
            .CODICE_MESSAGGIO = "O"
            .CODICE_PRODOTTO_FIDELITY = "000"
            .CAMPI_AGGIUNTIVI_1 = "2"
            .TIPO_CARTA_SCELTA = "9"
            If m_IsGift Then
                MyAVVIO_PAGAMENTO_MSG_O.TIPO_CARTA_SCELTA = "8"
            End If
            .GIFT_PAN = m_TheGiftPAN ' it is a 11 length string with all "0" chars or the gift pan
            .CAMPI_AGGIUNTIVI_CARTA_PAGAMENTO = "2"
            .CAMPO_PAN_MASCHERATO = "1"
            .FILLER = "00"
            .CONTROLLO_CAMPI_CARTA = "0"
            If m_IsRCard Then
                .CONTROLLO_CAMPI_CARTA = "2" ' 1=check 4 chars; 2=check 3 chars 
            End If
            .CAMPI_CARTA = m_TheRCardPAN ' it is a 4 length string with all "0" chars or the rcard pan
        End With

        s = Microsoft.VisualBasic.Chr(STX)
        With MyAVVIO_PAGAMENTO_MSG_O
            s += .CODICE_COMMERCIANTE & _
                .NUMERO_TERMINALE & _
                .CODICE_MESSAGGIO & _
                .CODICE_PRODOTTO_FIDELITY & _
                .CAMPI_AGGIUNTIVI_1 & _
                .TIPO_CARTA_SCELTA & _
                .GIFT_PAN & _
                .CAMPI_AGGIUNTIVI_CARTA_PAGAMENTO & _
                .CAMPO_PAN_MASCHERATO & _
                .FILLER & _
                .CONTROLLO_CAMPI_CARTA & _
                .CAMPI_CARTA
        End With
        s += Microsoft.VisualBasic.Chr(ETX)

        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function


    Private Function BuildEnablePrintOnECRMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        s = String.Format("{0}{1:D8}0E1{2}", Microsoft.VisualBasic.Chr(STX), Format(lWorkstationNmbr, "00000000"), Microsoft.VisualBasic.Chr(ETX))
        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function

    Private Function BuildDisablePrintOnECRMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        s = String.Format("{0}{1:D8}0E0{2}", Microsoft.VisualBasic.Chr(STX), Format(lWorkstationNmbr, "00000000"), Microsoft.VisualBasic.Chr(ETX))
        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function

    Private Function BuildTxTransactionAmountMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        s = String.Format("{0}{1:D8}0I{2:D8}1000000000{3}", Microsoft.VisualBasic.Chr(STX), Format(lWorkstationNmbr, "00000000"), Format(m_Amount * 100, "0").PadLeft(8, "0"), Microsoft.VisualBasic.Chr(ETX))
        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function
#End Region

#Region "Thread Functions"

    Protected Overridable Function DoStep( _
        ByRef DoTx As Boolean, ByRef TxState As TxStates, ByRef TxMsg As String, ByRef TxMsgState As States, _
        ByRef DoRx As Boolean, ByRef RxState As TxStates, ByRef RxMsg As String, ByRef RxMsgState As States, ByRef RxTimeout As Integer, _
        ByRef TxFirst As Boolean, ByRef NumRetry As Integer) As Boolean

        Dim Tx As Boolean = DoTx ' handled via Tx because we can receive a NAK and, in that case, we need to tx again
        Dim NAK_Count As Integer = 1

        DoStep = False

        For i As Integer = 0 To NumRetry
            Try
                If TxFirst Then
                    If Tx Then
                        UpdateState(TxMsg, TxMsgState)
                        Me.WriteToEFT(TxState)
                    End If
                    If DoRx Then
                        UpdateState(RxMsg, RxMsgState)
                        Me.ReadFromEFT(RxState, RxTimeout)
                    End If
                Else
                    If DoRx Then
                        UpdateState(RxMsg, RxMsgState)
                        Me.ReadFromEFT(RxState, RxTimeout)
                    End If
                    If Tx Then
                        UpdateState(TxMsg, TxMsgState)
                        Me.WriteToEFT(TxState)
                    End If
                End If
                Exit For
            Catch ex As EFTException
                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
                    If i >= NumRetry Then
                        Throw ex
                    End If
                    WriteToEFT(TxStates.TX_NAK)
                    Tx = False
                    Continue For
                ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
                    If NAK_Count >= NumRetry Then
                        Throw ex
                    End If
                    NAK_Count += 1
                    Tx = DoTx
                    Continue For
                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
                    Throw ex
                Else
                    Throw ex
                End If
            End Try
        Next

        DoStep = True

    End Function

    Protected Sub Pay()
        AbortP = False
        Dim MustSendAbort As Boolean = True
        Dim Tx As Boolean = True
        Dim NAK_Count As Integer = 1

        Dim bContinue As Boolean = True

        Try

            Dim msg As String = getPosTxtNew((m_TheModCntr.contxt), "UserMessage", TXT_EFT_START_PAYMENT)
            UpdateState("Apertura EFT", States.SETEFI_IN_PROGRESS)
            Me.OpenEFTSerial()

            '
            ' 1. enable the print of the eft receipt on the ecr printer
            '
            bContinue = DoStep(True, TxStates.TX_ENABLE_PRINT_ON_ECR, "Trasmissione abilitazione stampa", States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_ACK, "Ricezione conferma fisica", States.SETEFI_IN_PROGRESS, 30, True, 3)
            LOG_Error(getLocationString("Pay"), "Step 1")

            '
            ' 2. send an activation payment request to the eft
            '
            bContinue = DoStep(True, IIf(m_IsVoid, TxStates.TX_ACTIVATE_VOID, m_PaymentMessageToUse), IIf(m_IsVoid, "Trasmissione attivazione storno", "Trasmissione attivazione pagamento"), States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_ACK, "Ricezione conferma fisica", States.SETEFI_IN_PROGRESS, 30, True, 3)
            LOG_Error(getLocationString("Pay"), "Step 2")

            '
            ' 3. receive a request of amount message from the eft
            '
            bContinue = DoStep(True, TxStates.TX_ACK, "Trasmissione conferma fisica", States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_START_TRANSACTION_REQUEST, "Ricezione richiesta importo", States.SETEFI_IN_PROGRESS, 60, False, 3)
            LOG_Error(getLocationString("Pay"), "Step 3")

            '
            ' 4. send the transaciotn amount and read the confirm from the eft
            '
            bContinue = DoStep(True, TxStates.TX_TRANSACTION_AMOUNT, "Trasmissione importo", States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_ACK, "Ricezione conferma fisica", States.SETEFI_IN_PROGRESS, 60, True, 3)
            LOG_Error(getLocationString("Pay"), "Step 4")

            ' At this point the oprator can cancel the operation only directly pressing the red button on the eft terminal
            CancelOperationIsAvailable(False)

            '
            ' 5. receive the transaction result
            '
            Try
                bContinue = DoStep(True, TxStates.TX_ACK, "Trasmissione conferma fisica", States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_TRANSACTION_RESULT, "Ricezione esito transazione", States.SETEFI_IN_PROGRESS, 120, False, 3)
                LOG_Error(getLocationString("Pay"), "Step 5")
            Catch ex As EFTException
                If ex.KindOfException = KindOfException.TRANSACTION_FAILED Then
                    Me.WriteToEFT(TxStates.TX_ACK)
                    MustSendAbort = False
                End If
                Throw ex
            End Try

            ' From this point the transaction must be considered as executed
            UpdateState("Transazione eseguita", States.SETEFI_SUCCESS)

            ' Enable again the annulla button
            CancelOperationIsAvailable(True)

            '
            ' 6. confirm the received amount and get the receipt lines
            '
            For i As Integer = 0 To 6

                Try
                    LOG_Error(getLocationString("Pay"), "Ricezione scontrino. Loop : " & i.ToString)
                    bContinue = DoStep(True, TxStates.TX_ACK, "Trasmissione conferma fisica", States.SETEFI_SUCCESS, _
                                True, RxStates.RX_EFT_TICKET, "Ricezione scontrino", States.SETEFI_SUCCESS, 90, False, 1)
                    LOG_Error(getLocationString("Pay"), "Step 6, " & i.ToString)
                    ' confirm
                    'Me.WriteToEFT(TxStates.TX_ACK)
                    ' save the text
                    TheEFTTicket.Add(m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE)
                    Dim END_INDEX As Integer = m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE.IndexOf(Chr(&H1B))
                    If END_INDEX = -1 Then
                        LOG_Error(getLocationString("Pay"), "Ricezione scontrino. End of receipt has not been received, read again...")
                        Continue For
                    Else
                        'Exit For
                        ' The character before the &H1B must be &H7D if this is really the last receipt message
                        If AscW(m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE(END_INDEX - 1)) = &H7D Then
                            ' OK, we collected all receipt messages
                            LOG_Error(getLocationString("Pay"), "Ricezione scontrino. End of receipt has been received, exit...")
                            Exit For
                        Else
                            LOG_Error(getLocationString("Pay"), "Ricezione scontrino. End of receipt has been received, read again...")
                            Continue For
                        End If
                    End If
                    LOG_Error(getLocationString("Pay"), "Ricezione scontrino. Exit for")
                    Exit For
                Catch ex As EFTException
                    If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
                        If i >= 6 Then
                            Throw New EFTException(ex.Message)
                        End If
                        WriteToEFT(TxStates.TX_NAK)
                        Tx = False
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
                        If NAK_Count >= 6 Then
                            Throw New EFTException(ex.Message)
                        End If
                        NAK_Count += 1
                        Tx = True
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
                        Throw New EFTException(ex.Message)
                    End If
                End Try

            Next

        Catch ex As EFTException
            'Console.WriteLine(ex.Message)
            ' If an exception occurs after the transaction is completed, we have to consider the transaction as executed anyway.
            ' In case of error getting the receipt linees we will inform the operator.
            If Not m_State = States.SETEFI_SUCCESS Then
                UpdateState(ex.Message, States.SETEFI_ERROR)
                setErrorMessage(ex.Message)
            End If
        Finally
            If (m_State = States.SETEFI_ERROR) And Me.IsConnected() And MustSendAbort Then
                UpdateState("Annullamento transazione", States.SETEFI_ERROR)
                WriteToEFT(TxStates.TX_NAK) ' 
                WriteToEFT(TxStates.TX_NAK) ' Send 3 nak to abort the current operation
                WriteToEFT(TxStates.TX_NAK) ' 
            End If
            UpdateState("Chiusura EFT", m_State)
            'Me.CloseDaSistemi()
            Me.m_DialogActiv = False
        End Try

    End Sub

    Protected Sub Close()
        AbortP = False
        Dim MustSendAbort As Boolean = True
        Dim Tx As Boolean = True
        Dim NAK_Count As Integer = 1

        Dim bContinue As Boolean = True

        Try

            UpdateState("Apertura EFT", States.SETEFI_IN_PROGRESS)
            Me.OpenEFTSerial()

            '
            ' 1. enable the print of the eft receipt on the ecr printer
            '
            bContinue = DoStep(True, TxStates.TX_ENABLE_PRINT_ON_ECR, "Trasmissione abilitazione stampa", States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_ACK, "Ricezione conferma fisica", States.SETEFI_IN_PROGRESS, 30, True, 3)

            '
            ' 2. send the closure message
            '
            bContinue = DoStep(True, TxStates.TX_CLOSURE, "Trasmissione messaggio chiusura", States.SETEFI_IN_PROGRESS, _
                                True, RxStates.RX_ACK, "Ricezione conferma fisica", States.SETEFI_IN_PROGRESS, 120, True, 3)

            ' From this point the transaction must be considered as executed
            UpdateState("Transazione eseguita", States.SETEFI_SUCCESS)

            ' Enable again the annulla button
            CancelOperationIsAvailable(True)

            '
            ' 6. confirm the received amount and get the receipt lines
            '
            For i As Integer = 0 To 6

                Try

                    bContinue = DoStep(True, TxStates.TX_ACK, "Trasmissione conferma fisica", States.SETEFI_SUCCESS, _
                                True, RxStates.RX_EFT_TICKET, "Ricezione esito transazione", States.SETEFI_SUCCESS, 90, False, 1)
                    ' confirm
                    'Me.WriteToEFT(TxStates.TX_ACK)
                    ' save the text
                    TheEFTTicket.Add(m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE)
                    Dim END_INDEX As Integer = m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE.IndexOf(Chr(&H1B))
                    If END_INDEX = -1 Then
                        Continue For
                    Else
                        'Exit For
                        ' The character before the &H1B must be &H7D if this is really the last receipt message
                        If AscW(m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE(END_INDEX - 1)) = &H7D Then
                            ' OK, we collected all receipt messages
                            Exit For
                        Else
                            Continue For
                        End If
                    End If
                    Exit For
                Catch ex As EFTException
                    If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
                        If i >= 6 Then
                            Throw New EFTException(ex.Message)
                        End If
                        WriteToEFT(TxStates.TX_NAK)
                        Tx = False
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
                        If NAK_Count >= 6 Then
                            Throw New EFTException(ex.Message)
                        End If
                        NAK_Count += 1
                        Tx = True
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
                        Throw New EFTException(ex.Message)
                    End If
                End Try

            Next

        Catch ex As EFTException
            'Console.WriteLine(ex.Message)
            ' If an exception occurs after the transaction is completed, we have to consider the transaction as executed anyway.
            ' In case of error getting the receipt linees we will inform the operator.
            If Not m_State = States.SETEFI_SUCCESS Then
                UpdateState(ex.Message, States.SETEFI_ERROR)
                setErrorMessage(ex.Message)
            End If
        Finally
            If (m_State = States.SETEFI_ERROR) And Me.IsConnected() And MustSendAbort Then
                UpdateState("Annullamento transazione", States.SETEFI_ERROR)
                WriteToEFT(TxStates.TX_NAK) ' 
                WriteToEFT(TxStates.TX_NAK) ' Send 3 nak to abort the current operation
                WriteToEFT(TxStates.TX_NAK) ' 
            End If
            UpdateState("Chiusura EFT", m_State)
            'Me.CloseDaSistemi()
            Me.m_DialogActiv = False
        End Try

    End Sub

    'Protected Sub Pay()
    '    AbortP = False
    '    Dim MustSendAbort As Boolean = True
    '    Dim Tx As Boolean = True
    '    Dim NAK_Count As Integer = 1

    '    Dim bContinue As Boolean = True

    '    Try

    '        Dim msg As String = getPosTxtNew((m_TheModCntr.contxt), "UserMessage", TXT_EFT_START_PAYMENT)
    '        UpdateState("Apertura EFT", States.SETEFI_IN_PROGRESS)
    '        Me.OpenEFTSerial()

    '        '
    '        ' 1. enable the print of the eft receipt on the ecr printer
    '        '
    '        Tx = True
    '        NAK_Count = 1
    '        For i As Integer = 0 To 3
    '            Try
    '                If Tx Then
    '                    UpdateState("Trasmissione abilitazione stampa", States.SETEFI_IN_PROGRESS)
    '                    Me.WriteToEFT(TxStates.TX_ENABLE_PRINT_ON_ECR)
    '                End If
    '                UpdateState("Ricezione conferma fisica", States.SETEFI_IN_PROGRESS)
    '                Me.ReadFromEFT(RxStates.RX_ACK, 30)
    '                Exit For
    '            Catch ex As EFTException
    '                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
    '                    If i >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    WriteToEFT(TxStates.TX_NAK)
    '                    Tx = False
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
    '                    If NAK_Count >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    NAK_Count += 1
    '                    Tx = True
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
    '                    Throw New EFTException(ex.Message)
    '                End If
    '            End Try
    '        Next

    '        '
    '        ' 2. send an activation payment request to the eft
    '        '
    '        Tx = True
    '        NAK_Count = 1
    '        For i As Integer = 0 To 3
    '            Try
    '                If Tx Then
    '                    If m_IsVoid Then
    '                        UpdateState("Trasmissione attivazione storno", States.SETEFI_IN_PROGRESS)
    '                        Me.WriteToEFT(TxStates.TX_ACTIVATE_VOID)
    '                    Else
    '                        UpdateState("Trasmissione attivazione pagamento", States.SETEFI_IN_PROGRESS)
    '                        Me.WriteToEFT(m_PaymentMessageToUse)
    '                    End If
    '                End If
    '                UpdateState("Ricezione conferma fisica", States.SETEFI_IN_PROGRESS)
    '                Me.ReadFromEFT(RxStates.RX_ACK, 30)
    '                Exit For
    '            Catch ex As EFTException
    '                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
    '                    If i >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    WriteToEFT(TxStates.TX_NAK)
    '                    Tx = False
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
    '                    If NAK_Count >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    NAK_Count += 1
    '                    Tx = True
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
    '                    Throw New EFTException(ex.Message)
    '                End If
    '            End Try
    '        Next

    '        ' 3. receive a request of amount message from the eft
    '        'If Not m_IsVoid Then
    '        For i As Integer = 1 To 3

    '            Try
    '                UpdateState("Ricezione richiesta importo", States.SETEFI_IN_PROGRESS)
    '                Me.ReadFromEFT(RxStates.RX_START_TRANSACTION_REQUEST, 60)
    '                Me.WriteToEFT(TxStates.TX_ACK)
    '                Exit For
    '            Catch ex As EFTException
    '                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Or _
    '                   ex.KindOfException = KindOfException.NAK_RECEIVED Then ' NAK not allowed here, handle it as an incorrect message
    '                    If i >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    WriteToEFT(TxStates.TX_NAK)
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
    '                    Throw New EFTException(ex.Message)
    '                End If
    '            End Try
    '        Next

    '        '
    '        ' 4. send the transaciotn amount and read the confirm from the eft
    '        '
    '        Tx = True
    '        NAK_Count = 1
    '        For i As Integer = 0 To 3
    '            Try
    '                If Tx Then
    '                    UpdateState("Trasmissione importo", States.SETEFI_IN_PROGRESS)
    '                    Me.WriteToEFT(TxStates.TX_TRANSACTION_AMOUNT)
    '                End If
    '                UpdateState("Ricezione conferma fisica", States.SETEFI_IN_PROGRESS)
    '                Me.ReadFromEFT(RxStates.RX_ACK, 60)
    '                Exit For
    '            Catch ex As EFTException
    '                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
    '                    If i >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    WriteToEFT(TxStates.TX_NAK)
    '                    Tx = False
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
    '                    If NAK_Count >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    NAK_Count += 1
    '                    Tx = True
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
    '                    Throw New EFTException(ex.Message)
    '                End If
    '            End Try
    '        Next

    '        ' End If ' of isvoid

    '        ' At this point the oprator can cancel the operation only directly pressing the red button on the eft terminal
    '        CancelOperationIsAvailable(False)

    '        ' 5. receive the transaction result
    '        For i As Integer = 0 To 3
    '            Try
    '                UpdateState("Ricezione esito transazione", States.SETEFI_IN_PROGRESS)
    '                Me.ReadFromEFT(RxStates.RX_TRANSACTION_RESULT, 90)
    '                Exit For
    '            Catch ex As EFTException
    '                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Or _
    '                   ex.KindOfException = KindOfException.NAK_RECEIVED Then ' NAK not allowed here, handle it as an incorrect message
    '                    If i >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    WriteToEFT(TxStates.TX_NAK)
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
    '                    Throw New EFTException(ex.Message)
    '                ElseIf ex.KindOfException = KindOfException.TRANSACTION_FAILED Then
    '                    Me.WriteToEFT(TxStates.TX_ACK)
    '                    MustSendAbort = False
    '                    Throw New EFTException(ex.Message)
    '                End If

    '            End Try
    '        Next

    '        ' From this point the transaction must be considered as executed
    '        UpdateState("Transazione eseguita", States.SETEFI_SUCCESS)
    '        ' Enable again the annulla button
    '        CancelOperationIsAvailable(True)

    '        ' 6. confirm the received amount and get the receipt lines
    '        Tx = True
    '        NAK_Count = 1
    '        TheEFTTicket.Clear()
    '        For i As Integer = 0 To 6
    '            Try
    '                If Tx Then
    '                    'UpdateState("Trasmissione conferma fisica", States.SETEFI_IN_PROGRESS)
    '                    Me.WriteToEFT(TxStates.TX_ACK)
    '                End If
    '                Tx = False
    '                'UpdateState("Ricezione scontrino", States.SETEFI_IN_PROGRESS)
    '                Me.ReadFromEFT(RxStates.RX_EFT_TICKET, 90)
    '                'UpdateState("Trasmissione conferma fisica", States.SETEFI_IN_PROGRESS)
    '                Me.WriteToEFT(TxStates.TX_ACK)
    '                TheEFTTicket.Add(m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE)
    '                Dim END_INDEX As Integer = m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE.IndexOf(Chr(&H1B))
    '                If END_INDEX = -1 Then
    '                    Continue For
    '                Else
    '                    'Exit For
    '                    ' The character before the &H1B must be &H7D if this is really the last receipt message
    '                    If AscW(m_TheSCONTRINO_DA_TERMINALE.LINEE_DA_STAMPARE(END_INDEX - 1)) = &H7D Then
    '                        ' OK, we collected all receipt messages
    '                        Exit For
    '                    Else
    '                        Continue For
    '                    End If
    '                End If
    '                Exit For
    '            Catch ex As EFTException
    '                If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
    '                    If i >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    WriteToEFT(TxStates.TX_NAK)
    '                    Tx = False
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
    '                    If NAK_Count >= 3 Then
    '                        Throw New EFTException(ex.Message)
    '                    End If
    '                    NAK_Count += 1
    '                    Tx = True
    '                    Continue For
    '                ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
    '                    Throw New EFTException(ex.Message)
    '                End If
    '            End Try
    '        Next

    '    Catch ex As EFTException
    '        'Console.WriteLine(ex.Message)
    '        ' If an exception occurs after the transaction is completed, we have to consider the transaction as executed anyway.
    '        ' In case of error getting the receipt linees we will inform the operator.
    '        If Not m_State = States.SETEFI_SUCCESS Then
    '            UpdateState(ex.Message, States.SETEFI_ERROR)
    '            setErrorMessage(ex.Message)
    '        End If
    '    Finally
    '        If (m_State = States.SETEFI_ERROR) And Me.IsConnected() And MustSendAbort Then
    '            UpdateState("Annullamento transazione", States.SETEFI_ERROR)
    '            WriteToEFT(TxStates.TX_NAK) ' 
    '            WriteToEFT(TxStates.TX_NAK) ' Send 3 nak to abort the current operation
    '            WriteToEFT(TxStates.TX_NAK) ' 
    '        End If
    '        UpdateState("Chiusura EFT", m_State)
    '        'Me.CloseDaSistemi()
    '        Me.m_TheModCntr.DialogActiv = False
    '    End Try

    'End Sub

    Public TheEFTTicket As New System.Collections.ArrayList

    Public Function GetTicket() As String

        GetTicket = ""
        Dim AllText As String = ""

        Try

            For Each Part As String In TheEFTTicket
                'GetTicket += GetPrintableLinees(Part)
                AllText += Part
            Next

            GetTicket = GetPrintableLineesNew(AllText)

        Catch ex As Exception

        End Try

    End Function

    Protected Overridable Function GetPrintableLineesNew(ByVal AllReceivedLinees As String) As String

        Dim GetPrintableLines As String = ""
        Dim i As Integer = 0
        Dim CurrentChar As Char = ""
        Dim LineLenght As Integer = 24 ' default line lenght
        Dim ETXFound As Boolean = False

        GetPrintableLineesNew = ""

        Try


            For i = 0 To AllReceivedLinees.Length - 1 Step 1

                CurrentChar = AllReceivedLinees(i)

                If ETXFound Then
                    ' this char is the LRC, skip it
                    GetPrintableLines += ""
                    ETXFound = False
                ElseIf CurrentChar = "^" Then
                    ' consider the next 42 chars
                    LineLenght = 42
                    If LineLenght = 42 Then
                        GetPrintableLines = GetPrintableLines.Trim
                    End If
                    If GetPrintableLines.Trim.StartsWith("- - - - -") Then
                        GetPrintableLines = "*                      *" & vbCrLf & _
                                            " ______________________ "
                    End If
                    GetPrintableLineesNew += GetPrintableLines + vbCrLf
                    GetPrintableLines = ""
                ElseIf CurrentChar = "}" Then
                    ' new line
                    If LineLenght = 42 Then
                        GetPrintableLines = GetPrintableLines.Trim
                    End If
                    If GetPrintableLines.Trim.StartsWith("- - - - -") Then
                        GetPrintableLines = "*                      *" & vbCrLf & _
                                            " ______________________ "
                    End If
                    GetPrintableLineesNew += GetPrintableLines + vbCrLf
                    GetPrintableLines = ""
                ElseIf CurrentChar = "~" Then
                    ' bold print will be ignored
                    GetPrintableLines += ""
                ElseIf AscW(CurrentChar) = ETX Then
                    ' ETX, skip this end the next one because it is the LRC
                    GetPrintableLines += ""
                    ETXFound = True
                ElseIf Char.IsLetterOrDigit(CurrentChar) OrElse Char.IsWhiteSpace(CurrentChar) OrElse Char.IsPunctuation(CurrentChar) Then
                    GetPrintableLines += CurrentChar
                End If
                If GetPrintableLines.Length > 0 AndAlso GetPrintableLines.Length Mod LineLenght = 0 Then
                    If LineLenght = 42 Then
                        GetPrintableLines = GetPrintableLines.Trim
                    End If
                    If GetPrintableLines.Trim.StartsWith("- - - - -") Then
                        GetPrintableLines = "*                      *" & vbCrLf & _
                                            " ______________________ "
                    End If
                    GetPrintableLineesNew += GetPrintableLines + vbCrLf
                    GetPrintableLines = ""
                    LineLenght = 24
                End If

            Next

        Catch ex As Exception

        Finally

        End Try

    End Function

    Protected Overridable Function GetPrintableLinees(ByVal ReceivedLinees As String) As String

        Dim FormattedLinees As String = ReceivedLinees
        GetPrintableLinees = ReceivedLinees
        LOG_FuncStart(getLocationString("GetPrintableLinees"))

        Try

            Dim ETX_INDEX As Integer = ReceivedLinees.IndexOf(Chr(ETX))
            If ETX_INDEX <> -1 Then
                FormattedLinees = ReceivedLinees.Substring(0, ETX_INDEX)
            End If

            For index As Integer = 0 To &H1F Step 1
                FormattedLinees = FormattedLinees.Replace(Chr(index), "")
            Next

            FormattedLinees = FormattedLinees.Replace(Chr(&H7E), "") _
                                                .Replace(Chr(&H7F), "") _
                                                    .Replace(Chr(&H1B), "") _
                                                        .Replace(Chr(&H1B), "")
            Dim i As Integer = 0
            If FormattedLinees.Length > 0 Then
                GetPrintableLinees = ""
                While i < FormattedLinees.Length

                    Dim s As String = FormattedLinees.Substring(i, Math.Min(24, FormattedLinees.Length - i))
                    Dim ii As Integer = s.IndexOf(Chr(&H7D))
                    If ii <> -1 Then
                        s = FormattedLinees.Substring(i, ii + 1)
                        i += s.Length
                        GetPrintableLinees += s.Replace(Chr(&H7D), vbCrLf)
                    Else
                        i += s.Length
                        GetPrintableLinees += s & vbCrLf
                    End If

                End While

            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetPrintableLinees"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetPrintableLinees"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("GetPrintableLinees"), "Function GetPrintableLinees returns" & GetPrintableLinees)
        End Try

    End Function


    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

    Public Sub AbortPayment()
        AbortPayment("")
    End Sub

    Public Sub AbortPayment(ByVal Reason As String)
        LOG_Error(getLocationString("AbortPayment"), "Abort the payment")
        If Not thread Is Nothing AndAlso thread.IsAlive Then
            AbortP = True
            setErrorMessage(Reason)
        End If
    End Sub
#End Region

#Region "Status/Message/Results"
    Public Message As String = ""
    Private Sub UpdateState(ByVal msg As String, ByVal status As Integer)
        Dim RealLen As Integer = -1
        Me.m_State = status

        Me.Message = ""
        For i As Integer = 0 To msg.Length - 1
            If Convert.ToInt32(msg(i)) <> 0 Then
                Me.Message += msg(i)
            Else
                Exit For
            End If
        Next
        Me.Message = Me.Message.Replace(Chr(SOH), "") _
                         .Replace(Chr(EOT), "")
        RaiseEvent DaSistemiStatusChanged(Me.m_State, Me.Message)
    End Sub

    Public Function GetStatus() As Integer
        Return Me.m_State
    End Function

    Public ErrorMessage As String = ""
    Protected Sub setErrorMessage(ByVal ErrorMessage As String)
        Me.ErrorMessage = ErrorMessage
    End Sub

    Private Sub CancelOperationIsAvailable(ByVal enable As Boolean)
        RaiseEvent DaSistemiCancelOperationIsAvailableChanged(enable)
    End Sub

#End Region

#Region "Interface function"
    Public Function PaySetefi(ByVal Amount As Double, ByRef TheModCntr As ModCntr, ByVal IsGift As Boolean, ByVal TheGiftPAN As String, ByVal IsVoid As Boolean, ByVal IsRCard As Boolean, ByVal TheRCardPAN As String) As Integer

        If Amount <= 0 AndAlso Not IsVoid Then
            UpdateState("Importo non corretto", States.SETEFI_ERROR)
            setErrorMessage("Importo non corretto")
            Return -1
        End If

        m_Amount = Amount
        m_TheModCntr = TheModCntr
        m_IsGift = IsGift
        m_TheGiftPAN = TheGiftPAN 'Format(TheGiftPAN, "00000000000")
        m_IsRCard = IsRCard
        m_TheRCardPAN = TheRCardPAN 'Format(TheRCardPAN, "0000")
        m_IsVoid = IsVoid
        m_DialogActiv = True
        AbortP = False
        ClearStructures()
        TheEFTTicket.Clear()
        setErrorMessage("")
        CancelOperationIsAvailable(True)

        UpdateState("Inizio procedura di pagamento", States.SETEFI_IN_PROGRESS)

        thread = New Thread(AddressOf Pay)
        thread.Start()

        Return 0

    End Function

    Public Function CloseSetefi(ByRef TheModCntr As ModCntr) As Integer

        m_TheModCntr = TheModCntr
        m_DialogActiv = True
        AbortP = False
        ClearStructures()
        TheEFTTicket.Clear()
        setErrorMessage("")
        CancelOperationIsAvailable(True)

        UpdateState("Inizio procedura di chiusura", States.SETEFI_IN_PROGRESS)

        thread = New Thread(AddressOf Close)
        thread.Start()

        Return 0

    End Function
#End Region

#Region "Exception"
    Class EFTException
        Inherits Exception

        Private Retry As Boolean
        Public Overridable ReadOnly Property ShouldBeRetried() As Boolean
            Get
                Return Retry
            End Get
        End Property

        Private Kind As KindOfException
        Public Overridable ReadOnly Property KindOfException() As KindOfException
            Get
                Return Kind
            End Get
        End Property

        Private Const EFTExceptionMsg As String = ""
        '"Errore EFT" & vbNewLine

        Public Sub New()
            MyBase.New(EFTExceptionMsg)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String)
            MyBase.New(String.Format("{0}{1}", _
                EFTExceptionMsg, auxMessage))
            Me.Retry = False
        End Sub ' New

        'Public Sub New(ByVal auxMessage As String, ByVal inner As Exception)
        '    MyBase.New(String.Format("{0}{1}", _
        '        DaSistemiException, auxMessage), inner)
        '    Me.Retry = False
        'End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal retry As Boolean)
            MyBase.New(String.Format("{0}{1}", _
                EFTExceptionMsg, auxMessage))
            Me.Retry = retry
        End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal kind As KindOfException)
            MyBase.New(String.Format("{0}{1}", _
                EFTExceptionMsg, auxMessage))
            Me.Kind = kind
        End Sub ' New

    End Class ' EFTException
#End Region

#Region "New"
    Public Sub New()
        ' set up default setting
        EFT = New SerialPort("COM1", 1200, IO.Ports.Parity.None, 8, IO.Ports.StopBits.One)
        ReadIniFile() ' read custom parameters
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        ' CloseDaSistemi()
        MyBase.Finalize()
    End Sub
#End Region

End Class


