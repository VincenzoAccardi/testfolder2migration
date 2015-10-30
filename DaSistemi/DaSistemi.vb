Imports System
Imports System.IO.Ports
Imports System.Threading
Imports Microsoft.VisualBasic
Imports System.Xml
Imports TPDotnet.Pos

Public Class DaSistemi

#Region "Documentation"
    ' ********** ********** ********** **********
    ' E F T - DaSistemi protocol implementation
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region

    Public Event DaSistemiStatusChanged(ByVal state As Integer, ByVal message As String)
    Public Event DaSistemiRemainingSecond(ByVal seconds As Integer)
    Private AbortP As Boolean = False
    Private thread As Thread = Nothing

    Enum States
        DA_SISTEMI_ERROR = -1
        DA_SISTEMI_SUCCESS = 0
        DA_SISTEMI_IN_PROGRESS = 1
    End Enum

    Enum RxStates
        RX_START_TRANSACTION_REQUEST
        RX_ACK
        RX_TRANSACTION_RESULT
    End Enum

    Enum TxStates
        TX_ACK
        TX_TRANSACTION_AMOUNT
        TX_NAK
    End Enum

    Enum KindOfException
        MSG_IS_INCORRECT ' send nak and read again
        NAK_RECEIVED     ' write and read again
        ABORT_OPERATION
        TRANSACTION_FAILED
    End Enum

    Private Shared EFT As SerialPort = Nothing
    Public RxState As RxStates
    Public TxState As TxStates
    Public State As States
    Dim Amount As Double
    Dim TheModCntr As ModCntr

    ReadOnly MAX_MSG_SIZE As Integer = 37
    ReadOnly STX As Byte = &H2
    ReadOnly ETX As Byte = &H3
    ReadOnly ACK() As Byte = {&H6, &H3, &H7A}
    ReadOnly NAK() As Byte = {&H15, &H3, &H69}

    Public Sub ListSerial()
        Dim ports As String() = SerialPort.GetPortNames()

        Console.WriteLine("The following serial ports were found:")

        Dim port As String
        For Each port In ports
            Console.WriteLine(port)
        Next port

    End Sub

    Private Function CloseDaSistemi() As Boolean

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
                Throw New DaSistemiException("Chiusura porta : " + ex.Message)
                If i = 3 Then
                    Return False
                End If
                Continue For
            End Try
        Next

        EFT = Nothing

        Return True

    End Function

    Public Sub ReadIniFile()

        Dim document As XmlDocument
        Dim modelList As XmlNodeList
        Dim modelNode As XmlNode

        Try
            document = New XmlDocument()
            document.Load(getPosConfigurationPath() + "\\Eft.xml")

            modelList = document.SelectNodes("/Eft/Setting")
            For Each modelNode In modelList

                Dim Protocol As String = modelNode.Attributes.GetNamedItem("Protocol").Value
                If Protocol.Equals("DaSistemi") Then

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

                        End If

                    Next

                End If

            Next
        Catch errorVariable As Exception
            Console.Write(errorVariable.ToString())
        End Try

    End Sub

    Private Function OpenDaSistemi() As Boolean

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
                Throw New DaSistemiException("Apertura porta: " + ex.Message)
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

    Public Shared Function ToHexString(ByVal bytes() As Byte, ByVal count As Integer) As String

        Dim hexStr As String = ""
        Dim i As Integer
        For i = 0 To count - 1
            hexStr = hexStr + "<" + Hex(bytes(i)) + ">"
        Next i
        Return hexStr

    End Function

    Public Function IsConnected() As Boolean

        'If EFT.IsOpen Then
        'Return (EFT.CDHolding Or EFT.CtsHolding Or EFT.DsrHolding)
        'Else
        'Return False
        'End If
        Return EFT.IsOpen

    End Function

    Private Function ReadFromDaSistemi(ByVal what As Integer, ByVal timeout As Integer) As Boolean

        Dim currentBytesNo As Integer
        Dim currentBytes(1024) As Byte
        Dim totalBytesNo As Integer
        Dim totalBytes(1024) As Byte
        Dim EtxAlreadyFound As Boolean
        Dim RetryNo As Integer = 1

        Me.RxState = what

        Console.WriteLine("readFromDaSistemi : {0}", RxState)

        If Not IsConnected() Then
            Throw New DaSistemiException("Lettura dati : EFT non presente", KindOfException.ABORT_OPERATION)
        End If

        Dim startDateTime As DateTime = DateTime.Now

        EFT.ReadTimeout = 1000
        EtxAlreadyFound = False
        While True

            Try
                System.Array.Clear(currentBytes, 0, currentBytes.Length - 1)
                currentBytesNo = EFT.Read(currentBytes, 0, currentBytes.Length - 1)
                ' Checking the received buffer
                If currentBytesNo < 0 Then
                    Throw New DaSistemiException("Lettura dati : Errore lettura dati dalla porta specificata!", KindOfException.ABORT_OPERATION)
                ElseIf ((totalBytesNo + currentBytesNo) > MAX_MSG_SIZE) Then
                    'Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(Too many long message)")
                    Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(Too many long message)", KindOfException.MSG_IS_INCORRECT)
                End If

                System.Array.ConstrainedCopy(currentBytes, 0, totalBytes, totalBytesNo, currentBytesNo)
                totalBytesNo += currentBytesNo
                Dim index As Integer = System.Array.IndexOf(currentBytes, ETX)
                If index = -1 And Not EtxAlreadyFound Then ' Did I receive the ETX ?
                    Continue While ' No, I didn't
                Else ' Yes, I did
                    If EtxAlreadyFound And Not currentBytesNo = 1 Then ' Check if the Etx was already received
                        Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(EXT already found and only the CRC (1 byte) must follow the ETX)", KindOfException.MSG_IS_INCORRECT)
                    End If
                    EtxAlreadyFound = True
                    If totalBytes(totalBytesNo - 2) = ETX Then
                        ' Entire message received, check the CRC
                        Console.WriteLine("Received message : " + ToHexString(totalBytes, totalBytesNo))
                        Console.WriteLine("Received message length : {0}", totalBytesNo)
                        Dim crc As Byte = Me.CalculateCRC(totalBytes, 0, totalBytesNo - 2)
                        If crc = totalBytes(totalBytesNo - 1) Then
                            Exit While
                        Else
                            Console.WriteLine("Received message : CRC received {0} and calculated {1} are differ!", Hex(totalBytes(totalBytesNo - 1)), Hex(crc))
                            Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(CRC)", KindOfException.MSG_IS_INCORRECT)
                        End If
                    ElseIf totalBytes(totalBytesNo - 1) = ETX Then
                        Continue While ' I received the ETX, now I need to receive the last byte (CRC)
                    Else
                        Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(Only the CRC (1 byte) must follow the ETX)", KindOfException.MSG_IS_INCORRECT)
                    End If

                End If
            Catch ex As TimeoutException
                Dim ts As TimeSpan = DateTime.Now.Subtract(startDateTime)
                If ts.Seconds < timeout Then
                    Console.Write("{0}", (timeout - ts.Seconds))
                    RaiseEvent DaSistemiRemainingSecond((timeout - ts.Seconds))
                Else
                    Throw New DaSistemiException("Lettura dati : Tempo scaduto!", KindOfException.ABORT_OPERATION)
                End If
                If AbortP Then
                    Throw New DaSistemiException("Lettura dati : Interrotto dall'operatore!", KindOfException.ABORT_OPERATION)
                End If
            Catch ex As Exception
                Throw New DaSistemiException("Lettura dati : " + ex.Message, KindOfException.ABORT_OPERATION)
            End Try

        End While

        If totalBytesNo = 0 Then
            Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(Read 0 bytes)", KindOfException.ABORT_OPERATION)
        End If

        If (String.Compare(Me.ByteArrayToStr(totalBytes), Me.ByteArrayToStr(Me.NAK)) = 0) Then
            Throw New DaSistemiException("Lettura dati : NAK!", KindOfException.NAK_RECEIVED)
        End If

        Select Case what

            Case RxStates.RX_ACK
                If Not (String.Compare(Me.ByteArrayToStr(totalBytes), Me.ByteArrayToStr(Me.ACK)) = 0) Then
                    Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(Expected ACK)", KindOfException.MSG_IS_INCORRECT)
                End If

            Case RxStates.RX_START_TRANSACTION_REQUEST
                If Not ((totalBytes(0) = STX) And (totalBytes(totalBytesNo - 2) = ETX)) Then
                    Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(STX-ETX)", KindOfException.MSG_IS_INCORRECT)
                ElseIf Not (totalBytes(10) = AscW("I")) Then
                    Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(MessageCode)", KindOfException.MSG_IS_INCORRECT)
                End If

            Case RxStates.RX_TRANSACTION_RESULT
                If Not ((totalBytes(0) = STX) And (totalBytes(totalBytesNo - 2) = ETX)) Then
                    Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(STX-ETX)", KindOfException.MSG_IS_INCORRECT)
                ElseIf Not (totalBytes(10) = AscW("E")) Then
                    Throw New DaSistemiException("Lettura dati : Messaggio non corretto!" & vbNewLine & "(MessageCode)", KindOfException.MSG_IS_INCORRECT)
                ElseIf Not ((totalBytes(11) = AscW("0")) And (totalBytes(12) = AscW("0"))) Then
                    Throw New DaSistemiException("Lettura dati : Transazione non eseguita!", KindOfException.TRANSACTION_FAILED)
                Else
                    Dim TerminalID As String
                    Dim Pan As String
                    Dim tmp(1024) As Byte
                    Array.ConstrainedCopy(totalBytes, 1, tmp, 0, 9)
                    TerminalID = Me.ByteArrayToStr(tmp)
                    Array.Clear(tmp, 0, 1024)
                    Array.ConstrainedCopy(totalBytes, 13, tmp, 0, 19)
                    Pan = Me.ByteArrayToStr(tmp)
                End If

        End Select

        Return True

    End Function

    Private Function CalculateCRC(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Byte

        Dim i As Integer
        Dim crc As Byte = &H7F
        For i = offset To count
            crc = (crc Xor buffer(i))
        Next i

        Console.WriteLine("CRC = " + Hex(crc))
        Return crc

    End Function

    Private Function BuildTxACKMessage(ByRef b() As Byte) As Integer

        System.Array.Copy(Me.ACK, b, Me.ACK.Length)
        Return Me.ACK.Length

    End Function

    Private Function BuildTxNAKMessage(ByRef b() As Byte) As Integer

        System.Array.Copy(Me.NAK, b, Me.NAK.Length)
        Return Me.NAK.Length

    End Function


    Private Function TxACK() As Boolean

        Try
            EFT.Write(ACK, 0, ACK.Length)
        Catch ex As Exception
            Throw New DaSistemiException("Trasmissione ACK : " + ex.Message)
            Return False
        End Try

        Return True

    End Function

    Private Function TxNAK() As Boolean

        Try
            EFT.Write(NAK, 0, NAK.Length)
        Catch ex As Exception
            Throw New DaSistemiException("Trasmissione NAK : " + ex.Message)
            Return False
        End Try

        Return True

    End Function

    Private Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.ASCIIEncoding()
        Return encoding.GetBytes(str)
    End Function

    Private Function ByteArrayToStr(ByVal bytes() As Byte) As String
        Dim encoding As New System.Text.ASCIIEncoding()
        Return encoding.GetString(bytes)
    End Function

    Private Function BuildTxTransactionAmountMessage(ByRef b() As Byte) As Integer

        Dim s As String = ""
        ' 04-12-08 : Amount calculation : use Format() instead of Replace()
        's = String.Format("{0}000000000I{1:D8}0000{2:D32}{3}", Microsoft.VisualBasic.Chr(STX), Amount.ToString().Replace(",", "").PadLeft(8, "0"), 0, Microsoft.VisualBasic.Chr(ETX))
        s = String.Format("{0}000000000I{1:D8}0000{2:D32}{3}", Microsoft.VisualBasic.Chr(STX), Format(Amount * 100, "0").PadLeft(8, "0"), 0, Microsoft.VisualBasic.Chr(ETX))
        Dim crc As Byte = Me.CalculateCRC(Me.StrToByteArray(s), 0, s.Length - 1)
        s = String.Format("{0}{1}", s, Microsoft.VisualBasic.Chr(crc))
        System.Array.Copy(Me.StrToByteArray(s), b, s.Length)

        Return s.Length

    End Function

    Private Function WriteToDaSistemi(ByVal what As Integer) As Boolean

        Dim txBuffer(1024) As Byte
        Dim count As Integer = 0
        TxState = what
        Console.WriteLine("writeToDaSistemi : {0}", TxState)

        If Not IsConnected() Then
            Throw New DaSistemiException("Scrittura dati : EFT non presente", KindOfException.ABORT_OPERATION)
            Return False
        End If

        EFT.WriteTimeout = 3000
        System.Array.Clear(txBuffer, 0, txBuffer.Length - 1)

        Select Case TxState

            Case TxStates.TX_ACK
                count = BuildTxACKMessage(txBuffer)

            Case TxStates.TX_TRANSACTION_AMOUNT
                count = BuildTxTransactionAmountMessage(txBuffer)

            Case TxStates.TX_NAK
                count = BuildTxNAKMessage(txBuffer)

        End Select

        Try
            EFT.Write(txBuffer, 0, count)
            Console.WriteLine("Sent message : " + ToHexString(txBuffer, count))
            Console.WriteLine("Sent message length : {0}", count)
        Catch ex As Exception
            Throw New DaSistemiException(" Trasmissione importo : " + ex.Message, KindOfException.ABORT_OPERATION)
            Return False
        End Try

        Return True

    End Function

    Public Message As String = ""
    Private Sub UpdateState(ByVal msg As String, ByVal status As Integer)
        Me.State = status
        Me.Message = msg
        RaiseEvent DaSistemiStatusChanged(Me.State, Me.Message)
    End Sub

    Public Function GetStatus() As Integer
        Return Me.State
    End Function

    Public ErrorMessage As String = ""
    Public Sub setErrorMessage(ByVal ErrorMessage As String)
        Me.ErrorMessage = ErrorMessage
    End Sub

    Public Sub Pay()
        AbortP = False
        Dim MustSendAbort As Boolean = True
        Try
            UpdateState("Apertura EFT", States.DA_SISTEMI_IN_PROGRESS)
            Me.OpenDaSistemi()

            For i As Integer = 1 To 3

                Try
                    UpdateState("Ricezione richiesta importo", States.DA_SISTEMI_IN_PROGRESS)
                    Me.ReadFromDaSistemi(RxStates.RX_START_TRANSACTION_REQUEST, 30)
                    UpdateState("Trasmissione conferma fisica", States.DA_SISTEMI_IN_PROGRESS)
                    Me.WriteToDaSistemi(TxStates.TX_ACK)
                    Exit For
                Catch ex As DaSistemiException
                    If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Or _
                       ex.KindOfException = KindOfException.NAK_RECEIVED Then ' NAK not allowed here, handle it as an incorrect message
                        If i >= 3 Then
                            Throw New DaSistemiException(ex.Message)
                        End If
                        TxNAK()
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
                        Throw New DaSistemiException(ex.Message)
                    End If
                End Try

            Next

            Dim Tx As Boolean = True
            Dim NAK_Count As Integer = 1
            For i As Integer = 0 To 3
                Try
                    If Tx Then
                        UpdateState("Trasmissione importo", States.DA_SISTEMI_IN_PROGRESS)
                        Me.WriteToDaSistemi(TxStates.TX_TRANSACTION_AMOUNT)
                    End If
                    UpdateState("Ricezione conferma fisica", States.DA_SISTEMI_IN_PROGRESS)
                    Me.ReadFromDaSistemi(RxStates.RX_ACK, 15)
                    Exit For
                Catch ex As DaSistemiException
                    If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Then
                        If i >= 3 Then
                            Throw New DaSistemiException(ex.Message)
                        End If
                        TxNAK()
                        Tx = False
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.NAK_RECEIVED Then
                        If NAK_Count >= 3 Then
                            Throw New DaSistemiException(ex.Message)
                        End If
                        NAK_Count += 1
                        Tx = True
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
                        Throw New DaSistemiException(ex.Message)
                    End If
                End Try
            Next

            For i As Integer = 0 To 3
                Try
                    UpdateState("Ricezione esito transazione", States.DA_SISTEMI_IN_PROGRESS)
                    Me.ReadFromDaSistemi(RxStates.RX_TRANSACTION_RESULT, 60)
                    Exit For
                Catch ex As DaSistemiException
                    If ex.KindOfException = KindOfException.MSG_IS_INCORRECT Or _
                       ex.KindOfException = KindOfException.NAK_RECEIVED Then ' NAK not allowed here, handle it as an incorrect message
                        If i >= 3 Then
                            Throw New DaSistemiException(ex.Message)
                        End If
                        TxNAK()
                        Continue For
                    ElseIf ex.KindOfException = KindOfException.ABORT_OPERATION Then
                        Throw New DaSistemiException(ex.Message)
                    ElseIf ex.KindOfException = KindOfException.TRANSACTION_FAILED Then
                        Me.WriteToDaSistemi(TxStates.TX_ACK)
                        MustSendAbort = False
                        Throw New DaSistemiException(ex.Message)
                    End If

                End Try
            Next

            UpdateState("Trasmissione conferma fisica", States.DA_SISTEMI_IN_PROGRESS)
            Me.WriteToDaSistemi(TxStates.TX_ACK)
            UpdateState("Transazione eseguita", States.DA_SISTEMI_SUCCESS)

        Catch ex As DaSistemiException
            Console.WriteLine(ex.Message)
            UpdateState(ex.Message, States.DA_SISTEMI_ERROR)
            setErrorMessage(ex.Message)
        Finally
            If (State = States.DA_SISTEMI_ERROR) And Me.IsConnected() And MustSendAbort Then
                UpdateState("Annullamento transazione", States.DA_SISTEMI_ERROR)
                WriteToDaSistemi(TxStates.TX_NAK) ' 
                WriteToDaSistemi(TxStates.TX_NAK) ' Send 3 nak to abort the current operation
                WriteToDaSistemi(TxStates.TX_NAK) ' 
            End If
            UpdateState("Chiusura EFT", State)
            'Me.CloseDaSistemi()
            Me.TheModCntr.DialogActiv = False
        End Try

    End Sub

    Public Sub AbortPayment()
        If thread.IsAlive Then
            AbortP = True
        End If
    End Sub

    Public Function PayDaSistemi(ByVal Amount As Double, ByRef TheModCntr As ModCntr) As Integer

        If Amount <= 0 Then
            UpdateState("Importo non corretto", States.DA_SISTEMI_ERROR)
            setErrorMessage("Importo non corretto")
            Return -1
        End If

        Me.Amount = Amount
        Me.TheModCntr = TheModCntr
        Me.AbortP = False

        UpdateState("Inizio procedura di pagamento", States.DA_SISTEMI_IN_PROGRESS)

        thread = New Thread(AddressOf Pay)
        thread.Start()

        Return 0

    End Function

    Class DaSistemiException
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

        Private Const DaSistemiException As String = _
            "Errore DaSistemi" & vbNewLine

        Public Sub New()
            MyBase.New(DaSistemiException)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String)
            MyBase.New(String.Format("{0}{1}", _
                DaSistemiException, auxMessage))
            Me.Retry = False
        End Sub ' New

        'Public Sub New(ByVal auxMessage As String, ByVal inner As Exception)
        '    MyBase.New(String.Format("{0}{1}", _
        '        DaSistemiException, auxMessage), inner)
        '    Me.Retry = False
        'End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal retry As Boolean)
            MyBase.New(String.Format("{0}{1}", _
                DaSistemiException, auxMessage))
            Me.Retry = retry
        End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal kind As KindOfException)
            MyBase.New(String.Format("{0}{1}", _
                DaSistemiException, auxMessage))
            Me.Kind = kind
        End Sub ' New

    End Class ' DaSistemiException

    Public Sub New()
        ' set up default setting
        EFT = New SerialPort("COM1", 1200, IO.Ports.Parity.None, 8, IO.Ports.StopBits.One)
        ReadIniFile() ' read custom parameters
    End Sub

    Protected Overrides Sub Finalize()
        ' CloseDaSistemi()
        MyBase.Finalize()
    End Sub
End Class


