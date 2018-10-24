Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet
Imports TPDotnet.Pos
Imports System.Collections
Imports System.Xml.XPath
Imports System.Linq

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

#Region "IO Parameters"

#Region "GiftCard"

Public Class GiftCardActivationParameters
    Inherits GiftCardCommonParametersRecord

    Private _ArticleRecord As TPDotnet.Pos.TaArtSaleRec
    Public Property ArticleRecord() As TPDotnet.Pos.TaArtSaleRec
        Get
            Return _ArticleRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaArtSaleRec)
            _Record = value
            _ArticleRecord = value
        End Set
    End Property

    Public Overrides Property Value As Decimal
        Get
            If Not _ArticleRecord Is Nothing Then
                _Value = _ArticleRecord.dTaTotal
            End If
            Return _Value
        End Get
        Set(value As Decimal)
            _Value = value
        End Set
    End Property

    Public Overridable ReadOnly Property GiftCardBarcodeField As String
        Get
            Return "szITArgenteaGiftCardEAN"
        End Get
    End Property

    Public Overrides Property Barcode As String
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(GiftCardBarcodeField) Then
                Return _Record.GetPropertybyName(GiftCardBarcodeField).ToString()
            End If
            Return MyBase.Barcode
        End Get
        Set(value As String)
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(GiftCardBarcodeField) Then
                    _Record.AddField(GiftCardBarcodeField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(GiftCardBarcodeField, value.ToString)
            End If
            MyBase.Barcode = value
        End Set
    End Property

End Class

Public Class GiftCardRedeemParameters
    Inherits GiftCardCommonParametersRecord

    Protected _MediaRecord As TPDotnet.Pos.TaMediaRec
    Public Property MediaRecord() As TPDotnet.Pos.TaMediaRec
        Get
            Return _MediaRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaMediaRec)
            _Record = value
            _MediaRecord = value
        End Set
    End Property

    Public Overrides Property Value As Decimal
        Get
            If Not _MediaRecord Is Nothing Then
                _Value = _MediaRecord.dTaPaidTotal
            End If
            Return _Value
        End Get
        Set(value As Decimal)
            _Value = value
        End Set
    End Property

    Public Overrides Property Barcode() As String
        Get
            If Not _MediaRecord Is Nothing Then
                _Barcode = _MediaRecord.szBarcode
            End If
            Return _Barcode
        End Get
        Set(value As String)
            ' pay attention: Argentea overwrite the barcode with the transaction id
            ' with the following test, we avoid this behaviour
            If String.IsNullOrEmpty(_MediaRecord.szBarcode) Then
                _MediaRecord.szBarcode = value
            End If
        End Set
    End Property

End Class

Public Class GiftCardCommonParametersRecord
    Inherits CommonParametersRecord

    Public Overrides ReadOnly Property TransactionIdField As String
        Get
            Return "szArgenteaGiftCardTxID"
        End Get
    End Property

    Public Overrides ReadOnly Property StatusField As String
        Get
            Return "szArgenteaGiftCardStatus"
        End Get
    End Property

End Class

Public Class BalanceParameters
    Inherits CommonParameters

    Protected _GiftCardBalanceLineIdentifier As String = String.Empty
    Public Property GiftCardBalanceLineIdentifier() As String
        Get
            Return _GiftCardBalanceLineIdentifier
        End Get
        Set(value As String)
            _GiftCardBalanceLineIdentifier = value
        End Set
    End Property

    Public Overrides Property Value() As Decimal
        Get
            If Not String.IsNullOrEmpty(MessageOut) Then
                Dim linees As String() = MessageOut.Split(vbCrLf.ToCharArray)
                For Each l As String In linees
                    If l.ToUpper.StartsWith(_GiftCardBalanceLineIdentifier.ToUpper) Then
                        _Value = Convert.ToDecimal(l.ToUpper.Substring(_GiftCardBalanceLineIdentifier.Length).Trim)
                        Exit For
                    End If
                Next l
            End If
            Return _Value
        End Get
        Set(value As Decimal)
            _Value = value
        End Set
    End Property

    Public Overridable ReadOnly Property Receipt() As String
        Get
            Return _RefTo_MessageOut 'use message out as receipt
        End Get
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

#End Region

#Region "PhoneRecharge"

Public Class PhoneRechargeActivationParameters
    Inherits PhoneRechargeCommonParametersRecord

    Private _ArticleRecord As TPDotnet.Pos.TaArtSaleRec
    Public Property ArticleRecord() As TPDotnet.Pos.TaArtSaleRec
        Get
            Return _ArticleRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaArtSaleRec)
            _Record = value
            _ArticleRecord = value
        End Set
    End Property

    Public Overrides Property Value As Decimal
        Get
            If Not _ArticleRecord Is Nothing Then
                _Value = _ArticleRecord.dTaTotal
            End If
            Return _Value
        End Get
        Set(value As Decimal)
            _Value = value
        End Set
    End Property

    Public Overrides Property Barcode As String
        Get
            If Not _ArticleRecord Is Nothing Then
                Return _ArticleRecord.szInputString
            End If
            Return MyBase.Barcode
        End Get
        Set(value As String)
            _ArticleRecord.szInputString = value
        End Set
    End Property

    Public Const PinCounterObjectIdentifier As String = "PinCounter" + TPDotnet.IT.Common.Pos.PhoneRechargeItem

    Protected _GetNextPINCounter As Integer = -1
    Public Overridable ReadOnly Property GetNextPINCounter() As Integer
        Get
            If Not Controller.ObjectCash.ContainsKey(PinCounterObjectIdentifier) Then Controller.ObjectCash.Add(PinCounterObjectIdentifier, Transaction.lactTaNmbr.ToString + ";1")

            Dim content As String() = Controller.ObjectCash(PinCounterObjectIdentifier).Split(";".ToCharArray)
            _GetNextPINCounter = IIf(Transaction.lactTaNmbr = CInt(content(0)), CInt(content(1)) + 1, 1)
            Controller.ObjectCash(PinCounterObjectIdentifier) = Transaction.lactTaNmbr.ToString + ";" + _GetNextPINCounter.ToString

            Return _GetNextPINCounter
        End Get
    End Property

    Public Overridable ReadOnly Property PhoneRechargePinCounterField As String
        Get
            Return "lITArgenteaPhoneRechargePinCounter"
        End Get
    End Property

    Private _PINCounter As Integer
    Public Property PINCounter() As Integer
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(PhoneRechargePinCounterField) Then
                Return _Record.GetPropertybyName(PhoneRechargePinCounterField).ToString()
            End If
            Return _PINCounter
        End Get
        Set(value As Integer)
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(PhoneRechargePinCounterField) Then
                    _Record.AddField(PhoneRechargePinCounterField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(PhoneRechargePinCounterField, value.ToString)
            End If
            _PINCounter = value
        End Set
    End Property

    Public Overridable ReadOnly Property PhoneRechargePinIdField As String
        Get
            Return "szITArgenteaPhoneRechargePinID"
        End Get
    End Property

    Private _PinID As String = String.Empty
    Public Property PinID() As String
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(PhoneRechargePinIdField) Then
                Return _Record.GetPropertybyName(PhoneRechargePinIdField).ToString()
            End If
            Return _PinID
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then value = ""
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(PhoneRechargePinIdField) Then
                    _Record.AddField(PhoneRechargePinIdField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(PhoneRechargePinIdField, value.ToString)
            End If
            _PinID = value
        End Set
    End Property

End Class

Public Class PhoneRechargeCommonParametersRecord
    Inherits CommonParametersRecord

    Public Overrides ReadOnly Property TransactionIdField As String
        Get
            Return "szArgenteaPhoneRechargeTxID"
        End Get
    End Property

    Public Overrides ReadOnly Property StatusField As String
        Get
            Return "szArgenteaPhoneRechargeStatus"
        End Get
    End Property

End Class

#End Region

#Region "External GiftCard"

Public Class ExternalGiftCardActivationParameters
    Inherits ExternalGiftCardCommonParametersRecord

    Private _ArticleRecord As TPDotnet.Pos.TaArtSaleRec
    Public Property ArticleRecord() As TPDotnet.Pos.TaArtSaleRec
        Get
            Return _ArticleRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaArtSaleRec)
            _Record = value
            _ArticleRecord = value
        End Set
    End Property

    Public Overrides Property Value As Decimal
        Get
            If Not _ArticleRecord Is Nothing Then
                _Value = _ArticleRecord.dTaTotal
            End If
            Return _Value
        End Get
        Set(value As Decimal)
            _Value = value
        End Set
    End Property

    Public Overrides Property Barcode As String
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(ExternalGiftCardBarcodeField) Then
                Return _Record.GetPropertybyName(ExternalGiftCardBarcodeField).ToString()
            End If
            Return MyBase.Barcode
        End Get
        Set(value As String)
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(ExternalGiftCardBarcodeField) Then
                    _Record.AddField(ExternalGiftCardBarcodeField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(ExternalGiftCardBarcodeField, value.ToString)
            End If
            MyBase.Barcode = value
        End Set
    End Property

    Public Overridable ReadOnly Property ExternalGiftCardBarcodeField As String
        Get
            Return "szITArgenteaExternalGiftCardEAN"
        End Get
    End Property

End Class

Public Class ExternalGiftCardDeActivationParameters
    Inherits ExternalGiftCardCommonParametersRecord

    Private _ArticleRecord As TPDotnet.Pos.TaArtSaleRec
    Public Property ArticleRecord() As TPDotnet.Pos.TaArtSaleRec
        Get
            Return _ArticleRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaArtSaleRec)
            _Record = value
            _ArticleRecord = value
        End Set
    End Property

    Private _ArticleReturnRecord As TPDotnet.Pos.TaArtReturnRec
    Public Property ArticleReturnRecord() As TPDotnet.Pos.TaArtReturnRec
        Get
            Return _ArticleReturnRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaArtReturnRec)
            _Record = value
            _ArticleReturnRecord = value
        End Set
    End Property

    Public Overrides Property Value As Decimal
        Get
            If Not _ArticleRecord Is Nothing Then
                _Value = _ArticleRecord.dTaTotal
            ElseIf Not _ArticleReturnRecord Is Nothing Then
                _Value = _ArticleReturnRecord.dTaTotal
            End If
            Return _Value
        End Get
        Set(value As Decimal)
            _Value = value
        End Set
    End Property

    Public Overrides Property Barcode As String
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(ExternalGiftCardBarcodeField) Then
                Return _Record.GetPropertybyName(ExternalGiftCardBarcodeField).ToString()
            End If
            Return MyBase.Barcode
        End Get
        Set(value As String)
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(ExternalGiftCardBarcodeField) Then
                    _Record.AddField(ExternalGiftCardBarcodeField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(ExternalGiftCardBarcodeField, value.ToString)
            End If
            MyBase.Barcode = value
        End Set
    End Property

    Public Overridable ReadOnly Property ExternalGiftCardBarcodeField As String
        Get
            Return "szITArgenteaExternalGiftCardEAN"
        End Get
    End Property

End Class

Public Class ExternalGiftCardConfirmParameters
    Inherits ExternalGiftCardCommonParametersRecord

End Class

Public Class ExternalGiftCardCommonParametersRecord
    Inherits CommonParametersRecord

    Public Overrides ReadOnly Property TransactionIdField As String
        Get
            Return "szArgenteaExternalGiftCardTxID"
        End Get
    End Property

    Public Overrides ReadOnly Property StatusField As String
        Get
            Return "szArgenteaExternalGiftCardStatus"
        End Get
    End Property

End Class

#End Region

#Region "BP Parameters"

''' <summary>
'''     Parametri Common e  specializzati
'''     per l'uso in controller e service
'''     dedicati alla gestione dei BP sia
'''     Cartacei che Elettronici.
''' </summary>
Public Class BPParameters
    Inherits BPCommonParametersRecord

#Region "Variabili Private"

    Protected _MediaRecord As TPDotnet.Pos.TaMediaRec
    Protected _MediaMemberMediaRecord As TPDotnet.Pos.TaMediaMemberDetailRec

    Protected m_TerminalID As String = ""

#End Region


#Region "Property relative ad Argentea per definizione in Parametri"

    ''' <summary>
    '''     Return a Code of Terminal rescued
    '''     after call funztion remote Argentea
    ''' </summary>
    ''' <returns>String</returns>
    Public Property TerminalID() As String
        Get
            Return m_TerminalID
        End Get
        Set(ByVal value As String)
            m_TerminalID = value
        End Set
    End Property

#End Region


    ''' <summary>
    '''     Argomento specializzato ad avere la "TA Media Record"
    ''' </summary>
    ''' <returns>TPDotnet.Pos.TaMediaRec</returns>
    Public Property MediaRecord() As TPDotnet.Pos.TaMediaRec
        Get
            Return _MediaRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaMediaRec)
            _Record = value
            _MediaRecord = value
        End Set
    End Property

    ''' <summary>
    '''     Argomento specializzato ad avere la "TA Media Detail Record"
    ''' </summary>
    ''' <returns>TPDotnet.Pos.TaMediaMemberDetailRec</returns>
    Public Property MediaMemberDetailRecord() As TPDotnet.Pos.TaMediaMemberDetailRec
        Get
            Return _MediaMemberMediaRecord
        End Get
        Set(ByVal value As TPDotnet.Pos.TaMediaMemberDetailRec)
            _Record = value
            _MediaMemberMediaRecord = value
        End Set
    End Property


End Class

''' <summary>
'''     Strato per gli args (dinamici) di riferimento alla function Handler
''' </summary>
Public Class BPCommonParametersRecord
    Inherits CommonParametersRecord

    ''' <summary>
    '''     ID della Transazione (Comune)
    ''' </summary>
    ''' <returns></returns>
    Public Overrides ReadOnly Property TransactionIdField As String
        Get
            Return "szBPCTxID"
        End Get
    End Property

    ''' <summary>
    '''     Stato della Transazione
    ''' </summary>
    ''' <returns></returns>
    Public Overrides ReadOnly Property StatusField As String
        Get
            Return "szBPCStatus"
        End Get
    End Property

End Class


#End Region

#Region "Common"

Public Class CommonParametersRecord
    Inherits CommonParameters

    Protected _Record As TPDotnet.Pos.TaBaseRec
    Protected Property Record() As TPDotnet.Pos.TaBaseRec
        Get
            Return _Record
        End Get
        Set(ByVal value As TPDotnet.Pos.TaBaseRec)
            _Record = value
        End Set
    End Property

    Public Overridable ReadOnly Property TransactionIdField As String
        Get
            Return "TransactionIdField"
        End Get
    End Property

    Public Overridable ReadOnly Property StatusField As String
        Get
            Return "StatusField"
        End Get
    End Property

    Public Overrides Property TransactionID As String
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(TransactionIdField) Then
                Return _Record.GetPropertybyName(TransactionIdField).ToString()
            End If
            Return MyBase.TransactionID
        End Get
        Set(value As String)
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(TransactionIdField) Then
                    _Record.AddField(TransactionIdField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(TransactionIdField, value.ToString)
            End If
            MyBase.TransactionID = value
        End Set
    End Property

    Public Overrides Property Status As String
        Get
            If Not _Record Is Nothing AndAlso _Record.ExistField(StatusField) Then
                Return _Record.GetPropertybyName(StatusField).ToString()
            End If
            Return MyBase.Status
        End Get
        Set(value As String)
            If Not _Record Is Nothing Then
                If Not _Record.ExistField(StatusField) Then
                    _Record.AddField(StatusField, DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                End If
                _Record.setPropertybyName(StatusField, value.ToString)
            End If
            MyBase.Status = value
        End Set
    End Property

End Class


''' <summary>
'''     Parametri comuni in Uso da dll shared
''' </summary>
Public Class CommonParameters

    Friend _ParametersBase As System.Collections.Generic.Dictionary(Of String, Object)

#Region "Passati per riferimento per la formulazione"

    Protected _Controller As TPDotnet.Pos.ModCntr
    Protected _Transaction As TPDotnet.Pos.TA

    ''' <summary>
    '''     Il Controller passato come argomento dal chiamante con il metodo
    '''     reflection usato come argomento dinamico Common a tutti.
    ''' </summary>
    ''' <returns>Tipo atteso.: ModCntr <see cref="TPDotnet.Pos.ModCntr"/> </returns>
    Public Overridable Property Controller() As TPDotnet.Pos.ModCntr
        Get
            Return _Controller
        End Get
        Set(ByVal value As TPDotnet.Pos.ModCntr)
            _Controller = value
        End Set
    End Property

    ''' <summary>
    '''     LA TA passata come argomento dal chiamante con il metodo
    '''     reflection usata come argomento dinamico Common a tutti.
    ''' </summary>
    ''' <returns>Tipo atteso.: TA <see cref="TPDotnet.Pos.TA"/> </returns>
    Public Overridable Property Transaction() As TPDotnet.Pos.TA
        Get
            Return _Transaction
        End Get
        Set(ByVal value As TPDotnet.Pos.TA)
            _Transaction = value
        End Set
    End Property

#End Region

#Region "Interpretati e fillati per l'esito"

    Protected _Successfull As Boolean = False
    Protected _ErrorMessage As String = String.Empty
    Protected _SuccessMessage As String = String.Empty
    'Protected _CommittRequired As Boolean = False

    ''' <summary>
    '''     Usato subito dopo l'interpretazione della risposta
    '''     alla chiamata verso Argentea per l'esito OK o KO
    '''     basatado dalla risposta sul protocollo
    ''' </summary>
    Public Overridable Property Successfull() As Boolean
        Get
            Return _Successfull
        End Get
        Set(ByVal value As Boolean)
            _Successfull = value
        End Set
    End Property

    ''' <summary>
    '''     Usato subito dopo l'interpretazione della risposta
    '''     alla chiamata verso Argentea per l'esito OK o KO
    '''     definisce se il Successfull ha bisogno di una 
    '''     ulteriore chiamata di conferma per andare avanti
    ''' </summary>
    'Public Overridable Property CommittRequired() As Boolean
    'Get
    'Return _CommittRequired
    'End Get
    'Set(ByVal value As Boolean)
    '       _CommittRequired = value
    'End Set
    'End Property

    ''' <summary>
    '''     Usato per intermezzo di segnalazione errori
    '''     fillato dalla funzione Hundler e non dalla codifica
    '''     della risposta sul protocollo
    ''' </summary>
    ''' <returns>Stringa con errore impostato internamente</returns>
    Public Overridable Property ErrorMessage() As String
        Get
            Return _ErrorMessage
        End Get
        Set(ByVal value As String)
            _ErrorMessage = value
        End Set
    End Property

    ''' <summary>
    '''     Usato per intermezzo di segnalazione errori
    '''     fillato dalla funzione Hundler e non dalla codifica
    '''     della risposta sul protocollo
    ''' </summary>
    ''' <returns>Stringa con errore impostato internamente</returns>
    Public Overridable Property SuccessMessage() As String
        Get
            Return _SuccessMessage
        End Get
        Set(ByVal value As String)
            _SuccessMessage = value
        End Set
    End Property


#End Region

#Region "Comuni"

    Protected _Status As String = String.Empty
    Protected _Barcode As String = String.Empty
    Protected _TransactionID As String = String.Empty
    Protected _Value As Decimal = 0
    Protected _ValueExcedeed As Decimal = 0
    Protected _UndoBPCForExcedeed As Boolean = False

    ''' <summary>
    '''     Status di operazione verso un azione Argentea
    ''' </summary>
    ''' <returns>Una stringa con stati di Attivazione/Disattivazione o in genere per identificare lo stato dell'operazione</returns>
    Public Overridable Property Status() As String
        Get
            Return _Status
        End Get
        Set(ByVal value As String)
            _Status = value
        End Set
    End Property

    ''' <summary>
    '''     Il Barcode valorizzato dalla funzione
    '''     di ingresso Dematarialize o Undo Dematiralize
    ''' </summary>
    ''' <returns>L'EAN del barcode scandito dall'opratore</returns>
    Public Overridable Property Barcode() As String
        Get
            _Barcode = _Barcode
            Return _Barcode
        End Get
        Set(value As String)
            _Barcode = value
        End Set
    End Property

    ''' <summary>
    '''     L'ID della Transazione ripreso dalla Riposta Argentea
    '''     corrispondente alla transazione remota codificata.
    ''' </summary>
    ''' <returns>Stringa con id della transazione remota</returns>
    Public Overridable Property TransactionID() As String
        Get
            Return _TransactionID
        End Get
        Set(ByVal value As String)
            _TransactionID = value
        End Set
    End Property

    ''' <summary>
    '''     Il Valore del Buono Ticket validato
    '''     dalla risposta remota da Argentea.
    '''     (Dietro Parsing dopo la chiamata)
    ''' </summary>
    ''' <returns>Decimal in Valuta €.cc</returns>
    Public Overridable Property Value() As Decimal
        Get
            Return _Value
        End Get
        Set(ByVal value As Decimal)
            _Value = value
        End Set
    End Property

    ''' <summary>
    '''     Il Valore del Buono Ticket eventuale
    '''     in eccesso rispetto al totale dal pagare.
    '''     (Dietro Parsing dopo la chiamata)
    ''' </summary>
    ''' <returns>Decimal in Valuta €.cc</returns>
    Public Overridable Property ValueExcedeed() As Decimal
        Get
            Return _ValueExcedeed
        End Get
        Set(ByVal value As Decimal)
            _ValueExcedeed = value
        End Set
    End Property

    ''' <summary>
    '''     Di supporto alla Opzione di Operatore
    '''     per l'eccesso su totale rispetto al Buono.
    '''     (Permette di rimuovere l'utlima chiamata in relazione all'eccesso 
    '''     del totale buono rispetto al pagabile senza sollverare un eccezione)
    ''' </summary>
    ''' <returns>Boolea flag per non sollevare eccezioni in entrata su azione RemoveHandler richiamato internamente</returns>
    Protected Friend Property UndoBPCForExcedeed() As Boolean
        Get
            Return _UndoBPCForExcedeed
        End Get
        Set(ByVal value As Boolean)
            _UndoBPCForExcedeed = value
        End Set
    End Property

    ''' <summary>
    '''     Restituisce dalla TA il numero di Buoni Pasto già che
    '''     sono stati utilizzati nella vendita corrente.
    ''' </summary>
    ''' <returns>Il numero di Buoni Pasto già utilizzati in corso della vendita.</returns>
    Protected Friend Function GetAlreadyOnTAScanned() As Integer

        Try
            Dim doc As Xml.Linq.XDocument = _Transaction.TAtoXDocument(False, 0, False)
            Return doc.XPathSelectElements("/TAS/NEW_TA/MEDIA[szBPC or szBPE]").Count()
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Public Overridable ReadOnly Property IntValue() As Integer
        Get
            Return Math.Abs(CInt(Value * 100))
        End Get
    End Property

#End Region

#Region "Di supporto alle chiamate per il ritorno in risposta"

    Protected _RefTo_MessageOut As String = String.Empty

    ''' <summary>
    '''     Per compatibilità altrimenti usare <see cref="RefTo_MessageOut"/>
    ''' </summary>
    Public Overridable Property MessageOut() As String
        Get
            Return _RefTo_MessageOut
        End Get
        Set(ByVal value As String)
            _RefTo_MessageOut = value
        End Set
    End Property

#End Region

#Region "Funzione principe per il caricamento Dinamico dei Parametri tra funzioni"

    ''' <summary>
    '''     Usato per il Load dei parametri passati alle funzioni Handler
    '''     con la dictionary dei nomi degli argomenti e in richiamto al
    '''     loro valore tramite Reflection. 
    '''     (Funzione chiamata subito all'ingresso delle funzioni Hunlder)
    '''     Finalità.: Proiettare gli argomenti dinamicmente rispetto al chiamante.
    ''' </summary>
    ''' <param name="source">Dictionary dei nomi degli argomenti con value del Valore sull'argomento valorizzato dal chiamante</param>
    Public Sub LoadCommonFunctionParameter(ByRef source As System.Collections.Generic.Dictionary(Of String, Object))
        Try
            _ParametersBase = source
            For Each pi As System.Reflection.PropertyInfo In Me.GetType.GetProperties()
                If source.ContainsKey(pi.Name) Then
                    Try
                        pi.SetValue(Me, source(pi.Name), Nothing)
                    Catch ex As Exception

                    End Try
                End If
            Next pi

        Catch ex As Exception
        End Try
    End Sub

#End Region

End Class

#End Region

#End Region

