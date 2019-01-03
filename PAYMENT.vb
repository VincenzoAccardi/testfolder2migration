Option Strict Off
Option Explicit On

Imports System
Imports System.Text
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class PAYMENT : Inherits TPDotnet.Pos.PAYMENT : Implements TPDotnet.IT.Common.Pos.IFiscalPAYMENT

#Region "Documentation"
    ' PAYMENT
    ' ---------- ---------- ---------- ----------
    ' the class implements one PAYMENT
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2011, All rights reserved.
    ' -----------------------------------
#End Region

    Protected m_bITFiscalNotPaid As Integer
    Protected m_dITTxHALO As Decimal
    Protected m_bITCheckCashHalo As Decimal
    Protected m_bPaymentRoundableAt5Cent As Integer

#Region "PAYMENT Properties"

    ''' <summary>
    ''' gets / sets a flag indicating if this media member must be 
    ''' checked against the parameter ModPayment.CASH_HALO
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property bITCheckCashHalo() As Integer Implements IFiscalPAYMENT.bITCheckCashHalo
        Get
            bITCheckCashHalo = m_bITCheckCashHalo
        End Get
        Set(ByVal value As Integer)
            m_bITCheckCashHalo = value
            m.Fields_Value("bITCheckCashHalo") = value
        End Set
    End Property

    ''' <summary>
    ''' gets / sets a value indicating the HALO of this mediamember for the whole transaction
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dITTxHALO() As Decimal Implements IFiscalPAYMENT.dITTxHALO
        Get
            dITTxHALO = m_dITTxHALO
        End Get
        Set(ByVal value As Decimal)
            m_dITTxHALO = value
            m.Fields_Value("dITTxHALO") = value
        End Set
    End Property

    ''' <summary>
    ''' gets / sets a flag indicating if this media must be 
    ''' registered at fiscal printer as not paid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property bITFiscalNotPaid() As Integer Implements IFiscalPAYMENT.bITFiscalNotPaid
        Get
            bITFiscalNotPaid = m_bITFiscalNotPaid
        End Get
        Set(ByVal value As Integer)
            m_bITFiscalNotPaid = value
            m.Fields_Value("bITFiscalNotPaid") = value
        End Set
    End Property

    ''' <summary>
    ''' gets / sets a flag indicating if this media must be 
    ''' usated for payments roundables at 5 cent.
    ''' </summary>
    ''' <value></value>
    ''' <returns>True = Payment of type Roundable - False Payment Normal</returns>
    ''' <remarks></remarks>
    Public Overridable Property bPaymentRoundableAt5Cent() As Integer Implements IFiscalPAYMENT.bPaymentRoundableAt5Cent
        Get
            bPaymentRoundableAt5Cent = m_bPaymentRoundableAt5Cent
        End Get
        Set(ByVal value As Integer)
            m_bPaymentRoundableAt5Cent = value
            m.Fields_Value("bPaymentRoundableAt5Cent") = value
        End Set
    End Property

    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "PAYMENT"
        End Get
    End Property

#End Region

#Region "Overrides"

    Protected Overrides Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

    Protected Overrides Sub DefineFields()

        Try
            LOG_FuncStart(getLocationString("DefineFields"))

            MyBase.DefineFields()

            ' Custom fields
            ' ---------------
            m.Append("bITFiscalNotPaid", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("dITTxHALO", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("bITCheckCashHalo", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("bPaymentRoundableAt5Cent", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

            Exit Sub

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DefineFields"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DefineFields"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DefineFields"), "")
        End Try
    End Sub

    Protected Overrides Sub InitFields()
        Try
            LOG_FuncStart(getLocationString("InitFields"))

            MyBase.InitFields()

            ' Standard fields
            ' ---------------
            ' Standard fields, needs only to initialized when differ from "" (string) or 0 (numeric)
            ' In definefields the append do the default init: string = "" and numeric =0
            ' ---------------------------------------------------------------------------
            m.Fields_Value("bITFiscalNotPaid") = 0
            m.Fields_Value("dITTxHALO") = 9999999.99
            m.Fields_Value("bITCheckCashHalo") = 0
            m.Fields_Value("bPaymentRoundableAt5Cent") = 0

            Exit Sub


        Catch ex As Exception
            Try
                LOG_Error(getLocationString("InitFields"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("InitFields"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("InitFields"), "")
        End Try
    End Sub

    Public Sub New()

        MyBase.New()

    End Sub

#End Region

End Class
