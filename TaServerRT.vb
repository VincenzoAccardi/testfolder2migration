Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
Imports System
Public Class TaServerRT
    Inherits TPDotnet.Pos.TaBaseRec
#Region "Properties"
    ''' <summary>
    ''' gets the record type (sId)
    ''' </summary>
    ''' <value></value>
    ''' <returns>Italy_PosDef.TARecTypes.iTA_VLLBALANCE</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Italy_PosDef.TARecTypes.iTA_SERVER_RT
        End Get
    End Property

    ''' <summary>
    ''' gets the object name
    ''' </summary>
    ''' <value></value>
    ''' <returns>"VLLBALANCE"</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "IT_SERVERRT"
        End Get
    End Property

    ''' <summary>
    ''' gets/sets the header reference
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property lHdrRef() As Integer
        Get
            Return m.Fields_Value("lHdrRef")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lHdrRef") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets the header (TA)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property theHdr() As TaBaseHdr
        Get
            theHdr = m_Hdr
        End Get
    End Property
    Public Overridable Property szServerRTName() As String
        Get
            Return m.Fields_Value("szServerRTName")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szServerRTName") = value
        End Set

    End Property
    Public Overridable Property szFiscalCodeOrVatNumber() As String
        Get
            Return m.Fields_Value("szFiscalCodeOrVatNumber")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szFiscalCodeOrVatNumber") = value
        End Set

    End Property
    Public Overridable Property szLotteryStoreCode() As String
        Get
            Return m.Fields_Value("szLotteryStoreCode")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szLotteryStoreCode") = value
        End Set

    End Property
    Public Overridable Property szLotteryCode() As String
        Get
            Return m.Fields_Value("szLotteryCode")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szLotteryCode") = value
        End Set

    End Property
    Public Overridable Property szLotteryServerRTCode() As String
        Get
            Return m.Fields_Value("szLotteryServerRTCode")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szLotteryServerRTCode") = value
        End Set

    End Property
    Public Overridable Property lNumeroAzzeramento() As String
        Get
            Return m.Fields_Value("lNumeroAzzeramento")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("lNumeroAzzeramento") = value
        End Set

    End Property
    Public Overridable Property lNumeroDocumento() As String
        Get
            Return m.Fields_Value("lNumeroDocumento")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("lNumeroDocumento") = value
        End Set

    End Property
    Public Overridable Property szCCDC() As String
        Get
            Return m.Fields_Value("szCCDC")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szCCDC") = value
        End Set

    End Property
    Public Overridable Property szTaType() As String
        Get
            Return m.Fields_Value("szTaType")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szTaType") = value
        End Set
    End Property
    Public Overridable Property szDate() As String
        Get
            Return m.Fields_Value("szDate")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szDate") = value
        End Set
    End Property
    Public Overridable Property szRTFileName() As String
        Get
            Return m.Fields_Value("szRTFileName")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szRTFileName") = value
        End Set
    End Property
    Public Overridable Property TransactionType() As String
        Get
            Return m.Fields_Value("TransactionType")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("TransactionType") = value
        End Set
    End Property
    Public Overridable Property szOriginalRTDate() As String
        Get
            Return m.Fields_Value("szOriginalRTDate")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szOriginalRTDate") = value
        End Set
    End Property
#End Region

#Region "New/Finalize"

    ''' <summary>
    ''' Define standard fields for the TaBistroTableRec object
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub DefineFields()

        Try
            LOG_Info(getLocationString("DefineFields"), "starting")

            MyBase.DefineFields()

            ' create lHdrRef
            m.Append("lHdrRef", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

            '' Standard fields
            '' ---------------
            m.Append("szServerRTName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szFiscalCodeOrVatNumber", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szLotteryCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szLotteryStoreCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szLotteryServerRTCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("lNumeroAzzeramento", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("lNumeroDocumento", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szCCDC", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szTaType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szDate", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szRTFileName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("TransactionType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szOriginalRTDate", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

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

    ''' <summary>
    ''' Initialize the fields defined in DefineFields
    ''' Set default values
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub InitFields()

        Try
            LOG_Info(getLocationString("InitFields"), "starting")

            MyBase.InitFields()

            'Alpha numeric fields are initialized by '' and numeric fields are initialized by 0.

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
#End Region

#Region "Overwritten functionality"
    ''' <summary>
    ''' create new instance of this class
    ''' </summary>
    ''' <returns>the new instance of this class</returns>
    ''' <remarks></remarks>
    Public Overrides Function CreateMe() As TPDotnet.Pos.TaBaseRec
        ' create new instance of this class
        Return New TaServerRT
    End Function


#End Region

#Region "Private Functions"

    ''' <summary>
    ''' Gets the name of this object and appends it with the actMethod.
    ''' </summary>
    ''' <param name="actMethode">The actual method as String</param>
    ''' <returns>TypeName + method name</returns>
    ''' <remarks></remarks>
    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
