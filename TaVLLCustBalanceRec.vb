Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
Imports System
Public Class TaVLLCustBalanceRec
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
            Return Italy_PosDef.TARecTypes.iTA_VLL_CUST_BALANCE
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
            Return "IT_VLLBALANCE"
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

    ''' <summary>
    ''' gets/sets the type
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lType() As Integer
        Get
            Return m.Fields_Value("lType")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lType") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the balance value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dValue() As Double
        Get
            Return m.Fields_Value("dValue")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dValue") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the last update
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szLastUpdate() As String
        Get
            Return m.Fields_Value("szLastUpdate")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szLastUpdate") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the last update
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szDiscountDescription() As String
        Get
            Return m.Fields_Value("szDiscountDescription")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szDiscountDescription") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the Number of TA
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lNumber() As Integer
        Get
            Return m.Fields_Value("lNumber")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lNumber") = Value
        End Set
    End Property
    ''' <summary>
    ''' gets/sets the barcode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szBarcode() As String
        Get
            Return m.Fields_Value("szBarcode")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szBarcode") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the CodeAvantage
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szCodeAvantage() As String
        Get
            Return m.Fields_Value("szCodeAvantage")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szCodeAvantage") = Value
        End Set
    End Property
    ''' <summary>
    ''' gets/sets the CampagneIdentifiant
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szCampagneIdentifiant() As String
        Get
            Return m.Fields_Value("szCampagneIdentifiant")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szCampagneIdentifiant") = Value
        End Set
    End Property

    Private _isVisible As Boolean = False
    Public Property bIsVisible() As Boolean
        Get
            Return _isVisible
        End Get
        Set(ByVal Value As Boolean)
            _isVisible = Value
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
            m.Append("lType", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("dValue", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("szLastUpdate", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szDiscountDescription", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("lNumber", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("szBarcode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szCodeAvantage", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szCampagneIdentifiant", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("bIsVisible", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

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
        Return New TaVLLCustBalanceRec
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

    Public Overrides Function GetPresentation(ByRef theDecive As Short, ByRef thegReceipt As gReceipt, ByRef bTrainingMode As Integer) As String
        Dim PresentationKey As String

        GetPresentation = ""

        Try
            LOG_Info(getLocationString("GetPresentation"), "starting")
            If Not bIsVisible Then
                Exit Function
            End If
            ' we will now get the Presentation from Global Object Receipt
            PresentationKey = Italy_PosDef.TARecTypes.iTA_VLL_CUST_BALANCE & "." & lType

            'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            GetPresentation = GetTheLines(theDecive, Me, thegReceipt, bTrainingMode, PresentationKey)

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetPresentation"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetPresentation"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("GetPresentation"), "Function GetPresentation returns " & GetPresentation.ToString)
        End Try

    End Function
#End Region

End Class
