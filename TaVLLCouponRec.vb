Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
Imports System
Public Class TaVLLCouponRec
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
            Return Italy_PosDef.TARecTypes.iTA_VLL_COUPON
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
            Return "IT_VLLCOUPON"
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
    ''' gets/sets the start validation of coupon
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szValidFrom() As String
        Get
            Return m.Fields_Value("szValidFrom")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szValidFrom") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the end validation
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szValidTo() As String
        Get
            Return m.Fields_Value("szValidTo")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szValidTo") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the value of coupon
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
    ''' gets/sets the template number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lTemplateNumber() As Integer
        Get
            Return m.Fields_Value("lTemplateNumber")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lTemplateNumber") = Value
        End Set
    End Property

    Public Overridable Property lCodeType() As Integer
        Get
            Return m.Fields_Value("lCodeType")
        End Get
        Set(value As Integer)
            m.Fields_Value("lCodeType") = value
        End Set
    End Property


    Private _isVisible As Boolean

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
            m.Append("szBarcode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("lType", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("szValidFrom", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szValidTo", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("dValue", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("lTemplateNumber", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("lCodeType", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

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
        Return New TaVLLCouponRec
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

    ''' <summary>
    ''' Sets the presentation key and gets the associated presentation lines
    ''' </summary>
    ''' <param name="theDecive"></param>
    ''' <param name="thegReceipt"></param>
    ''' <param name="bTrainingMode"></param>
    ''' <returns>
    ''' The string which will be shown on the receipt
    ''' </returns>
    ''' <remarks></remarks>
    Public Overrides Function GetPresentation(ByRef theDecive As Short, ByRef thegReceipt As gReceipt, ByRef bTrainingMode As Integer) As String
        Dim PresentationKey As String

        GetPresentation = ""

        Try
            LOG_Info(getLocationString("GetPresentation"), "starting")
            If Not bIsVisible Then
                Exit Function
            End If
            ' we will now get the Presentation from Global Object Receipt
            PresentationKey = Italy_PosDef.TARecTypes.iTA_VLL_COUPON & "." & Me.lCodeType & Me.lTemplateNumber

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
