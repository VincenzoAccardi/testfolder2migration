Imports TPDotnet.Pos
Imports System
Imports Microsoft.VisualBasic
Public Class TaMediaGroupRec
    Inherits TPDotnet.Pos.TaBaseRec

    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Pos.TARecTypes.iTA_MEDIA_GROUP
        End Get
    End Property

    ''' <summary>
    ''' gets the object name
    ''' </summary>
    ''' <value>"MEDIA_GROUP"</value>
    ''' <remarks>This property will be serialized to ta xml file.</remarks>
    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "MEDIA_GROUP"
        End Get
    End Property

    Public Overridable Property dAmount() As Decimal
        Get
            dAmount = m.Fields_Value("dAmount")
        End Get
        Set(ByVal Value As Decimal)
            m.Fields_Value("dAmount") = Value
        End Set
    End Property
    Public Overridable Property lCount() As Integer
        Get
            lCount = m.Fields_Value("lCount")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lCount") = Value
        End Set
    End Property
    Public Overridable Property szDescription() As String
        Get
            szDescription = m.Fields_Value("szDescription")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szDescription") = Value
        End Set
    End Property
    Public Overridable Property szMediaType() As String
        Get
            szMediaType = m.Fields_Value("szMediaType")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szMediaType") = Value
        End Set
    End Property
    Public Overridable Property szGroupType() As String
        Get
            szGroupType = m.Fields_Value("szGroupType")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szGroupType") = Value
        End Set
    End Property

    Public Overridable Property szPrintCode() As String
        Get
            szPrintCode = m.Fields_Value("szPrintCode")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szPrintCode") = Value
        End Set
    End Property


#Region "New/Finalize"

    ''' <summary>
    ''' Define standard fields for zipcode object
    ''' </summary>
    Protected Overrides Sub DefineFields()
        Try
            LOG_FuncStart(getLocationString("DefineFields"))

            MyBase.DefineFields()

            ' create lHdrRef
            m.Append("lHdrRef", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

            ' append the fields to the recordset
            m.Append("dAmount", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("lCount", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("szDescription", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szMediaType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szGroupType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szPrintCode", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
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

#End Region

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
End Class
