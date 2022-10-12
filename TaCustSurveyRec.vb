Imports TPDotnet.Pos
Imports Microsoft.VisualBasic
Imports System

Public Class TaCustSurveyRec
    Inherits TPDotnet.Pos.TaCustSurveyRec


    ''' <summary>
    ''' sets / gets the szITAnswer6.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITAnswer6() As String
        Get
            szITAnswer6 = m.Fields_Value("szITAnswer6")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITAnswer6") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITAnswer7.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITAnswer7() As String
        Get
            szITAnswer7 = m.Fields_Value("szITAnswer7")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITAnswer7") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITAnswer8.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITAnswer8() As String
        Get
            szITAnswer8 = m.Fields_Value("szITAnswer8")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITAnswer8") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITAnswer9.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITAnswer9() As String
        Get
            szITAnswer9 = m.Fields_Value("szITAnswer9")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITAnswer9") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITAnswer10.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITAnswer10() As String
        Get
            szITAnswer10 = m.Fields_Value("szITAnswer10")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITAnswer10") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITText6.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITText6() As String
        Get
            szITText6 = m.Fields_Value("szITText6")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITText6") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITText7.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITText7() As String
        Get
            szITText7 = m.Fields_Value("szITText7")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITText7") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITText8.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITText8() As String
        Get
            szITText8 = m.Fields_Value("szITText8")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITText8") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITText9.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITText9() As String
        Get
            szITText9 = m.Fields_Value("szITText9")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITText9") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the szITText10.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szITText10() As String
        Get
            szITText10 = m.Fields_Value("szITText10")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITText10") = Value
        End Set
    End Property

#Region "New/Finalize"
    ''' <summary>
    ''' Define standard fields for the TaBistroTableRec object
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub DefineFields()

        Try
            LOG_Info(getLocationString("DefineFields"), "starting")

            MyBase.DefineFields()

            m.Append("szITAnswer6", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITAnswer7", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITAnswer8", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITAnswer9", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITAnswer10", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITText6", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITText7", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITText8", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITText9", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szITText10", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            ' Standard fields
            ' ---------------
            'm.Append("dValPaySummarized", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)

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
    ''' Gets the name of this object and appends it with the actMethod.
    ''' </summary>
    ''' <param name="actMethode">The actual method as String</param>
    ''' <returns>TypeName + method name</returns>
    ''' <remarks></remarks>
    Protected Overrides Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region
End Class
