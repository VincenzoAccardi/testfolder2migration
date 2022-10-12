Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class TaCustDataRegRec
    Inherits TPDotnet.Pos.TaBaseRec
    Implements TPDotnet.IT.Common.Pos.ITaCustDataRegRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaFreeCustomerDataRec
    ' ---------- ---------- ---------- ----------
    ' In this record we fill the common customer data
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2012, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Properties"

    ''' <summary>
    ''' gets the record type (sId)
    ''' </summary>
    ''' <value></value>
    ''' <returns>Italy_PosDef.TARecTypes.iTA_ZREPORT</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Italy_PosDef.TARecTypes.iTA_CUST_DATA_REG_REC
        End Get
    End Property


    ''' <summary>
    ''' gets the object name
    ''' </summary>
    ''' <value></value>
    ''' <returns>"ZREPORT"</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "TA_CUST_DATA_REG"
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
    ''' gets/sets the customer first name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szFirstName() As String Implements ITaCustDataRegRec.szFirstName
        Get
            Return m.Fields_Value("szFirstName")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szFirstName") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the customer last name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szLastName() As String Implements ITaCustDataRegRec.szLastName
        Get
            Return m.Fields_Value("szLastName")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szLastName") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the customer birth date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szBirthDate() As String Implements ITaCustDataRegRec.szBirthDate
        Get
            Return m.Fields_Value("szBirthDate")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szBirthDate") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the customer birth city
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szBirthCity() As String Implements ITaCustDataRegRec.szBirthCity
        Get
            Return m.Fields_Value("szBirthCity")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szBirthCity") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the customer address
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szAddress() As String Implements ITaCustDataRegRec.szAddress
        Get
            Return m.Fields_Value("szAddress")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szAddress") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the customer Country
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szState() As String Implements ITaCustDataRegRec.szState
        Get
            Return m.Fields_Value("szState")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szState") = Value
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

            ' Standard fields
            ' ---------------
            m.Append("szFirstName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szLastName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szBirthDate", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szBirthCity", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szAddress", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szState", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

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
        Return New TaCustDataRegRec
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

            ' we will now get the Presentation from Global Object Receipt
            PresentationKey = Italy_PosDef.TARecTypes.iTA_CUST_DATA_REG_REC & ".1"

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
