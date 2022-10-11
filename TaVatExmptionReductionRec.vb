Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class TaVatExmptionReductionRec
    Inherits TaBaseRec


#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaVatExmptionReductionRec
    ' ---------- ---------- ---------- ----------
    ' Ticket with items sold with exempt or reduced VAT
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2010, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Properties"

    ''' <summary>
    ''' gets the record type (sId)
    ''' </summary>
    ''' <value></value>
    ''' <returns>Italy_PosDef.TARecTypes.iTA_VAT_EXEMPTION_REDUCTION</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Italy_PosDef.TARecTypes.iTA_VAT_EXEMPTION_REDUCTION
        End Get
    End Property


    ''' <summary>
    ''' gets the object name
    ''' </summary>
    ''' <value></value>
    ''' <returns>"VAT_EXEMPTION_REDUCTION"</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "VAT_EXEMPTION_REDUCTION"
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
            Return m.Fields_value("lHdrRef")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_value("lHdrRef") = Value
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
    ''' gets/sets the szItemTaxGroupID to use for all items in the ticket
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szItemTaxGroupID() As String
        Get
            szItemTaxGroupID = m.Fields_Value("szItemTaxGroupID")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szItemTaxGroupID") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the szDescription of the selected ItemTaxGroup
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szDescription() As String
        Get
            szDescription = m.Fields_Value("szDescription")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szDescription") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the lPresentation. Used presentation, but for type "coupon" always presentation
    '''  2 is used
    ''' </summary>
    ''' <value>This is the presentation's extension</value>
    ''' <remarks></remarks>
    Public Overridable Property lPresentation() As Integer
        Get
            lPresentation = m.Fields_Value("lPresentation")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lPresentation") = Value
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
            m.Append("szItemTaxGroupID", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szDescription", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("lPresentation", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

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
        Return New TaVatExmptionReductionRec
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
            'PresentationKey = Italy_PosDef.TARecTypes.iTA_VAT_EXEMPTION_REDUCTION & ".1"
            'GetPresentation = GetTheLines(theDecive, Me, thegReceipt, bTrainingMode, PresentationKey)

            ' add specific lines...
            PresentationKey = Italy_PosDef.TARecTypes.iTA_VAT_EXEMPTION_REDUCTION & "." & Me.lPresentation
            GetPresentation = GetPresentation & GetTheLines(theDecive, Me, thegReceipt, bTrainingMode, PresentationKey)


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
