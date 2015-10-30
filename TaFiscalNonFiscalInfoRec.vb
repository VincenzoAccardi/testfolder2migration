Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class TaFiscalNonFiscalInfoRec
    Inherits TPDotnet.Pos.TaBaseRec


#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaFiscalNonFiscalInfoRec
    ' ---------- ---------- ---------- ----------
    ' In this record we fill all informations about a fiscal receipt that contains also non fiscal sale.
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Properties"

    ''' <summary>
    ''' gets the record type (sId)
    ''' </summary>
    ''' <value></value>
    ''' <returns>TPDotnet.Italy.Common.Pos.TARecTypes.iTA_HELLO_WORLD</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Italy_PosDef.TARecTypes.iTA_FISCAL_NON_FISCAL_INFO
        End Get
    End Property

    ''' <summary>
    ''' gets the object name
    ''' </summary>
    ''' <value></value>
    ''' <returns>"BISTROTABLE"</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "FISCAL_NON_FISCAL_INFO"
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
    ''' gets/sets the fiscal total value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dFiscalTotal() As Double
        Get
            dFiscalTotal = m.Fields_Value("dFiscalTotal")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dFiscalTotal") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the non fiscal total value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dNonFiscalTotal() As Double
        Get
            dNonFiscalTotal = m.Fields_Value("dNonFiscalTotal")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dNonFiscalTotal") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the entire TA total value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dTaTotal() As Double
        Get
            dTaTotal = m.Fields_Value("dTaTotal")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dTaTotal") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the entire TA paid value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dTaPaid() As Double
        Get
            dTaPaid = m.Fields_Value("dTaPaid")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dTaPaid") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the real change in amount
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dChange() As Double
        Get
            dChange = m.Fields_Value("dChange")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dChange") = Value
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
            m.Append("dFiscalTotal", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("dNonFiscalTotal", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("dTaTotal", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("dTaPaid", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("dChange", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)

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
        Return New TaFiscalNonFiscalInfoRec
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
            PresentationKey = Italy_PosDef.TARecTypes.iTA_FISCAL_NON_FISCAL_INFO & ".1"

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
