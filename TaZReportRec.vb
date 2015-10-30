Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class TaZReportRec
    Inherits TPDotnet.Pos.TaBaseRec
    Implements TPDotnet.IT.Common.Pos.ITaZReportRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaZReportRec
    ' ---------- ---------- ---------- ----------
    ' In this record we fill all informations returned from the fiscal printer before calling the PrintZReport
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
    ''' <returns>Italy_PosDef.TARecTypes.iTA_ZREPORT</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Italy_PosDef.TARecTypes.iTA_ZREPORT
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
            Return "IT_ZREPORT"
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
    ''' gets/sets the fiscal daily total value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dFiscalDailyTotal() As Double Implements ITaZReportRec.dFiscalDailyTotal
        Get
            dFiscalDailyTotal = m.Fields_Value("dFiscalDailyTotal")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dFiscalDailyTotal") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the fiscal grand total value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property dFiscalGrandTotal() As Double Implements ITaZReportRec.dFiscalGrandTotal
        Get
            dFiscalGrandTotal = m.Fields_Value("dFiscalGrandTotal")
        End Get
        Set(ByVal Value As Double)
            m.Fields_Value("dFiscalGrandTotal") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the fiscal zreport number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lZReportNmbr() As Integer Implements ITaZReportRec.lZReportNmbr
        Get
            lZReportNmbr = m.Fields_Value("lZReportNmbr")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lZReportNmbr") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the fiscal receipt number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lFiscalReceiptNmbr() As Integer Implements ITaZReportRec.lFiscalReceiptNmbr
        Get
            lFiscalReceiptNmbr = m.Fields_Value("lFiscalReceiptNmbr")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lFiscalReceiptNmbr") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the non fiscal receipt number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lNonFiscalReceiptNmbr() As Integer Implements ITaZReportRec.lNonFiscalReceiptNmbr
        Get
            lNonFiscalReceiptNmbr = m.Fields_Value("lNonFiscalReceiptNmbr")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lNonFiscalReceiptNmbr") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the fiscal printer ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szFiscalPrinterID() As String Implements ITaZReportRec.szFiscalPrinterID
        Get
            szFiscalPrinterID = m.Fields_Value("szFiscalPrinterID")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szFiscalPrinterID") = Value
        End Set
    End Property

    ''' <summary>
    ''' gets/sets the fiscal printer firmware
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szFiscalPrinterFW() As String Implements ITaZReportRec.szFiscalPrinterFW
        Get
            szFiscalPrinterFW = m.Fields_Value("szFiscalPrinterFW")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szFiscalPrinterFW") = Value
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
            m.Append("dFiscalDailyTotal", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)
            m.Append("dFiscalGrandTotal", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)

            m.Append("lFiscalReceiptNmbr", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("lNonFiscalReceiptNmbr", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("lZReportNmbr", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

            m.Append("szFiscalPrinterFW", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szFiscalPrinterID", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

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
        Return New TaZReportRec
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
            PresentationKey = Italy_PosDef.TARecTypes.iTA_ZREPORT & ".1"

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
