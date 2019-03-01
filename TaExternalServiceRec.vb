Option Strict Off
Option Explicit On
Imports System
Imports TPDotnet.Pos
Imports TPDotnet.IT.Common.Pos
Imports Microsoft.VisualBasic

Public Class TaExternalServiceRec : Inherits TPDotnet.Pos.TaBaseRec

#Region "Documentation"
    ' Tp.net Pos DataClass TaExternalServiceRec
    ' -----------------------------------
    ' Author : Emanuele Gualtierotti
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Variablen"
#End Region

    Public Enum ExternalServiceStatus
        Unkown
        PreChecked
        Activated
        Deleted
    End Enum

#Region "Properties"

    ''' <summary>
    ''' gets the record type (sId)
    ''' </summary>
    ''' <value></value>
    ''' <returns>TPDotnet.Italy.Common.Pos.TARecTypes.iTA_BITSTRO_TABLE</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE
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
            Return "EXTERNAL_SERVICE"
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
    ''' gets/sets the receipt returned fron host
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property szReceipt() As String
        Get
            szReceipt = m.Fields_Value("szReceipt")
        End Get
        Set(ByVal Value As String)
            szOriginalReceipt = Value
            Dim szRec As String = Value
            szRec = Replace(szRec, Microsoft.VisualBasic.vbCrLf, Microsoft.VisualBasic.vbLf)
            szRec = Replace(szRec, Microsoft.VisualBasic.vbCr, Microsoft.VisualBasic.vbLf)
            szRec = Replace(szRec, Microsoft.VisualBasic.vbLf, Microsoft.VisualBasic.vbCrLf)

            m.Fields_Value("szReceipt") = "\x{FONTA,CENTER} " + szRec.Replace(vbCrLf, " \x" + vbCrLf + "\x{FONTA,CENTER} ") + " \x" + vbCrLf

        End Set
    End Property

    Public Overridable Property szOriginalReceipt() As String
        Get
            szOriginalReceipt = m.Fields_Value("szOriginalReceipt")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szOriginalReceipt") = Value
        End Set
    End Property

    ''' <summary>
    ''' flag : print this node on receipt? (used by GetPresentation)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property bPrintReceipt() As Integer
        Get
            bPrintReceipt = m.Fields_Value("bPrintReceipt")
        End Get
        Set(ByVal value As Integer)
            m.Fields_Value("bPrintReceipt") = value
        End Set
    End Property

    ''' <summary>
    ''' Type of service.
    ''' Values for Item:
    ''' GIFTCARD
    ''' PHONERECHARGE
    ''' EXTERNALGIFTCARD
    ''' Values for Media:
    ''' EMEAL
    ''' MEAL
    ''' EFT
    ''' GIFTCARD
    ''' SATISPAY
    ''' JIFFY
    ''' BITCOIN
    ''' PAYFAST
    ''' </summary>
    ''' <returns></returns>
    Public Overridable Property szServiceType() As String
        Get
            szServiceType = m.Fields_Value("szServiceType")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szServiceType") = value
        End Set
    End Property

    ''' <summary>
    ''' Number of Copies
    ''' </summary>
    ''' <returns></returns>
    Public Overridable Property lCopies() As Integer
        Get
            lCopies = m.Fields_Value("lCopies")
        End Get
        Set(ByVal value As Integer)
            m.Fields_Value("lCopies") = value
        End Set
    End Property
    ''' <summary>
    ''' Record Status
    ''' Unkown
    ''' PreChecked
    ''' Activated
    ''' Deleted
    ''' </summary>
    ''' <returns></returns>
    Public Overridable Property szStatus() As String
        Get
            szStatus = m.Fields_Value("szStatus")
        End Get
        Set(ByVal value As String)
            m.Fields_Value("szStatus") = value
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
            m.Append("szReceipt", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("bPrintReceipt", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("lCopies", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

            m.Append("szServiceType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szStatus", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

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
            m.Fields_Value("bPrintReceipt") = 1

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
        Return New TaExternalServiceRec
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

            If bPrintReceipt = 0 Then
                Exit Function
            End If

            Dim szRec As String = szReceipt

            szRec = Replace(szRec, Microsoft.VisualBasic.vbCrLf, Microsoft.VisualBasic.vbLf)
            szRec = Replace(szRec, Microsoft.VisualBasic.vbCr, Microsoft.VisualBasic.vbLf)
            szRec = Replace(szRec, Microsoft.VisualBasic.vbLf, Microsoft.VisualBasic.vbCrLf)

            'GetPresentation = "\x{FONTA,CENTER} " + szRec.Replace(vbCrLf, " \x" + vbCrLf + "\x{FONTA,CENTER} ") + "\x"
            GetPresentation = szRec
            'If Not String.IsNullOrEmpty(szReceipt) AndAlso szReceipt.IndexOf(Microsoft.VisualBasic.Constants.vbLf) >= 0 Then _
            '    GetPresentation = szReceipt.Replace(Microsoft.VisualBasic.Constants.vbLf, _
            '                                        Microsoft.VisualBasic.Constants.vbCrLf)

            Exit Function

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
