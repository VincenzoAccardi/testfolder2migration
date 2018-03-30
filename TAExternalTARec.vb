Imports TPDotnet.Pos
Imports System
Imports Microsoft.VisualBasic
''' <summary>
''' DataClass TAExternalTARec stores all needed information of an external transaction
''' </summary>
''' <remarks></remarks>
Public Class TAExternalTARec : Inherits TPDotnet.Pos.TaBaseRec

#Region "Documentation"
    ' Tp.net Pos DataClass TAExternalTARec
    ' -----------------------------------
    ' Author : Manuel Schultze
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf International GmbH,
    ' 33106 Paderborn, Germany, 2012 , All rights reserved.
    ' -----------------------------------
#End Region


#Region "Properties"


    ''' <summary>
    ''' Gets the sid. This function has to be overwritten
    ''' </summary>
    ''' <value>lSid</value>
    Public Overrides ReadOnly Property sId() As Short
        Get
            Return Pos.TARecTypes.iTA_EXTERNAL_TA
        End Get
    End Property

    ''' <summary>
    ''' gets the object name
    ''' </summary>
    ''' <value>"EXTERNAL_TA"</value>
    ''' <remarks>This property will be serialized to ta xml file.</remarks>
    Public Overrides ReadOnly Property szObjectName() As String
        Get
            Return "EXTERNAL_TA"
        End Get
    End Property

    ''' <summary>
    ''' sets / gets the external transactions file name
    ''' </summary>
    ''' <value>The filename of the transaction</value>
    ''' <remarks>This property will be serialized to ta xml file.</remarks>
    Public Overridable Property szExternalFileName() As String
        Get
            szExternalFileName = m.Fields_Value("szExternalFileName")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szExternalFileName") = Value
        End Set
    End Property

    Public Overridable Property szLinkToFileName() As String
        Get
            szLinkToFileName = m.Fields_Value("szLinkToFileName")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szLinkToFileName") = Value
        End Set
    End Property
    ''' <summary>
    ''' sets / gets the external transactionnumber
    ''' </summary>
    ''' <value>The external transactionnumber</value>
    ''' <remarks>This property will be serialized to ta xml file.</remarks>
    Public Overridable Property szExternalTransactionNumber() As String
        Get
            szExternalTransactionNumber = m.Fields_Value("szExternalTransactionNumber")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szExternalTransactionNumber") = Value
        End Set
    End Property

    ''' <summary>
    ''' sets / gets the customer number
    ''' </summary>
    ''' <value>The customer number</value>
    ''' <remarks>This property will be serialized to ta xml file.</remarks>
    Public Overridable Property szExternalCustomerNumber() As String
        Get
            szExternalCustomerNumber = m.Fields_Value("szExternalCustomerNumber")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szExternalCustomerNumber") = Value
        End Set
    End Property

    Public Overridable Property szExternalWorkStationNumber() As String
        Get
            szExternalWorkStationNumber = m.Fields_Value("szExternalWorkStationNumber")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szExternalWorkStationNumber") = Value
        End Set
    End Property

    Public Overridable Property bRescan() As Integer
        Get
            bRescan = m.Fields_Value("bRescan")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("bRescan") = Value
        End Set
    End Property

    Public Overridable Property lModulNmbrExt() As Integer
        Get
            lModulNmbrExt = m.Fields_Value("lModulNmbrExt")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lModulNmbrExt") = Value
        End Set
    End Property

    Public Overridable Property szExternalTaType() As String
        Get
            szExternalTaType = m.Fields_Value("szExternalTaType")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szExternalTaType") = Value
        End Set
    End Property
    Public Overridable Property szMovePath As String
        Get
            szMovePath = m.Fields_Value("szMovePath")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szMovePath") = Value
        End Set
    End Property

    ' Start / End
    Public Shared OPERATION_TYPE_START_EXTERNAL As String = "Start"
    Public Shared OPERATION_TYPE_END_EXTERNAL As String = "End"

    Public Shared DELETE_FILE_NO As String = "NO"
    Public Shared DELETE_FILE_AFTER_IMPORT As String = "AFTER_IMPORT"
    Public Shared DELETE_FILE_END_TA As String = "END_TA"
    Public Shared MOVE_FILE_AFTER_IMPORT As String = "MOVE_AFTER_IMPORT"
    Public Shared MOVE_FILE_END_TA As String = "MOVE_END_TA"

    Public Overridable Property szOperationType() As String
        Get
            szOperationType = m.Fields_Value("szOperationType")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szOperationType") = Value
        End Set
    End Property

    Public Overridable Property bConfirm() As Integer
        Get
            bConfirm = m.Fields_Value("bConfirm")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("bConfirm") = Value
        End Set
    End Property

    Public Overridable Property szTemplateFileIn() As String
        Get
            szTemplateFileIn = m.Fields_Value("szTemplateFileIn")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szTemplateFileIn") = Value
        End Set
    End Property
    Public Overridable Property szTemplateFileOut() As String
        Get
            szTemplateFileOut = m.Fields_Value("szTemplateFileOut")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szTemplateFileOut") = Value
        End Set
    End Property

    Public Overridable Property szStyleSheetName() As String
        Get
            szStyleSheetName = m.Fields_Value("szStyleSheetName")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szStyleSheetName") = Value
        End Set
    End Property

    Public Overridable Property szFilePathIn() As String
        Get
            szFilePathIn = m.Fields_Value("szFilePathIn")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szFilePathIn") = Value
        End Set
    End Property

    Public Overridable Property szFilePathOut() As String
        Get
            szFilePathOut = m.Fields_Value("szFilePathOut")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szFilePathOut") = Value
        End Set
    End Property

    Public Overridable Property szDeleteFile() As String
        Get
            szDeleteFile = m.Fields_Value("szDeleteFile")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szDeleteFile") = Value
        End Set
    End Property

    Public Overridable Property bConfirmEndRescan() As Integer
        Get
            bConfirmEndRescan = m.Fields_Value("bConfirmEndRescan")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("bConfirmEndRescan") = Value
        End Set
    End Property

    Public Overridable Property bAllowVoidLine() As Integer
        Get
            bAllowVoidLine = m.Fields_Value("bAllowVoidLine")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("bAllowVoidLine") = Value
        End Set
    End Property

    Public Overridable Property szRescanType() As String
        Get
            szRescanType = m.Fields_Value("szRescanType")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szRescanType") = Value
        End Set
    End Property

#End Region

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
            m.Append("szExternalFileName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szExternalTransactionNumber", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szExternalCustomerNumber", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szExternalWorkStationNumber", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("bIsBasket", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("bRescan", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("lModulNmbrExt", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("szExternalTaType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szOperationType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("bConfirm", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("szTemplateFileIn", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szTemplateFileOut", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szStyleSheetName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szFilePathIn", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szFilePathOut", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szDeleteFile", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("bConfirmEndRescan", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("bAllowVoidLine", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
            m.Append("szLinkToFileName", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szMovePath", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
            m.Append("szRescanType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

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

#Region "Overwritten functionalities"
    ''' <summary>
    ''' create new instance of this class
    ''' </summary>
    ''' <returns>
    ''' the new instance of this class
    ''' </returns>
    ''' <remarks></remarks>
    Public Overrides Function CreateMe() As TPDotnet.Pos.TaBaseRec
        ' create new instance of this class
        Return New TAExternalTARec
    End Function

    Public Overrides Function GetPresentation(ByRef theDecive As Short, ByRef thegReceipt As gReceipt, ByRef bTrainingMode As Integer) As String
        Dim PresentationKey As String
        Dim lPresentationKey As Integer = 0

        GetPresentation = ""

        Try
            LOG_Info(getLocationString("GetPresentation"), "starting")

            If Me.szOperationType = OPERATION_TYPE_START_EXTERNAL Then
                lPresentationKey = 1000
            ElseIf Me.szOperationType = OPERATION_TYPE_END_EXTERNAL Then
                lPresentationKey = 2000 ' End
            End If

            lPresentationKey = lPresentationKey + Me.lModulNmbrExt
            PresentationKey = Italy_PosDef.TARecTypes.iTA_EXTERNAL_TA & "." & lPresentationKey.ToString

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

    Public Overrides Function Clone(ByRef CloneFromObj As TPDotnet.Pos.TaBaseRec, _
                                    ByVal iCreateNmbr As Integer) As Boolean
        Dim myCloneObj As TAExternalTARec
        Clone = False

        Try
            If TypeOf (CloneFromObj) Is TAExternalTARec Then
                myCloneObj = CloneFromObj

                m.Fields_Value("szExternalFileName") = myCloneObj.szExternalFileName
                m.Fields_Value("szExternalTransactionNumber") = myCloneObj.szExternalTransactionNumber
                m.Fields_Value("szExternalCustomerNumber") = myCloneObj.szExternalCustomerNumber
                m.Fields_Value("szExternalWorkStationNumber") = myCloneObj.szExternalWorkStationNumber
                m.Fields_Value("bConfirm") = myCloneObj.bConfirm
                m.Fields_Value("bRescan") = myCloneObj.bRescan
                m.Fields_Value("lModulNmbrExt") = myCloneObj.lModulNmbrExt
                m.Fields_Value("szExternalTaType") = myCloneObj.szExternalTaType
                m.Fields_Value("szOperationType") = myCloneObj.szOperationType
                m.Fields_Value("szTemplateFileIn") = myCloneObj.szTemplateFileIn
                m.Fields_Value("szTemplateFileOut") = myCloneObj.szTemplateFileOut
                m.Fields_Value("szStyleSheetName") = myCloneObj.szStyleSheetName
                m.Fields_Value("szFilePathIn") = myCloneObj.szFilePathIn
                m.Fields_Value("szFilePathOut") = myCloneObj.szFilePathOut
                m.Fields_Value("szDeleteFile") = myCloneObj.szDeleteFile
                m.Fields_Value("bConfirmEndRescan") = myCloneObj.bConfirmEndRescan
                m.Fields_Value("bAllowVoidLine") = myCloneObj.bAllowVoidLine
                m.Fields_Value("szLinkToFileName") = myCloneObj.szLinkToFileName
                m.Fields_Value("szMovePath") = myCloneObj.szMovePath
                m.Fields_Value("szRescanType") = myCloneObj.szRescanType

                theHdr.bTaValid = myCloneObj.theHdr.bTaValid
                theHdr.lTaRefToCreateNmbr = iCreateNmbr

                Clone = True
            End If

        Catch ex As Exception
            LOG_Error(getLocationString("Clone"), ex)
        End Try
    End Function

#End Region


#Region "Private Functions"
    ''' <summary>
    ''' Gets the location string.
    ''' </summary>
    ''' <param name="actMethode">The actual method as String</param>
    ''' <returns>TypeName + method name</returns>
    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function
#End Region


End Class
