Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Xml
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility

Imports VB = Microsoft.VisualBasic
Module UpdWSStatus

    Public Const MAX_RETRIES As Short = 5

    ' update the status in ComputerComponentStatus table by using the TPWSStatusService
    '  Parameter like - calling Program
    '                           Workstation
    '                           Status
    '                           Language
    '                           Operator
    '                           Mode
    '                           SessionID
    '  will be transfered via xml via command line

    Public Sub Main()
        Dim xmlDoc As Xml.XmlDocument
        Dim myXMLRootNode As Xml.XmlNode
        Dim myXMLNode As Xml.XmlNode

        Dim myTPWSStatusService As TPDotnet.Services.TPWSStatusService.TPWSStatusHelper

        Dim szParameter As String
        Dim szProgram As String = ""
        Dim szStatus As String = ""
        Dim szLanguage As String = ""
        Dim szMode As String = ""
        Dim szSessionID As String = ""
        Dim iOperator As Integer = 0
        Dim iWorkstation As Integer = 0
        Dim sStatus As Short = 0

        Dim bRet As Boolean = False
        Dim i As Short

        Try
            LOG_Error(getLocationString("UpdWSStatus.Main"), "Start")

            Dim pid As Integer = System.Diagnostics.Process.GetCurrentProcess().Id
            For Each p As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess.ProcessName)
                If p.Id <> pid Then
                    LOG_Error(getLocationString("UpdWSStatus.Main"), "Another instance of this program is already running and will be killed now (ID=" + p.Id.ToString + ")")
                    p.Kill()
                Else
                    LOG_Error(getLocationString("UpdWSStatus.Main"), "Don't kill this process because it is the current instance (ID=" + p.Id.ToString + ")")
                End If
            Next

            ' get the parameter
            szParameter = VB.Command()

            LOG_Error(getLocationString("UpdWSStatus.Main"), "Parameter " & szParameter)

            If szParameter.Length > 0 Then
                ' we got parameter
                ' ok lets create a DOM Document
                xmlDoc = New Xml.XmlDocument
                xmlDoc.LoadXml(szParameter)

                ' DOMDocument is filled, now we read all the childnotes
                myXMLRootNode = xmlDoc.DocumentElement
                For Each myXMLNode In myXMLRootNode.ChildNodes
                    Select Case myXMLNode.Name.ToUpper(New CultureInfo("en-US", False))
                        Case "SZPROGRAM"
                            szProgram = myXMLNode.InnerText
                        Case "IWORKSTATION"
                            iWorkstation = myXMLNode.InnerText
                        Case "SSTATUS"
                            sStatus = myXMLNode.InnerText
                        Case "SZLANGUAGE"
                            szLanguage = myXMLNode.InnerText
                        Case "IOPERATOR"
                            iOperator = myXMLNode.InnerText
                        Case "SZMODE"
                            szMode = myXMLNode.InnerText
                        Case "SZSESSIONID"
                            szSessionID = myXMLNode.InnerText
                    End Select
                Next myXMLNode

                LOG_Error(getLocationString("UpdWSStatus.Main"), "used fields: " & _
                                                                 " Program: " & szProgram & _
                                                                 " Workstation: " & iWorkstation & _
                                                                 " Status: " & sStatus & _
                                                                 " Language: " & szLanguage & _
                                                                 " Operator: " & iOperator & _
                                                                 " Mode: " & szMode & _
                                                                 " SessionID: " & szSessionID)

                ' set status
                ' ==========
                Select Case sStatus
                    Case STATUS_SIGN_ON
                        ' sign on
                        szStatus = TPDotnet.Services.TPWSStatusService.Constants.ComponentStatus.SIGNON

                    Case STATUS_SIGN_ON_OFFLINE
                        ' sign on offline
                        szStatus = TPDotnet.Services.TPWSStatusService.Constants.ComponentStatus.SIGNONOFFLINE

                    Case STATUS_SIGN_OFF
                        ' sign off, short sign off
                        szStatus = TPDotnet.Services.TPWSStatusService.Constants.ComponentStatus.SIGNOFF

                    Case STATUS_SHORT_SIGN_OFF
                        szStatus = "SHORTSIGNOFF"

                    Case STATUS_BREAK, STATUS_LOCK
                        ' break, lock
                        szStatus = TPDotnet.Services.TPWSStatusService.Constants.ComponentStatus.BREAK

                    Case STATUS_RUNNING
                        szStatus = TPDotnet.Services.TPWSStatusService.Constants.ComponentStatus.RUNNING

                    Case STATUS_TERMINATED
                        szStatus = TPDotnet.Services.TPWSStatusService.Constants.ComponentStatus.TERMINATED

                    Case 999
                        szStatus = "FUNCTION_IS_RUNNING"

                    Case Else
                        szStatus = "UNKNOWN"

                End Select

                ' set status in ComputerComponentStatus
                ' =====================================
                myTPWSStatusService = New TPDotnet.Services.TPWSStatusService.TPWSStatusHelper _
                                          (szProgram, iWorkstation)
                While Not bRet AndAlso i < MAX_RETRIES

                    bRet = myTPWSStatusService.UpdateWSComponentStatus(szStatus, _
                                                                       szLanguage, _
                                                                       iOperator, _
                                                                       "Mode " & szMode & " SessionID " & szSessionID)
                    i = i + 1
                End While
                myTPWSStatusService.Dispose()

                If bRet Then
                    LOG_Error("UpdWSStatus.Main", "Status updated to " & szStatus & ", tries: " & i.ToString)
                Else
                    LOG_Error("UpdWSStatus.Main", "Status NOT updated to " & szStatus)
                End If
            End If

            LOG_Error(getLocationString("UpdWSStatus.Main"), "End")

            Exit Sub


        Catch ex As Exception
            Try
                LOG_Error(getLocationString("Main"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("Main"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("Main"), "")
        End Try

    End Sub

    Public Function getLocationString(ByRef actMethode As String) As String
        getLocationString = actMethode & " "
    End Function

End Module
