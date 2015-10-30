Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports System.Drawing

Public Class PhoneRechargeController
    Implements IPhoneRechargeActivation
    Implements IPhoneRechargeActivationPreCheck

#Region "Documentation"
    ' ********** ********** ********** **********
    ' Argentea EFT
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Basiglio, 2014, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Argentea specific"

    Protected ArgenteaCOMObject As ARGLIB.argpay

#End Region

#Region "Instance related functions"

    Public Sub New()
        Parameters = New ArgenteaParameters()
    End Sub

#End Region

#Region "Parameters"

    Protected Parameters As ArgenteaParameters

#End Region

#Region "Public functions"

#End Region

#Region "IPhoneRechargeActivationPreCheck"

    Public Function CheckPhoneRecharge(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IPhoneRechargeReturnCode Implements IPhoneRechargeActivationPreCheck.CheckPhoneRecharge
        CheckPhoneRecharge = IPhoneRechargeReturnCode.OK
        Dim funcName As String = "CheckPhoneRecharge"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As PhoneRechargeActivationParameters = New PhoneRechargeActivationParameters

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea CheckPhoneRecharge function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")
            p.LoadCommonFunctionParameter(parameters)

            ' call in check mode
            ShowWaitScreen(p.Controller, False, frm)
            ArgenteaCOMObject = New ARGLIB.argpay()

            Dim CSV As String = String.Empty
            Dim RetchkCode As Integer = Nothing
            RetchkCode = ArgenteaCOMObject.RichiestaPIN(p.PINCounter = p.GetNextPINCounter, p.Barcode, p.IntValue, p.ErrorMessage, p.TransactionID)
            If RetchkCode <> ArgenteaFunctionsReturnCode.OK Then
                CheckPhoneRecharge = IPhoneRechargeReturnCode.KO
                LOG_Debug(getLocationString(funcName), "Activation check for phone recharge  " & p.Barcode & " returns error: " & p.ErrorMessage & " RetCode:" & RetchkCode)
                CSV = "KO" & ";" & p.MessageOut & ";" & p.ErrorMessage
            Else
                LOG_Debug(getLocationString(funcName), "phone recharge number " & p.Barcode & " successfuly checked for activation")
                CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
            End If
            Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
            objTPTAHelperArgentea.HandleReturnString(p.Transaction, _
                                                        p.Controller, _
                                                        CSV, _
                                                        InternalArgenteaFunctionTypes.PhoneRechargeCheck, _
                                                        Me.Parameters)
            p.Status = ArgenteaPhoneRechargeStatus.ActivatedWithCheckMode

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally
            ShowWaitScreen(p.Controller, True, frm)
            ShowError(p)
        End Try
    End Function

#End Region

#Region "IPhoneRechargeActivation"

    Public Function ActivatePhoneRecharge(ByRef parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IPhoneRechargeReturnCode Implements IPhoneRechargeActivation.ActivatePhoneRecharge
        ActivatePhoneRecharge = IPhoneRechargeReturnCode.OK
        Dim funcName As String = "ActivatePhoneRecharge"

        Dim frm As System.Windows.Forms.Form = Nothing
        Dim p As PhoneRechargeActivationParameters = New PhoneRechargeActivationParameters

        Try
            ' collect the input parameters
            p.LoadCommonFunctionParameter(parameters)

            ' call in check mode
            ArgenteaCOMObject = New ARGLIB.argpay()
            For i As Integer = 1 To p.Transaction.taCollection.Count
                Dim MyTaBaseRec As TPDotnet.Pos.TaBaseRec = p.Transaction.GetTALine(i)

                Select Case MyTaBaseRec.sid
                    Case TPDotnet.Pos.PosDef.TARecTypes.iTA_ART_SALE
                        Try
                            p.ArticleRecord = MyTaBaseRec
                            If p.ArticleRecord.theHdr.bIsVoided = 0 AndAlso TypeOf p.ArticleRecord.ARTinArtSale Is TPDotnet.IT.Common.Pos.ART Then
                                Dim ITART As TPDotnet.IT.Common.Pos.ART = p.ArticleRecord.ARTinArtSale
                                If ITART.szITSpecialItemType = TPDotnet.IT.Common.Pos.PhoneRechargeItem Then

                                    Try
                                        Dim CSV As String = String.Empty
                                        ShowWaitScreen(p.Controller, False, frm)
                                        Dim RetActCode As Integer = ArgenteaFunctionsReturnCode.KO
                                        Try
                                            RetActCode = ArgenteaCOMObject.ConfermaPIN(p.TransactionID, p.PINCounter, p.IntValue, p.Barcode, p.PinID, p.MessageOut, p.ErrorMessage)
                                        Catch ex As Exception
                                            LOG_Error(getLocationString(funcName), "Erorr:" & ex.StackTrace)
                                            ShowWaitScreen(p.Controller, True, frm)
                                        End Try
                                        If RetActCode <> ArgenteaFunctionsReturnCode.OK Then
                                            ActivatePhoneRecharge = IPhoneRechargeReturnCode.KO
                                            ' Show an error for each gift card that cannot be definitely activated
                                            LOG_Debug(getLocationString(funcName), "Activation phone recharge  " & p.Barcode & " returns error: " & p.ErrorMessage & " RetCode:" & RetActCode)
                                            CSV = "KO" & ";" & p.MessageOut & vbCrLf & _
                                                    "ERRORE EMISSIONE RICARICA" & vbCrLf & _
                                                    p.ErrorMessage & vbCrLf & _
                                                    "Tran. ID: " & p.TransactionID & vbCrLf & _
                                                    "Barcode: " & p.Barcode & vbCrLf & _
                                                    "Value: " & Math.Round(p.IntValue / 100, 2) & vbCrLf & vbCrLf & " " & ";" & p.ErrorMessage
                                        Else
                                            LOG_Debug(getLocationString(funcName), "Phone recharge number " & p.Barcode & " successfuly activated")
                                            CSV = "OK" & ";" & p.MessageOut & ";" & p.ErrorMessage
                                        End If

                                        Dim objTPTAHelperArgentea As New TPTAHelperArgentea()
                                        objTPTAHelperArgentea.HandleReturnString(p.Transaction, _
                                                                                 p.Controller, _
                                                                                 CSV, _
                                                                                 InternalArgenteaFunctionTypes.PhoneRechargeActivation, _
                                                                                 Me.Parameters)
                                        p.Status = ArgenteaPhoneRechargeStatus.ActivatedDefinitively
                                    Catch ex As Exception
                                        LOG_Error(getLocationString(funcName), ex.Message)
                                    Finally
                                        ShowWaitScreen(p.Controller, True, frm)
                                        ShowError(p)
                                    End Try

                                End If

                            End If

                        Catch ex As Exception
                            LOG_Error(getLocationString(funcName), ex.Message)
                        End Try
                        Exit Select

                End Select
            Next i

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        Finally

        End Try
    End Function

#End Region

#Region "Overridable"

    Protected Overridable Sub ShowError(ByRef TheModCntr As TPDotnet.Pos.ModCntr, _
                                        ByRef err As String)
        Dim funcName As String = "ShowError"
        Dim szTranslatedError As String = err

        Try
            If Not TheModCntr Is Nothing AndAlso Not String.IsNullOrEmpty(err) Then

                LOG_Debug(getLocationString(funcName), err)

                szTranslatedError = getPosTxtNew(TheModCntr.contxt, "LevelITCommonArgentea" & err, 0)
                If String.Equals(szTranslatedError, "message  0 not found", StringComparison.InvariantCultureIgnoreCase) Then
                    LOG_Error(getLocationString(funcName), "Message does not exists:" & err & ". Use the original one.")
                    szTranslatedError = err
                End If

                ' not nice but we don't have a list of error codes
                TPMsgBox(PosDef.TARMessageTypes.TPERROR,
                             szTranslatedError,
                             0,
                             TheModCntr,
                             "LevelITCommonArgentea" & err)

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Sub

    Protected Overridable Sub ShowError(ByRef p As CommonParameters)
        Dim funcName As String = "ShowError"

        Try
            If Not p Is Nothing AndAlso Not String.IsNullOrEmpty(p.ErrorMessage) Then

                LOG_Debug(getLocationString(funcName), p.ErrorMessage)

                ShowError(p.Controller, p.ErrorMessage)

            End If

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
        End Try

    End Sub

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

    Protected Overridable Sub ShowWaitScreen(ByRef TheModCntr As ModCntr, ByVal bClear As Boolean, ByRef form As System.Windows.Forms.Form, Optional ByVal customMsg As String = "", Optional ByVal addCustomMsg As String = "")

        Dim i As Integer = -1
        Dim resolution As String = String.Empty

        Try
            LOG_FuncStart(getLocationString("ShowInfo"), "function started")

            If bClear Then

                If form IsNot Nothing Then
                    form.Close()
                    If Not form Is Nothing Then
                        If TheModCntr IsNot Nothing Then
                            TheModCntr.EndForm()
                        End If
                    End If
                    form = Nothing
                End If

            Else

                Dim msg As String = IIf(Not String.IsNullOrEmpty(customMsg), customMsg, getPosTxtNew((TheModCntr.contxt), "Message", TEXT_PLEASE_WAIT))
                msg &= customMsg

                form = TPMsg(msg, TEXT_PLEASE_WAIT, TheModCntr, "Message")
                Dim lx As Integer = (TheModCntr.GUICntr.ThePosForm.Width / 2) - (form.Width / 2)
                Dim ly As Integer = (TheModCntr.GUICntr.ThePosForm.Height / 2) - (form.Height / 2)
                form.Location = New System.Drawing.Point(lx, ly)

                'Try
                '    resolution = TheModCntr.GUICntr.ThePosForm.Width & "x" & TheModCntr.GUICntr.ThePosForm.Height
                '    i = Array.FindIndex(TheModCntr.GUICntr.POSGUIConfig.SubFormSizes, _
                '                                       Function(x As TPDotnet.Pos.SubFormSize) _
                '                                           x.Type = NO_STRETCH.ToString _
                '                                           AndAlso _
                '                                           x.Resolution = resolution)
                '    If i >= 0 Then




                '        form.BackgroundImageLayout = Windows.Forms.ImageLayout.Stretch

                '        form.SetBounds(TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.X, _
                '                       TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.Y, _
                '                       TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.Width, _
                '                       TheModCntr.GUICntr.POSGUIConfig.SubFormSizes(i).SubFormRectangle.Height)
                '    End If
                'Catch ex As Exception

                'End Try
                System.Windows.Forms.Application.DoEvents()

            End If

            Exit Sub

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("ShowInfo"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("ShowInfo"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("ShowInfo"), "end of function")
        End Try
    End Sub

#End Region


End Class
