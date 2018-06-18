Imports System
Imports System.Windows.Forms
Imports TPDotnet.Pos

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

Public Class ModSignOff
    Inherits TPDotnet.Pos.ModSignOff

    Public Overrides Function ModBase_run(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Short

        Dim myclsChangeLanguage As IclsChangeLanguage
        Dim szLanguage As String = ""

        ModBase_run = 0

        Try
            LOG_FuncStart(getLocationString("ModBase_run"))

            ' check emergency session
            ' =======================
            If m_bCheckSession Then
                ' do not check again, when calling from ModSignOff4Other, then it is an automatic declaration 
                ' of an emergency session. Check not needed again. Otherwise "never ending loop". Will be
                ' set in ModSignOff4Other to false to avoid always calling again and again.
                If CheckSession(taobj, TheModCntr) = 1 Then
                    ' an automatic signOff is done
                    ModBase_run = 1
                    Exit Function
                End If
            End If


            ' check recalled transactions, but not in silent mode
            ' ===================================================
            If TheModCntr.bExternalDialog = False AndAlso
               TheModCntr.bCalledFromWebService = False AndAlso
               TheModCntr.ModulNmbrExt = 0 Then
                If CheckRecallTX(taobj, TheModCntr) = False Then
                    Exit Function
                End If
            End If



            If DoStartWork(taobj, TheModCntr) = False Then
                Exit Function
            End If

            'Prepare the Form, perform dialog
            '================================
            If PerformDialog(taobj, TheModCntr) = False Then
                ' The Cancel-button was pressed, or an error occured
                ' Clear the changed fields in EmplMoveExt
                ' =======================================================
                CleanUp(taobj, TheModCntr)
                TheModCntr.GUICntr.ThePosForm.panelFunctionKeys.Visible = True
                LOG_Debug(getLocationString("ModBase_run"), "Dialog was aborted or failed --> exit function")
                Exit Function
            End If
            ' Open the drawer now
            If HandleDrawer(taobj, TheModCntr) = False Then
                ' an error occured
                ' Clear the changed fields in EmplMoveExt
                ' =======================================================
                CleanUp(taobj, TheModCntr)
                TheModCntr.GUICntr.ThePosForm.panelFunctionKeys.Visible = True
                LOG_Debug(getLocationString("ModBase_run"), "The drawer could not be opened properly --> exit function")
                Exit Function
            End If

            'everything went well
            'Counting were done
            'no differences detected or all differences were authorized
            'Withdraw or float was entered or taken automatically
            'Now do the work
            If DoEndWork(taobj, TheModCntr) = False Then
                LOG_Debug(getLocationString("ModBase_run"), "The end work could not be performed properly --> exit function")
                Exit Function
            End If

            If m_bIsMediaCount = False Then
                'we are in cashier declaration

                If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
                    ' Restore default GUI
                    TheModCntr.GUICntr.ChangeSkin("", False)
                End If

                ' change to default language ?
                ' ============================

                myclsChangeLanguage = createPosObject(Of IclsChangeLanguage)(TheModCntr, "clsChangeLanguage", 0)
                If Not myclsChangeLanguage Is Nothing Then
                    szLanguage = myclsChangeLanguage.GetDefaultLanguage(TheModCntr)
                    If TheModCntr.language <> szLanguage Then
                        myclsChangeLanguage.SetNewLanguage(taobj, TheModCntr, szLanguage)
                    End If
                    myclsChangeLanguage = Nothing
                End If
            End If

            If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then

                Dim szstring As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".posformcode.show_calculator").Trim
                If Not String.IsNullOrEmpty(szstring) Then
                    Dim showcalculator As Boolean = True
                    If szstring.ToUpper.Equals("y") Then
                        showcalculator = True
                    Else
                        showcalculator = False
                    End If
                    TheModCntr.GUICntr.ThePosForm.cmdCalculator.Visible = showcalculator
                End If

            End If

            If TheModCntr.ModulNmbrExt = 0 Then
                Try
                    EFTController.Instance.Close(taobj, TheModCntr)
                Catch ex As Exception
                    LOG_Debug(getLocationString("ModBase_run"), "argentea close function raises an error: " & ex.Message)
                End Try
            End If

            ModBase_run = 1

        Catch ex As Exception
            ModBase_run = 0
            LOG_Error(getLocationString("ModBase_run"), ex)
        Finally
            m_ModeObject = Nothing
            LOG_FuncExit(getLocationString("ModBase_run"), String.Concat("Function ModBase_run returns ", ModBase_run.ToString))
        End Try
    End Function

    Protected Overrides Function DoStartWork(ByVal taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean

        Dim MyTaSignOffRec As TPDotnet.Pos.TaSignOffRec
        Dim MyTaMediaCount As TaMediaCount
        Dim MyModTrainingStatistic As TPDotnet.Pos.ModTrainingStatistic
        Dim szSQL As String
        Dim szString As String
        Dim iLength As Integer
        Dim bRet As Boolean

        DoStartWork = True

        Try
            TheModCntr.SetFuncKeys((False))

            If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
                TheModCntr.GUICntr.ThePosForm.panelFunctionKeys.Hide()
                'Application.DoEvents()
            End If

            If TheModCntr.ModulNmbrExt <> 0 Then
                'we are in the function "MediaCount" = "Kassensturz"
                m_bIsMediaCount = True
            Else
                'we are in cashier declaration
                m_bIsMediaCount = False
            End If
            LOG_Debug(getLocationString("DoStartWork"), String.Concat("MediaCount = ", m_bIsMediaCount.ToString))

            ' Check the parameters
            ' ======================
            CheckParameters(TheModCntr)

            ' get the class for handling the mode
            ' ===================================
            m_ModeObject = GetModeObject(TheModCntr, m_szMode)
            If m_ModeObject Is Nothing Then
                DoStartWork = False
                LOG_Error(getLocationString("DoStartWork"), "Modul not present for mode (inherited of PosModeBase)")
                TheModCntr.GUICntr.ThePosForm.panelFunctionKeys.Visible = True
                Exit Function
            End If

            'Clear the details from old stati
            '================================
            ClearDetails(TheModCntr, 0, 0)

            'Clear the empty entries in MoveExt
            '================================
            RemoveEmptyEntriesFromMoveExt(taobj, TheModCntr)

            ' Read sign off information
            ' =========================
            m_EMPLMOVE = New ADODB_Recordset
            m_POSMOVE = New ADODB_Recordset

            szSQL = String.Concat("SELECT * FROM EmplMove WHERE lOperatorID = ", TheModCntr.lActOperatorID.ToString)
            LOG_Debug(getLocationString("DoStartWork"), String.Concat("open table: ", szSQL))
            m_EMPLMOVE.Open(szSQL, TheModCntr.con, ADODB_CursorTypeEnum.adOpenDynamic, ADODB_LockTypeEnum.adLockOptimistic)
            szSQL = "SELECT * FROM PosMove"
            LOG_Debug(getLocationString("DoStartWork"), String.Concat("open table: ", szSQL))
            m_POSMOVE.Open(szSQL, TheModCntr.con, ADODB_CursorTypeEnum.adOpenDynamic, ADODB_LockTypeEnum.adLockOptimistic)

            m_szSignOffName = m_EMPLMOVE.Fields_value("szEmplName")
            m_szSignOffNmbr = m_EMPLMOVE.Fields_value("szSignOnName")
            LOG_Debug(getLocationString("DoStartWork"), String.Concat("SignOffName =  ", m_szSignOffName))
            LOG_Debug(getLocationString("DoStartWork"), String.Concat("SignOffNmbr =  ", m_szSignOffNmbr))

            ' show "Closed" on customerdisplay during this function
            ' ====================================================
            szString = getPosTxtNew(TheModCntr.contxt, "Message", TEXT_INST_CLOSE)

            If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
                TheModCntr.GUICntr.ThePosForm.lblLineDisplay.Text = szString
                iLength = TheModCntr.OposLineDisplay_DeviceColumns * TheModCntr.OposLineDisplay_DeviceRows
                While szString.Length < iLength
                    szString = String.Concat(szString, " ")
                End While
                TheModCntr.OposShowLineDisplayString(szString)
            End If

            ' we have to get the pin from the EmplTable
            ' =========================================
            If m_bIsMediaCount = False Then
                ' check, sign on on Web service
                If CheckWebSignOn(taobj, TheModCntr) = False Then
                    ' do not continue
                    DoStartWork = False
                    LOG_Warning(1, getLocationString("DoStartWork"), "CheckWebSignOn failed. Do not continue declaration. -->exit function")
                    TheModCntr.GUICntr.ThePosForm.panelFunctionKeys.Visible = True
                    Exit Function
                End If

                MyTaSignOffRec = taobj.CreateTaObject(PosDef.TARecTypes.iTA_SIGN_OFF)
                bRet = MyTaSignOffRec.DBRead(TheModCntr.con, m_EMPLMOVE.Fields_value("szSignOnName").ToString.Trim)
                m_szSignOffPin = MyTaSignOffRec.EMPLinSignOff.szSignOnPassword

                If TheModCntr.bTrainingMode <> False Then
                    ' we are in training, update some statistic
                    MyModTrainingStatistic = createPosModelObject(Of ModTrainingStatistic)(TheModCntr, "ModTrainingStatistic", 0, False)
                    If Not MyModTrainingStatistic Is Nothing Then
                        If m_szMode = POS_MODE Then
                            ' need statistic for all up to now logged on operator
                            MyModTrainingStatistic.FillTrainingStatistic_NEW(TheModCntr, taobj)
                        Else
                            ' only the statistic of the current operator
                            MyModTrainingStatistic.FillTrainingStatistic_NEW(TheModCntr, taobj, True)
                        End If
                    End If
                End If
            Else
                'we are in MediaCount
                MyTaMediaCount = taobj.CreateTaObject(PosDef.TARecTypes.iTA_MEDIA_COUNT)
                bRet = MyTaMediaCount.DBRead(TheModCntr.con, m_EMPLMOVE.Fields_value("szSignOnName").ToString.Trim)
                m_szSignOffPin = MyTaMediaCount.EMPLinMediaCount.szSignOnPassword
                LOG_Debug(getLocationString("DoStartWork"), "MyTaMediaCount created")
            End If

        Catch ex As Exception
            DoStartWork = False
            LOG_Error(getLocationString("DoStartWork"), ex)
        Finally
            MyTaSignOffRec = Nothing
            MyTaMediaCount = Nothing
            MyModTrainingStatistic = Nothing
            LOG_FuncExit(getLocationString("DoStartWork"), String.Concat("Function DoStartWork returns ", DoStartWork.ToString))
        End Try
    End Function

    Protected Overrides Function PerformDialog(ByRef taobj As TPDotnet.Pos.TA, ByRef theModCntr As ModCntr) As Short

        If theModCntr.bCalledFromWebService OrElse theModCntr.bExternalDialog Then
            Return True
        End If

        Return MyBase.PerformDialog(taobj, theModCntr)

    End Function

    Protected Overrides Function DoEndWork(ByVal taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean

        Dim MyModPrintFloat As TPDotnet.Pos.ModPrintFloat
        Dim bRet As Boolean

        DoEndWork = False

        Try
            If m_bIsMediaCount = False Then
#If Not CF_FRAMEWORK Then
                ' do not check the drawer until the user signed on again
                TheModCntr.OposDrawer_CheckforOpenFlag = False
#End If
            End If

            If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
                ' Unload signoff-form
                ' ===================================
                m_FormSignOff.Close()
                Application.DoEvents()

                TheModCntr.SubForm2 = Nothing
                TheModCntr.Dialog2FormName = ""
                TheModCntr.Dialog2Activ = False
                TheModCntr.EndForm()
            End If

            If m_bIsMediaCount = False Then
                'we are in cashier declaration
                'add SignOff-object to transaction
                '=================================
                AddSignOffObject(taobj, TheModCntr)
            Else
                'we are in "Media Count" = "Kassensturz"
                'add MediaCount-object to transaction
                '===========================================
                AddMediaCountObject(taobj, TheModCntr)
            End If

            ' for each lMediaNmbr in EmplMoveExt or PosMoveExt, we create one Record
            ' and add it to ta, unless it doesn't already exist
            '========================================================================
            AddEmplNewMoveObjects(taobj, TheModCntr)

            If m_bIsMediaCount <> False Then
                ' we are in MediaCounting
                ' Reset the changed fields in EmplMoveExt and EmplMoveDetail
                ' ==========================================================
                ResetLocalDB(taobj, TheModCntr)
            End If

            'close the local Empl-tables
            '======================================
            CloseRecSet(m_EMPLMOVEEXT)
            CloseRecSet(m_EMPLMOVEDETAIL)

            ' now create the footerentry in the ta
            '=====================================
            taobj.TAEnd(fillFooterLines((TheModCntr.con), taobj, TheModCntr))

            If m_bIsMediaCount = False Then
                TheModCntr.szSignOnName = ""
                TheModCntr.szPrintCode = ""
                TheModCntr.szActEmployeeName = ""
                TheModCntr.lActOperatorID = 0
            End If

            'everything with the transaction went fine so far
            '================================================
            DoEndWork = True
            taobj.bPrintReceipt = True
            taobj.bDelete = True ' ok , we will delete this TA
            taobj.bTAtoFile = True ' now we will write the ta to the tafile

            If m_bIsMediaCount <> False Then
                'we are in media count, increase counter for media count
                '=======================================================
                IncreaseMediaCounter()
            Else
                'we are in cashier declaration
                'we must export the EmplMove, EmplMoveExt and delete all entries from DB
                '=======================================================================

                ' sync Return Web Service
                ' =======================
                ResyncReturnWebService(TheModCntr)

                ' sync CVM Web Service
                ' =======================
                ResyncCVMWebService(TheModCntr)

                ' close all
                CloseRecSet(m_POSMOVE)
                CloseRecSet(m_EMPLMOVE)
                CloseRecSet(m_EMPLMOVEEXT)
                CloseRecSet(m_EMPLMOVEDETAIL)

                ' do the sign off corresponding to the mode
                '===============================================
                bRet = m_ModeObject.DoSignOff(TheModCntr, taobj)

                TheModCntr.bEmplSignedOnOffline = False

                ' set status "SignOff" in ComputerComponentStatus
                ' ===============================================
                UpdateComponentStatus(taobj, TheModCntr, Me.GetType.Name(), STATUS_SIGN_OFF)

            End If

            TheModCntr.SetFuncKeys((True))

            If Not TheModCntr.bExternalDialog AndAlso Not TheModCntr.bCalledFromWebService Then
                'show the Calculator-Button again
                TheModCntr.GUICntr.ThePosForm.cmdCalculator.Visible = True
            End If

            ' in training write a valid Ta with training statistics
            ' ================================================================
            If TheModCntr.bTrainingMode <> False And m_bIsMediaCount = False Then
                TheModCntr.ModulAgain = "ModTrainingStatistic"
            End If

            'Create the Object of the ModPrintFloat
            'to print an additional receipt for the float only in case of declaration
            '========================================================================
            If m_bIsMediaCount = False Then
                MyModPrintFloat = createPosModelObject(Of ModPrintFloat)(TheModCntr, "ModPrintFloat", 0, False)
                If MyModPrintFloat Is Nothing Then
                    LOG_Debug(getLocationString("DoEndWork"), "createPosModelObject ModPrintFloat failed")
                Else
                    If MyModPrintFloat.ModBase_run(taobj, TheModCntr) <> 1 Then
                        'TODO something went wrong
                    End If
                End If
            End If

        Catch ex As Exception
            DoEndWork = False
            LOG_Error(getLocationString("DoEndWork"), ex)
        Finally
            MyModPrintFloat = Nothing
            LOG_FuncExit(getLocationString("DoEndWork"), String.Concat("Function DoEndWork returns ", DoEndWork.ToString))
        End Try
    End Function


End Class
