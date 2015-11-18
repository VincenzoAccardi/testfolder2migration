Option Strict Off
Option Explicit On
Imports System
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Imports System.Collections.Generic
Imports System.Xml.Linq
Imports System.Xml.XPath

Public Class TA : Inherits TPDotnet.Pos.TA : Implements TPDotnet.IT.Common.Pos.IFiscalTA

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TA
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Private Class fields"
    Protected m_bFiscalPrinted As Boolean
#End Region

#Region "Properties"

    ''' <summary>
    ''' gets / sets the fiscal tickt is printed flag.
    ''' When the flag is set, we allow only non fiscal ticket printing (eg. second receipt)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property bFiscalPrinted() As Boolean Implements IFiscalTA.bFiscalPrinted
        Get
            bFiscalPrinted = m_bFiscalPrinted
        End Get
        Set(ByVal Value As Boolean)
            m_bFiscalPrinted = Value
        End Set
    End Property

#End Region

#Region "Overrides"

    Public Overrides Sub AddObject(ByRef obj As TPDotnet.Pos.TaBaseRec, iPos As Short)
        Dim funcName As String = "AddObject"
        Dim parameters As Dictionary(Of String, Object)

        Try
            parameters = New Dictionary(Of String, Object) From _
                { _
                    {"Transaction", Me} _
                  , {"Record", obj} _
                  , {"Index", iPos} _
                }
            If TPDotnet.IT.Common.Pos.TransactionRecordPlugIn.Instance.Add(parameters) <> ITransactionRecordPlugInReturnCode.KO Then
                MyBase.AddObject(obj, iPos)
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Protected Overrides Sub RemoveObject(ByRef i As Short)
        Dim funcName As String = "AddObject"
        Dim parameters As Dictionary(Of String, Object)

        Try
            parameters = New Dictionary(Of String, Object) From _
                { _
                    {"Transaction", Me} _
                  , {"Index", i} _
                }
            If TPDotnet.IT.Common.Pos.TransactionRecordPlugIn.Instance.Remove(parameters) <> ITransactionRecordPlugInReturnCode.KO Then
                MyBase.RemoveObject(i)
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Protected Overrides Sub Class_Initialize_Renamed()
        MyBase.Class_Initialize_Renamed()
        m_bFiscalPrinted = False
    End Sub

    Public Overrides Function TATmpValues2NodeControl() As String

        ' in this function we fill the TA_CONTROL Node with the temporary TA values
        TATmpValues2NodeControl = ""

        Try

            TATmpValues2NodeControl = MyBase.TATmpValues2NodeControl()
            If TATmpValues2NodeControl <> "" Then

                TATmpValues2NodeControl = TATmpValues2NodeControl.Replace("</TA_CONTROL>" & vbCrLf, "")
                If m_bFiscalPrinted = False Then
                    TATmpValues2NodeControl = TATmpValues2NodeControl & "<bFiscalPrinted>0</bFiscalPrinted>" & vbCrLf
                Else
                    TATmpValues2NodeControl = TATmpValues2NodeControl & "<bFiscalPrinted>1</bFiscalPrinted>" & vbCrLf
                End If

            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("TATmpValues2NodeControl"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("TATmpValues2NodeControl"), InnerEx)
            End Try
        Finally
            Try
                If TATmpValues2NodeControl.Length > 0 Then
                    TATmpValues2NodeControl = TATmpValues2NodeControl & "</TA_CONTROL>" & vbCrLf
                End If
                LOG_FuncExit(getLocationString("TATmpValues2NodeControl"), "Function TATmpValues2NodeControl returns " & TATmpValues2NodeControl.ToString)
            Catch innerEx As Exception
                LOG_Error(getLocationString("TATmpValues2NodeControl (Finally)"), innerEx)
            End Try
        End Try

    End Function
    Public Overrides Function TAfromControlNode2TA(ByRef XMLTARecordNode As System.Xml.Linq.XElement) As Boolean

        Dim myXMLValueNode As System.Xml.Linq.XElement
        Dim szTmpValue As String

        ' in this function we fill TA from the TA_CONTROL node        
        Try

            TAfromControlNode2TA = False
            TAfromControlNode2TA = MyBase.TAfromControlNode2TA(XMLTARecordNode)
            If TAfromControlNode2TA Then

                For Each myXMLValueNode In XMLTARecordNode.Descendants()

                    szTmpValue = myXMLValueNode.Name.LocalName

                    If szTmpValue = "m_bFiscalPrinted" Then
                        m_bFiscalPrinted = CBool(myXMLValueNode.Value)
                    End If

                Next myXMLValueNode

                TAfromControlNode2TA = True
            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("TAfromControlNode2TA"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("TAfromControlNode2TA"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("TAfromControlNode2TA"), "Function TAfromControlNode2TA returns " & TAfromControlNode2TA.ToString)
        End Try

    End Function

    Public Overrides Sub TAEnd(ByRef taftrrecobj As TPDotnet.Pos.TaFtrRec)

        MyBase.TAEnd(taftrrecobj)

        If TypeOf taftrrecobj Is TPDotnet.IT.Common.Pos.TaFtrRec Then

            Dim MyTaFtrRec As TPDotnet.IT.Common.Pos.TaFtrRec
            MyTaFtrRec = taftrrecobj
            MyTaFtrRec.lSoldItems = Me.GetQuantity

        End If

    End Sub

    ' Ema 06-04-09 : reset the transaction number by date.
    Public Overrides Sub NewTa2Reg()

        Dim lRolloverTrans As Integer
        Dim lActTanmbr As Integer       ' store the ta, which were working with
        Dim lNewActTanmbr As Integer       ' store the ta, which were working with
        Dim szITTransactionResetDate As String = "29991231"
        Dim TransactionResetDate As Date = New Date(2999, 12, 31)
        Dim bTransactionResetDate As Boolean = False

        ' write back TANmbr to registry
        Try
            ' already increased
            ' =================
            If m_bUpdateTaNmbr <> False Then
                LOG_Warning(99, getLocationString("NewTa2Reg"), "already updated for ta:" + m_lactTaNmbr.ToString)
                Exit Sub
            End If

            ' store it for later compare
            lActTanmbr = m_lactTaNmbr       ' store the ta, which were working with

            GetLastTransactionNmbr(m_lactTaNmbr, lRolloverTrans)
            GetTransactionResetDate(szITTransactionResetDate)

            ' check ta number, whether correct
            If m_lactTaNmbr <> lActTanmbr Then
                ' ta number differs!!!!
                ' bring error message
                LOG_Error(getLocationString("NewTa2Reg"), _
                          "Transaction numbers differ, m_lactTaNmbr=" & m_lactTaNmbr.ToString & _
                          " and lActTanmbr=" & lActTanmbr.ToString)
            End If

            If Date.TryParseExact(szITTransactionResetDate, "yyyyMMdd", Nothing, Globalization.DateTimeStyles.None, TransactionResetDate) AndAlso _
                Date.Compare(TransactionResetDate, Date.Now) < 0 Then
                ' transaction reset date reached
                bTransactionResetDate = True
            End If

            ' we will check now for overflow
            If m_lactTaNmbr >= lRolloverTrans OrElse bTransactionResetDate Then
                lNewActTanmbr = 1
            Else
                lNewActTanmbr = m_lactTaNmbr + 1
            End If

            SetLastTransactionNmbr(lNewActTanmbr, lRolloverTrans)
            If bTransactionResetDate Then
                TransactionResetDate.AddYears(1) ' increase the year
                SetTransactionResetDate(Format(TransactionResetDate, "yyyyMMdd"))
            End If

            ' update done
            m_bUpdateTaNmbr = True

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("NewTa2Reg"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("NewTa2Reg"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("NewTa2Reg"), "")
        End Try
    End Sub

    Public Overrides Sub Clone(ByRef ta As TPDotnet.Pos.TA)

        Try

            MyBase.Clone(ta)
            Me.bFiscalPrinted = DirectCast(ta, TPDotnet.IT.Common.Pos.IFiscalTA).bFiscalPrinted

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("Clone"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("Clone"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("Clone"), "")
        End Try

    End Sub

    Public Overrides Function GetFiscalReturns() As Decimal
        Dim TaBase As TPDotnet.Pos.TaBaseRec
        Dim MyTaArtReturnRec As TaArtReturnRec
        Dim i As Short

        '=================================================================
        ' get the amount of fiscal returns
        '=================================================================
        Try
            LOG_FuncStart(getLocationString("GetFiscalReturns"))
            GetFiscalReturns = 0

            For i = TARecords.Count() To 1 Step -1
                TaBase = TARecords.Item(i)
                If TaBase.theHdr.bIsVoided <> TAdefine.TaAllHdrTypes.IS_VOIDED AndAlso TaBase.sid = PosDef.TARecTypes.iTA_ART_RETURN Then
                    MyTaArtReturnRec = TARecords.Item(i)
                    If MyTaArtReturnRec.ARTinArtReturn.bIsFiscalArt <> 0 Then
                        GetFiscalReturns = GetFiscalReturns + (MyTaArtReturnRec.dTaTotal + (Math.Abs(MyTaArtReturnRec.dTaQty) * MyTaArtReturnRec.dTaDiscount))
                    End If
                    MyTaArtReturnRec = Nothing
                End If
                TaBase = Nothing
            Next

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetFiscalReturns"), ex)
                GetFiscalReturns = 0

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetFiscalReturns"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("GetFiscalReturns"), "Function GetFiscalReturns returns " & GetFiscalReturns.ToString)
        End Try
    End Function

    Public Overridable Function GetPayedValueForMediaMember(ByVal lMediaMember As Integer) As Decimal Implements IFiscalTA.GetPayedValueForMediaMember

        Dim MyTaMediaRec As TaMediaRec

        Try

            GetPayedValueForMediaMember = 0
            For Each TARecord In TARecords
                If TARecord.sid = PosDef.TARecTypes.iTA_MEDIA Then
                    MyTaMediaRec = TARecord
                    If lMediaMember = MyTaMediaRec.PAYMENTinMedia.lMediaMember Then
                        GetPayedValueForMediaMember = GetPayedValueForMediaMember + TARecord.GetPayedValue()
                    End If
                End If
            Next TARecord

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetPayedValueForMediaMember"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetPayedValueForMediaMember"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("GetPayedValueForMediaMember"), "Function GetPayedValueForMediaMember returns " & GetPayedValueForMediaMember.ToString)
        End Try

    End Function

    Public Overridable Function GetPaidValueOfMediaMemberToCheckCashHalo() As Decimal Implements IFiscalTA.GetPaidValueOfMediaMemberToCheckCashHalo

        Dim MyTaMediaRec As TaMediaRec

        Try

            GetPaidValueOfMediaMemberToCheckCashHalo = 0
            For Each TARecord In TARecords
                If TARecord.sid = TPDotnet.Pos.PosDef.TARecTypes.iTA_MEDIA Then
                    MyTaMediaRec = TARecord
                    If CType(MyTaMediaRec.PAYMENTinMedia, TPDotnet.IT.Common.Pos.IFiscalPAYMENT).bITCheckCashHalo Then
                        GetPaidValueOfMediaMemberToCheckCashHalo = GetPaidValueOfMediaMemberToCheckCashHalo + TARecord.GetPayedValue()
                    End If
                End If
            Next TARecord

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetPaidValueOfMediaMemberToCheckCashHalo"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetPaidValueOfMediaMemberToCheckCashHalo"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("GetPaidValueOfMediaMemberToCheckCashHalo"), "Function GetPaidValueOfMediaMemberToCheckCashHalo returns " & GetPaidValueOfMediaMemberToCheckCashHalo.ToString)
        End Try

    End Function


    Public Overrides Function TAfromFile(ByRef szFileName As String, _
                                           ByRef bReadHdr As Boolean, _
                                           ByRef bReadStatistik As Boolean, _
                                           ByRef bCheckTraining As Boolean, _
                                           ByRef bFileIsStream As Boolean, _
                                           ByRef bReadTA_Control As Boolean, _
                                           ByRef bRefresh As Boolean, _
                                           ByRef bCheckCRC As Boolean, _
                                        ByVal bCheckRescan As Boolean) As Boolean



        TAfromFile = False
        Try

            Dim xmlDoc As System.Xml.Linq.XDocument
            Dim myXMLRootNode As System.Xml.Linq.XElement = Nothing
            Dim myXMLNewTANode As System.Xml.Linq.XElement = Nothing
            Dim myXMLValueNode As System.Xml.Linq.XElement
            Dim myControlNode As System.Xml.Linq.XElement
            Dim bRet As Boolean

            Dim myXmlTaBase As TaBaseRec

            ' for not within the standard defined TA Objects
            Dim szTmpValue As String

            Dim iPriceWithVat As Integer
            Dim bCreateNewTax As Boolean = False
            Dim bAddNewTax As Boolean
            Dim bTaxFreeCust As Boolean = False

            Dim sNmbrRecs As Short

            Dim szSaveFileName As String = String.Empty

            TAfromFile = False

            Try

                ' error code
                m_sErrTAfromFile = 0

                ' get up to now stored objects within the TA
                ' ==========================================
                sNmbrRecs = Me.GetNmbrofRecs


                ' ok lets create a DOM Document
                xmlDoc = New System.Xml.Linq.XDocument()

                ' store the fileName, because will be probably overwritten in ReadTA
                ' ==================================================================
                If bFileIsStream = False Then
                    szSaveFileName = szFileName
                End If

                ' read file either from directory or from DBShare
                ' ===============================================
                ReadTA(szFileName, bFileIsStream, xmlDoc)
                ' error ?
                If m_sErrTAfromFile <> 0 Then
                    ' error appears
                    If String.IsNullOrEmpty(szSaveFileName) Then
                        LOG_Error(getLocationString("TAfromFileOnlyInfo"), String.Concat("could not read: ", szFileName))
                    Else
                        LOG_Error(getLocationString("TAfromFileOnlyInfo"), String.Concat("could not read: ", szSaveFileName))
                    End If
                    Exit Function
                End If


                ' get node of TAS and NEW_TA
                ' ==========================
                GetStartXDoc(xmlDoc, myXMLRootNode, myXMLNewTANode)

                ' for older TAs we assume, prices are incl. VAT
                iPriceWithVat = True

                ' now we read all the childnotes
                For Each myXMLValueNode In myXMLNewTANode.Elements()

                    ' here we reload the control types
                    ' --------------------------------
                    szTmpValue = myXMLValueNode.Name.LocalName

                    If szTmpValue.ToUpper(CultureInfo_EnUs) = _
                         TPDotnet.Services.CRCService.CRCService.CRC_SERVICE_TAG_NAME.ToUpper(CultureInfo_EnUs) Then
                        ' checksum, no object, only for checking
                        ' therefore read the next one
                        Continue For
                    End If

                    If szTmpValue = "TA_CONTROL" Then
                        myControlNode = myXMLValueNode

                        If bReadTA_Control = True Then
                            bRet = TAfromControlNode2TA(myXMLValueNode)
                        End If

                        'For TPiSHOP the bMustRescan flag has to be checked in some cases.
                        If bCheckRescan Then
                            '1st read the bMustRescan flag from XML
                            TAbMustRescanFromControlNode2TA(myXMLValueNode)
                            If MustRescan Then
                                'If rescan flag is set we exit the function and return false.
                                LOG_Debug(getLocationString("TAfromFile"), "Transaction must be rescanned - will not be loaded.")
                                Exit Function
                            End If
                        End If

                        ' in some older versions we store the information about price with / without tax in 
                        ' the control object, now try to get this info
                        iPriceWithVat = Me.GetPricesWithVat(myXMLValueNode)
                        Continue For
                    End If

                    myXmlTaBase = CreateTaObject(Of TaBaseRec)(szTmpValue)

                    If Not myXmlTaBase Is Nothing Then
                        Dim myTaBase As TPDotnet.Pos.TaBaseRec = myXmlTaBase
                        If Not bReadHdr AndAlso myXmlTaBase.sid = TPDotnet.Pos.TARecTypes.iTA_HEADER Then
                            myTaBase = Me.GetTALine(1)
                            If Not myTaBase.DataFieldObject.ExistField("lOriginalOperatorID") Then
                                myTaBase.DataFieldObject.AddField("lOriginalOperatorID", TPDotnet.Pos.DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)
                            End If
                            Dim doc As XDocument = XDocument.Parse(myXMLValueNode.ToString(), LoadOptions.None)
                            Dim xElement As XElement = doc.XPathSelectElement("//lOperatorID")
                            myTaBase.DataFieldObject.SetPropertybyName("lOriginalOperatorID", CInt(xElement.Value))
                        End If
                        Dim xDoc As XDocument = XDocument.Parse(myXMLValueNode.ToString(), LoadOptions.None)
                        Dim xElements As List(Of XElement) = xDoc.XPathSelectElements("//*[@Type]").ToList
                        For Each xEl As XElement In xElements 
                            If Not myTaBase.DataFieldObject.ExistField(xEl.Name.ToString()) Then
                                myTaBase.DataFieldObject.AddField(xEl.Name.ToString(), CType(xEl.FirstAttribute.Value.ToString(), TPDotnet.Pos.DataField.FIELD_TYPES))
                            End If
                            myTaBase.DataFieldObject.SetPropertybyName(xEl.Name.ToString(), xEl.Value.ToString())
                        Next

                        If m_bCheckTAObj Then
                            ' check, which objects should be added to the current TA
                            bRet = checkTAObj(szFileName, myXMLValueNode, szTmpValue, myXmlTaBase, bReadHdr, bReadStatistik, bCheckTraining, _
                                       bFileIsStream, bRefresh, bCheckCRC, iPriceWithVat, bCreateNewTax, bAddNewTax, bTaxFreeCust)
                            If bRet = False Then
                                Exit Function
                            End If
                        Else
                            ' add every object without any check
                            bRet = myXmlTaBase.FillFromXmlNode(myXMLValueNode)
                            Me.Add(myXmlTaBase, , bRefresh)
                        End If
                    Else
                        ' try to create this unknown object
                        CreateUnknownTAObj(myXMLValueNode, szTmpValue, bCreateNewTax, bAddNewTax, bTaxFreeCust, bRefresh)
                    End If

                Next myXMLValueNode

                ' correct ta: check, whether creation number is unique due to existing and now loaded objects
                '             header on pos 1 and statistic on 2
                ' ===========================================================================================
                CorrectLoadedTA(bReadHdr, bReadStatistik, sNmbrRecs)

                TAfromFile = True
                Exit Function

            Catch ex As Exception
                Try
                    LOG_Error(getLocationString("TAfromFile"), ex)

                Catch InnerEx As Exception
                    LOG_ErrorInTry(getLocationString("TAfromFile"), InnerEx)
                End Try
            Finally
                Try
                    If Not String.IsNullOrEmpty(szSaveFileName) Then
                        ' restore the real
                        szFileName = szSaveFileName
                        bFileIsStream = False
                    End If
                Catch ex As Exception
                    LOG_Error(getLocationString("TAfromFile (Finally)"), ex)
                End Try
                LOG_FuncExit(getLocationString("TAfromFile"), String.Concat("Function TAfromFile returns ", TAfromFile.ToString))
            End Try

            


        Catch ex As Exception
            Try
                LOG_Error(getLocationString("TAfromFile"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("TAfromFile"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("TAfromFile"), String.Concat("Function TAfromFile returns ", TAfromFile.ToString))
        End Try

    End Function
    'Public Overrides Sub TAStart(ByRef contxt As ADODB_Connection, _
    '                               ByRef conTPPosDB As ADODB_Connection) 
    '    Dim MyTaHdrRec As TaHdrRec
    '    Dim szWorkstationGroupID As String = ""
    '    Dim szWorkstationID As String = ""
    '    Dim lWorkstationNmbr As String = ""
    '    Dim szDate As String = ""

    '    Try
    '        MyBase.TAStart(contxt, conTPPosDB)

    '        szWorkstationGroupID = getTransactionPastWorkstationGroup()
    '        szWorkstationID = getTransactionPastWorkstationID()
    '        lWorkstationNmbr = getTransactionPastWorkstationNumber()
    '        szDate = getTransactionPastDate()

    '        ' check if registry has been filled to make a past transaction
    '        If Not String.IsNullOrEmpty(szWorkstationGroupID) AndAlso _
    '           Not String.IsNullOrEmpty(szWorkstationID) AndAlso _
    '           Not String.IsNullOrEmpty(lWorkstationNmbr) AndAlso _
    '           Not String.IsNullOrEmpty(szDate) _
    '        Then
    '            m_szStartTaTime = szDate
    '            MyTaHdrRec = GetTALine(1)
    '            MyTaHdrRec.szDate = m_szStartTaTime
    '        End If

    '    Catch ex As Exception
    '        Try
    '            LOG_Error(getLocationString("TAStart"), ex)
    '        Catch InnerEx As Exception
    '            LOG_ErrorInTry(getLocationString("TAStart"), InnerEx)
    '        End Try
    '    Finally
    '        LOG_FuncExit(getLocationString("TAStart"), "Exit")
    '    End Try

    'End Sub

    Public Sub SplitArticleIfNeededForTotalDiscountRounding(ByRef DiscInfo As TPDotnet.Pos.TaDiscInfoRec, ByRef MyTaArtSaleRec As TPDotnet.Pos.TaArtSaleRec, ByRef MyTaDiscInfoRec As TPDotnet.Pos.TaDiscInfoRec, ByRef dOriginalVal As Decimal, ByRef dVal As Decimal)

        Dim lNewPosition As Integer
        Dim MyOneTaArtSaleRec As TPDotnet.Pos.TaArtSaleRec
        Dim MyOneTaDiscInfoRec As TPDotnet.Pos.TaDiscInfoRec
        Dim dOriginalQty As Decimal

        Try
            LOG_FuncStart(getLocationString("SplitArticleIfNeededForTotalDiscountRounding"))

            MyTaArtSaleRec.dDiscQty = MyTaArtSaleRec.dTaQty
            dOriginalQty = MyTaArtSaleRec.dTaQty

            MyOneTaArtSaleRec = Me.CreateTaObject(Of TPDotnet.Pos.TaArtSaleRec)(TPDotnet.Pos.PosDef.TARecTypes.iTA_ART_SALE)
            MyTaArtSaleRec.ClonetoArtSale(MyOneTaArtSaleRec)
            MyOneTaArtSaleRec.dTaQty = 1
            MyOneTaArtSaleRec.dDiscQty = 1
            MyOneTaArtSaleRec.ActualizeQuantity(1, 1, 1, 1, m_iExactness)
            MyOneTaArtSaleRec.dPartOfTotalDiscount = TPDotnet.Services.Rounding.Rounding.dRounding(MyTaArtSaleRec.dPartOfTotalDiscount / dOriginalQty, TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, 1, m_iExactness)

            MyTaArtSaleRec.dPartOfTotalDiscount = TPDotnet.Services.Rounding.Rounding.dRounding((MyTaArtSaleRec.dPartOfTotalDiscount / dOriginalQty) * (dOriginalQty - 1), TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, 1, m_iExactness)
            MyTaArtSaleRec.dTaQty -= 1
            MyTaArtSaleRec.dDiscQty -= 1
            MyTaArtSaleRec.ActualizeQuantity(MyTaArtSaleRec.dTaQty, dModMeasurementEntry1, dModMeasurementEntry2, dModMeasurementEntry3, m_iExactness)

            MyOneTaArtSaleRec.dPartOfTotalDiscount += dOriginalVal - dVal
            MyOneTaArtSaleRec.dTaDiscount += dOriginalVal - dVal
            Add(MyOneTaArtSaleRec, GetPositionFromCreationNmbr(MyTaDiscInfoRec.theHdr.lTaCreateNmbr))

            MyOneTaDiscInfoRec = Me.CreateTaObject(Of TPDotnet.Pos.TaDiscInfoRec)(TPDotnet.Pos.PosDef.TARecTypes.iTA_DISC_INFO)
            MyTaDiscInfoRec.ClonetoDiscInfoRec(MyOneTaDiscInfoRec)
            MyOneTaDiscInfoRec.lManualDiscountID = MyTaDiscInfoRec.lManualDiscountID
            MyOneTaDiscInfoRec.lArtRefNmbr = MyOneTaArtSaleRec.theHdr.lTaCreateNmbr
            MyOneTaDiscInfoRec.theHdr.lTaRefToCreateNmbr = MyOneTaArtSaleRec.theHdr.lTaCreateNmbr
            MyOneTaDiscInfoRec.dTotalDiscount = TPDotnet.Services.Rounding.Rounding.dRounding(MyOneTaDiscInfoRec.dTotalDiscount / dOriginalQty, TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, 1, m_iExactness)

            MyTaDiscInfoRec.dTotalDiscount = TPDotnet.Services.Rounding.Rounding.dRounding((MyTaDiscInfoRec.dTotalDiscount / dOriginalQty) * (dOriginalQty - 1), TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, 1, m_iExactness)

            MyOneTaDiscInfoRec.dTotalDiscount += dOriginalVal - dVal
            Add(MyOneTaDiscInfoRec, GetPositionFromCreationNmbr(MyOneTaArtSaleRec.theHdr.lTaCreateNmbr))

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("SplitArticleIfNeededForTotalDiscountRounding"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("SplitArticleIfNeededForTotalDiscountRounding"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("SplitArticleIfNeededForTotalDiscountRounding"), "end")
        End Try

    End Sub

    Public Overrides Sub AddTotalDiscInfo(ByRef DiscInfo As TPDotnet.Pos.TaDiscInfoRec)
        Dim MyTaDiscInfoRec As TPDotnet.Pos.TaDiscInfoRec
        Dim MyTaArtSaleRec As TPDotnet.Pos.TaArtSaleRec
        Dim MyTaArtVSetRec As TPDotnet.Pos.TaArtVSetRec
        Dim TaBase As TPDotnet.Pos.TaBaseRec
        Dim i As Short
        Dim dVal As Decimal
        Dim dBasic As Decimal
        Dim dDisc As Decimal
        Dim lHelp As Integer
        Dim lRef As Integer
        Dim bDummy As Boolean

        Dim dOriginalVal As Decimal = 0
        Dim lTaCreateNmbrArtSaleWithHighAmount As Integer = -1
        Dim lTaCreateNmbrDiscInfoWithHighAmount As Integer = -1
        Dim dItemWithHighAmount As Decimal = 0
        Dim bItemWithHighAmountSavedHasMultipleQty As Boolean = False
        Dim bItemWithHighAmountHasBeebSaved As Boolean = False

        'Set discount into the article sales records
        '============================================
        Try
            LOG_FuncStart(getLocationString("SetTotalDiscInfo"))
            dVal = 0

            If DiscInfo.lDiscListType <> PosDef.DiscountTypes.LIST_TOTAL_DISCOUNT_IN_PERC AndAlso _
                DiscInfo.lDiscListType <> PosDef.DiscountTypes.LIST_TOTAL_DISCOUNT_IN_AMOUNT Then
                Exit Sub
            End If

            dOriginalVal = TPDotnet.Services.Rounding.Rounding.dRounding(DiscInfo.dTotalAmount * DiscInfo.dDiscValue / 100, _
                                         TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, _
                                         1, m_iExactness) * (-1)

            Me.Add(DiscInfo)
            TaBase = DiscInfo
            lRef = TaBase.theHdr.lTaCreateNmbr
            TaBase = Nothing

            ' we create TaDiscInfo entrys for all articles getting total discount
            '====================================================================
            For i = TARecords.Count() To 1 Step -1
                If CType(TARecords.Item(i), TaBaseRec).sid = PosDef.TARecTypes.iTA_ART_SALE AndAlso _
                    CType(TARecords.Item(i), TaBaseRec).theHdr.bIsVoided <> TAdefine.TaAllHdrTypes.IS_VOIDED Then

                    MyTaArtSaleRec = TARecords.Item(i)

                    If (MyTaArtSaleRec.ARTinArtSale.bTotalDiscountAllowed <> 0) Then

                        ' subtract previously given discount from basic total amount
                        dBasic = MyTaArtSaleRec.dTaTotal + (MyTaArtSaleRec.dTaDiscount * MyTaArtSaleRec.dTaQty)

                        ' calculate the whole amount of discount for this item
                        dDisc = ((DiscInfo.dDiscValue * dBasic) / 100) * (-1)

                        ' round discount 
                        dDisc = TPDotnet.Services.Rounding.Rounding.dRounding(dDisc / MyTaArtSaleRec.dTaQty, _
                                         TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, _
                                         1, m_iExactness)
                        dDisc = TPDotnet.Services.Rounding.Rounding.dRounding(dDisc * MyTaArtSaleRec.dTaQty, TPDotnet.Services.Rounding.ROUNDINGMETHOD.ROUND_ARITHMETIC, 1, m_iExactness)

                        'accumulate for total disc info line
                        dVal = dVal + dDisc

                        'write disc info behind each article
                        If DiscInfo.lIsArtorDept <> isTotal Then
                            MyTaDiscInfoRec = CreateTaObject(PosDef.TARecTypes.iTA_DISC_INFO)
                            TaBase = MyTaArtSaleRec
                            lHelp = TaBase.theHdr.lTaCreateNmbr
                            TaBase = MyTaDiscInfoRec
                            TaBase.theHdr.lTaRefToCreateNmbr = lHelp
                            TaBase = Nothing

                            ' store the basic and the whole amount of discount
                            MyTaDiscInfoRec.dTotalAmount = dBasic
                            MyTaDiscInfoRec.dTotalDiscount = dDisc

                            ' setup all other information from passed in discount info record
                            MyTaDiscInfoRec.dDiscValue = DiscInfo.dDiscValue
                            MyTaDiscInfoRec.lDiscExtNmbr = DiscInfo.lDiscExtNmbr
                            MyTaDiscInfoRec.lDiscGroupNmbr = DiscInfo.lDiscGroupNmbr
                            MyTaDiscInfoRec.lDiscListType = PosDef.DiscountTypes.iTOTAL_DISCOUNT
                            MyTaDiscInfoRec.lIsArtorDept = DiscInfo.lIsArtorDept
                            MyTaDiscInfoRec.szArtDptNmbr = MyTaArtSaleRec.ARTinArtSale.szPOSItemID
                            MyTaDiscInfoRec.szDiscDesc = DiscInfo.szDiscDesc
                            MyTaDiscInfoRec.szDiscGroupDesc = DiscInfo.szDiscGroupDesc
                            MyTaDiscInfoRec.lArtRefNmbr = lHelp
                            MyTaDiscInfoRec.lManualDiscountID = lRef
                            Me.Add(MyTaDiscInfoRec, i)

                            ' save the discount position
                            If dItemWithHighAmount < (dBasic / MyTaArtSaleRec.dTaQty) Then

                                If Not bItemWithHighAmountHasBeebSaved Then
                                    ' if not already saved
                                    dItemWithHighAmount = (dBasic / MyTaArtSaleRec.dTaQty)
                                    lTaCreateNmbrArtSaleWithHighAmount = MyTaArtSaleRec.theHdr.lTaCreateNmbr
                                    lTaCreateNmbrDiscInfoWithHighAmount = MyTaDiscInfoRec.theHdr.lTaCreateNmbr
                                    bItemWithHighAmountSavedHasMultipleQty = IIf(MyTaArtSaleRec.dTaQty > 1, True, False)
                                    bItemWithHighAmountHasBeebSaved = True
                                Else
                                    ' something has been saved
                                    If bItemWithHighAmountSavedHasMultipleQty Then
                                        ' the saved item has qty > 1, so try using the current item
                                        dItemWithHighAmount = (dBasic / MyTaArtSaleRec.dTaQty)
                                        lTaCreateNmbrArtSaleWithHighAmount = MyTaArtSaleRec.theHdr.lTaCreateNmbr
                                        lTaCreateNmbrDiscInfoWithHighAmount = MyTaDiscInfoRec.theHdr.lTaCreateNmbr
                                        bItemWithHighAmountSavedHasMultipleQty = IIf(MyTaArtSaleRec.dTaQty > 1, True, False)
                                        bItemWithHighAmountHasBeebSaved = True
                                    Else
                                        ' the saved item has 1, chech if we have to use the current item
                                        If MyTaArtSaleRec.dTaQty = 1 Then
                                            ' overwrite only if the current item has qty 1
                                            dItemWithHighAmount = (dBasic / MyTaArtSaleRec.dTaQty)
                                            lTaCreateNmbrArtSaleWithHighAmount = MyTaArtSaleRec.theHdr.lTaCreateNmbr
                                            lTaCreateNmbrDiscInfoWithHighAmount = MyTaDiscInfoRec.theHdr.lTaCreateNmbr
                                            bItemWithHighAmountSavedHasMultipleQty = IIf(MyTaArtSaleRec.dTaQty > 1, True, False)
                                            bItemWithHighAmountHasBeebSaved = True
                                        End If
                                    End If
                                End If

                            ElseIf dItemWithHighAmount >= (dBasic / MyTaArtSaleRec.dTaQty) AndAlso _
                                bItemWithHighAmountHasBeebSaved AndAlso bItemWithHighAmountSavedHasMultipleQty AndAlso _
                                MyTaArtSaleRec.dTaQty = 1 Then
                                ' this item has qty 1, an item has been already saved but, the saved item has qty > 1.
                                ' so we prefer to use this item even if its amount is lower than the saved item amount.

                                dItemWithHighAmount = (dBasic / MyTaArtSaleRec.dTaQty)
                                lTaCreateNmbrArtSaleWithHighAmount = MyTaArtSaleRec.theHdr.lTaCreateNmbr
                                lTaCreateNmbrDiscInfoWithHighAmount = MyTaDiscInfoRec.theHdr.lTaCreateNmbr
                                bItemWithHighAmountSavedHasMultipleQty = IIf(MyTaArtSaleRec.dTaQty > 1, True, False)
                                bItemWithHighAmountHasBeebSaved = True

                            End If

                            MyTaDiscInfoRec = Nothing
                        End If

                        'calculate the discount for a single item and store it in the article record
                        'this makes the call of SetPartOfTotalDisc obsolete !!!
                        MyTaArtSaleRec.dPartOfTotalDiscount = MyTaArtSaleRec.dPartOfTotalDiscount + dDisc
                        MyTaArtSaleRec.dTaDiscount = MyTaArtSaleRec.dTaDiscount + (dDisc / MyTaArtSaleRec.dTaQty)

                        ' update the tax
                        ' ==============
                        If MyTaArtSaleRec.ARTinArtSale.bArtKVSet = False Then
                            ' not for a KV Set
                            Me.CalculateArticleTax(MyTaArtSaleRec, i + 1, 0, False, bDummy)
                        End If
                        MyTaArtSaleRec = Nothing
                    End If

                End If

                If CType(TARecords.Item(i), TaBaseRec).sid = PosDef.TARecTypes.iTA_ARTSET AndAlso _
                   CType(TARecords.Item(i), TaBaseRec).theHdr.bIsVoided <> TAdefine.TaAllHdrTypes.IS_VOIDED Then
                    ' only for kv set recal the tax
                    MyTaArtVSetRec = TARecords.Item(i)
                    TaBase = TARecords.Item(GetPositionFromCreationNmbr(MyTaArtVSetRec.theHdr.lTaRefToCreateNmbr))
                    If TaBase.sid = PosDef.TARecTypes.iTA_ART_SALE Then
                        MyTaArtSaleRec = TaBase
                        If MyTaArtSaleRec.ARTinArtSale.bArtKVSet <> False Then
                            ' it is a kv set, then total is filled for this sub article
                            ' we need to recalculate the tax, because the tax is stored for each sub article of a KV set 
                            MyTaArtVSetRec.dTaTotal = MyTaArtVSetRec.dTaTotal - _
                                                      ((DiscInfo.dDiscValue * MyTaArtVSetRec.dTaTotal) / 100)
                            Me.CalculateArticleTax(MyTaArtVSetRec, i + 1, MyTaArtVSetRec.dTaTotal, True, bDummy)
                        End If
                    End If

                End If

            Next i

            'Prepare discount total line for receipt
            If DiscInfo.lDiscListType = PosDef.DiscountTypes.LIST_TOTAL_DISCOUNT_IN_AMOUNT Then
                DiscInfo.dDiscValue = 0
            End If

            If dVal <> dOriginalVal AndAlso lTaCreateNmbrArtSaleWithHighAmount > 0 AndAlso lTaCreateNmbrDiscInfoWithHighAmount > 0 Then

                MyTaArtSaleRec = Me.GetTALine(Me.GetPositionFromCreationNmbr(lTaCreateNmbrArtSaleWithHighAmount))
                MyTaDiscInfoRec = Me.GetTALine(Me.GetPositionFromCreationNmbr(lTaCreateNmbrDiscInfoWithHighAmount))

                If MyTaArtSaleRec.dTaQty > 1 Then
                    SplitArticleIfNeededForTotalDiscountRounding(DiscInfo, MyTaArtSaleRec, MyTaDiscInfoRec, dOriginalVal, dVal)
                Else
                    MyTaArtSaleRec.dPartOfTotalDiscount += dOriginalVal - dVal
                    MyTaArtSaleRec.dTaDiscount += dOriginalVal - dVal
                    MyTaDiscInfoRec.dTotalDiscount += dOriginalVal - dVal
                End If

                dVal = dOriginalVal

                ' update the tax
                ' ==============
                If MyTaArtSaleRec.ARTinArtSale.bArtKVSet = False Then
                    ' not for a KV Set
                    Me.CalculateArticleTax(MyTaArtSaleRec, lTaCreateNmbrArtSaleWithHighAmount + 1, 0, False, bDummy)
                End If
                TARefresh(False)
            End If

            DiscInfo.dTotalDiscount = dVal
            DiscInfo.lIsArtorDept = 0

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("AddTotalDiscInfo"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("AddTotalDiscInfo"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("AddTotalDiscInfo"), "end")
        End Try
    End Sub

    ''' <summary>
    ''' Reorgs this TA. Delets second Header / Footer / Total. 
    ''' </summary>
    Public Overrides Sub Reorg()
        Dim i As Short
        Dim j As Short
        Dim k As Short
        Dim Creation As Integer

        Dim aktsid As Short
        Dim lastsId As Short

        Try

            LOG_FuncStart(getLocationString("Reorg"))

            i = 0
            j = TARecords.Count()
            ' less than 2 elements
            If j < 2 Then
                Exit Sub
            End If

            Do While i < j
                aktsid = 0
                For i = 1 To j
                    lastsId = aktsid
                    aktsid = CType(TARecords.Item(i), TaBaseRec).sid
                    ' we will delete the second TA_HEADER
                    If i <> 1 AndAlso aktsid = PosDef.TARecTypes.iTA_HEADER Then
                        Remove(i)
                        j = j - 1 ' because we have removed 1 record
                        i = 0
                        Exit For
                    End If
                    ' we will delete the second TA_TOTAL
                    If aktsid = PosDef.TARecTypes.iTA_TOTAL AndAlso lastsId = PosDef.TARecTypes.iTA_TOTAL Then
                        Remove(i - 1)
                        j = j - 1 ' because we have removed 1 record
                        i = 0
                        Exit For
                    End If
                    ' we will delete the second TA_FOOTER
                    If aktsid = PosDef.TARecTypes.iTA_FOOTER AndAlso lastsId = PosDef.TARecTypes.iTA_FOOTER Then
                        Remove(i - 1)
                        j = j - 1 ' because we have removed 1 record
                        i = 0
                        Exit For
                    End If
                    ' we will change TA_TOTAL
                    '                TA_ART_SALE
                    If aktsid = PosDef.TARecTypes.iTA_ART_SALE AndAlso lastsId = PosDef.TARecTypes.iTA_TOTAL Then
                        TaChangeRec(i - 1, i)
                        i = 0
                        Exit For
                    End If


                    ' we will change TA_MEDIA and Total
                    If aktsid = PosDef.TARecTypes.iTA_TOTAL AndAlso lastsId = PosDef.TARecTypes.iTA_MEDIA Then
                        TaChangeRec(i - 1, i)
                        i = 0
                        Exit For
                    End If

                    ' check for taxes
                    If CheckReorgTax(aktsid, lastsId, i) Then
                        ' we need to change
                        TaChangeRec(i - 1, i)
                        i = 0
                        Exit For
                    End If

                    ' we will change TA_FOOTER
                    '                TA_MEDIA
                    If aktsid = PosDef.TARecTypes.iTA_MEDIA AndAlso lastsId = PosDef.TARecTypes.iTA_FOOTER Then
                        TaChangeRec(i - 1, i)
                        i = 0
                        Exit For
                    End If

                    ' we will change TA_MEDIA or iTA_FOOTER
                    '                TA_?????
                    If aktsid <> PosDef.TARecTypes.iTA_FOOTER AndAlso _
                       aktsid <> PosDef.TARecTypes.iTA_MEDIA AndAlso _
                       aktsid <> PosDef.TARecTypes.iTA_LOYALTY AndAlso _
                       aktsid <> PosDef.TARecTypes.iTA_TAX_INCLUDED AndAlso _
                       aktsid <> PosDef.TARecTypes.iTA_CUSTSURVEY AndAlso _
                       aktsid <> PosDef.TARecTypes.iTA_COMMENT Then

                        If lastsId = PosDef.TARecTypes.iTA_MEDIA OrElse lastsId = PosDef.TARecTypes.iTA_FOOTER Then
                            TaChangeRec(i - 1, i)
                            i = 0
                            Exit For
                        End If
                    End If
                    'put loyalty record just before the footer
                    If aktsid = PosDef.TARecTypes.iTA_LOYALTY Then
                        k = Me.getLastMediaRecNr
                        If i < k Then
                            TaFromLineafterLine(i, j - 1)
                        End If
                    End If

                    If aktsid = PosDef.TARecTypes.iTA_COMMENT Then
                        If CType(TARecords.Item(i), TaComment).lPresentation.ToString().StartsWith(CInt(TPDotnet.IT.Common.Pos.Italy_PosDef.TARecTypes.iTA_VLL_CUST_MESSAGE)) Then
                            k = Me.getLastMediaRecNr
                            If i < k Then
                                TaFromLineafterLine(i, j - 1)
                            End If
                        Else
                            If lastsId = PosDef.TARecTypes.iTA_MEDIA OrElse lastsId = PosDef.TARecTypes.iTA_FOOTER Then
                                TaChangeRec(i - 1, i)
                                i = 0
                                Exit For
                            End If
                        End If


                    End If
                Next i
            Loop

            ' last one a custsurvey, then change before footer
            i = GetNmbrofRecs()
            If TARecords(i).sid = PosDef.TARecTypes.iTA_CUSTSURVEY AndAlso _
               TARecords(i - 1).sid = PosDef.TARecTypes.iTA_FOOTER Then
                TaChangeRec(i - 1, i)
            End If

            ' Here we check for refs in the TA
            For i = 1 To TARecords.Count()
                Creation = CType(TARecords.Item(i), TaBaseRec).theHdr.lTaRefToCreateNmbr
                If Creation > 0 AndAlso _
                    Creation <> CType(TARecords.Item(i - 1), TaBaseRec).theHdr.lTaCreateNmbr AndAlso _
                    Creation <> CType(TARecords.Item(i - 1), TaBaseRec).theHdr.lTaRefToCreateNmbr Then

                    k = GetPositionFromCreationNmbr(Creation) ' get the act. TaNmbr from the CreateNmbr
                    If k > i Then       'search for chain length of line k if it is behind i
                        ' get length of chain
                        GetLengthOfChain(k, j, True, True)
                        'For j = 0 To TARecords.Count()
                        '    If (j + k + 1) > TARecords.Count Then
                        '        Exit For
                        '    End If
                        '    If CType(TARecords.Item(j + k + 1), TaBaseRec).theHdr.lTaRefToCreateNmbr = 0 Then
                        '        Exit For
                        '    End If
                        'Next
                        TaFromLineafterLine(i, j)   ' drop line i and put it behind the chain of line k
                        i = i - 1                       ' it may be that there are more refs in direct row

                    ElseIf k > 0 Then
                        ' get length of chain
                        GetLengthOfChain(k, j, True, False)
                        ' when i is between current object i and the end object of its main object (j)
                        '  we assume the chain is correct
                        If i > j Then
                            'For LineToChange = k + 1 To j
                            'If Creation <> CType(TARecords.Item(LineToChange), TaBaseRec).theHdr.lTaRefToCreateNmbr Then
                            TaFromLineafterLine(i, j) ' if equal it refs already to k
                            'Exit For
                            'End If
                            'Next LineToChange
                        End If
                    End If
                End If
            Next i

            TARefresh()

            LOG_FuncExit(getLocationString("Reorg"), "end")

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("Reorg"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("Reorg"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("Reorg"), "")
        End Try
    End Sub

#End Region

#Region "Reset TA-Number by date"

    ' Returs the configured transaction reset date.
    ' If it is not defined or an error occurs then returns 29991231
    Public Function GetTransactionResetDate(ByRef szITTransactionResetDate As String) As Boolean
        '=========================================================================================
        '      transaction reset date
        '=========================================================================================
        Dim RECSET As New ADODB_Recordset
        Dim SqlString As String
        Dim TransactionResetDate As Date

        GetTransactionResetDate = False

        ' fill the FooterLines in taobj from database store
        Try
            LOG_FuncStart(getLocationString("GetTransactionResetDate"))

            ' read data
            ' =========
            ' read transaction counter
            ' ========================
            SqlString = "SELECT * FROM TxControlTransactionNmbr " & _
                         "WHERE lRetailStoreID = " & lRetailStoreID & "  " & _
                           "AND lWorkstationNmbr = " & lWorkStationNmbr

            ' set default value
            szITTransactionResetDate = "29991231"

            LOG_Debug(getLocationString("GetTransactionResetDate"), SqlString)
            RECSET.Open(SqlString, conTPPosDB, ADODB_CursorTypeEnum.adOpenForwardOnly, ADODB_LockTypeEnum.adLockReadOnly)
            If RECSET.EOF = False Then

                ' ignore, when szITTransactionResetDate does not exist
                Try
                    If IsDBNull(RECSET.Fields_value("szITTransactionResetDate")) <> True Then
                        szITTransactionResetDate = RECSET.Fields_value("szITTransactionResetDate")
                        GetTransactionResetDate = True
                    End If
                Catch ex As Exception
                    LOG_Error(getLocationString("GetTransactionResetDate"), _
                              "Field 'szITTransactionResetDate' does not exist, 29991231 is used!")
                End Try

                If Not Date.TryParseExact(szITTransactionResetDate, "yyyyMMdd", Nothing, Globalization.DateTimeStyles.None, TransactionResetDate) Then
                    LOG_Error(getLocationString("GetTransactionResetDate"), _
                                                  "Field 'szITTransactionResetDate' not in format YYYYMMDD : " & szITTransactionResetDate & ". 29991231 is used!")
                    szITTransactionResetDate = "29991231"
                    GetTransactionResetDate = False
                End If

            Else
                LOG_Error(getLocationString("GetTransactionResetDate"), _
                          "Entry not present for Workstation " & lWorkStationNmbr & _
                          " not found. Use 29991231 as rollover transaction date!")
            End If

            RECSET.Close()
            RECSET = Nothing

            Exit Function

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetTransactionResetDate"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetTransactionResetDate"), InnerEx)
            End Try
            szITTransactionResetDate = "29991231"
        Finally
            LOG_FuncExit(getLocationString("GetTransactionResetDate"), "Function GetTransactionResetDate returns " & GetTransactionResetDate.ToString)
        End Try

    End Function

    Public Function SetTransactionResetDate(ByRef szITTransactionResetDate As String) As Boolean

        Dim RECSET As New ADODB_Recordset
        Dim SqlString As String

        SetTransactionResetDate = False

        ' fill the FooterLines in taobj from database store
        Try
            LOG_FuncStart(getLocationString("SetTransactionResetDate"))

            ' read data
            ' =========
            SqlString = "SELECT * FROM TxControlTransactionNmbr " & _
                         "WHERE lRetailStoreID = " & lRetailStoreID & "  " & _
                           "AND lWorkstationNmbr = " & lWorkStationNmbr

            LOG_Debug(getLocationString("SetTransactionResetDate"), SqlString)
            RECSET.Open(SqlString, conTPPosDB, ADODB_CursorTypeEnum.adOpenDynamic, ADODB_LockTypeEnum.adLockOptimistic)
            If RECSET.EOF = False Then

                ' ignore, when szITTransactionResetDate does not exist
                Try
                    LOG_Error(getLocationString("SetTransactionResetDate"), _
                              "Set 'szITTransactionResetDate' to " & szITTransactionResetDate)
                    RECSET.Fields_value("szITTransactionResetDate") = szITTransactionResetDate
                    SetTransactionResetDate = True
                Catch ex As Exception
                    LOG_Error(getLocationString("SetTransactionResetDate"), _
                              "Field 'szITTransactionResetDate' does not exist!")
                End Try

                RECSET.Update()

            End If

            RECSET.Close()
            RECSET = Nothing

            Exit Function

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("SetLastTransactionNmbr"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("SetLastTransactionNmbr"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("SetLastTransactionNmbr"), "Function SetLastTransactionNmbr returns " & SetTransactionResetDate.ToString)
        End Try

    End Function

#End Region

End Class
