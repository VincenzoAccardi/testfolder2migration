Imports TPDotnet.IT.Common.Pos
Imports TPDotnet.Pos
Imports ARGLIB = PAGAMENTOLib
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Xml.XPath
Imports System.Linq
Imports System.IO

Public Class ValassisCouponController
#Region "IValassisCouponValidation"
    Public Function ValidationValassis(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "ValidationValassis"
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IValassisCouponValidation function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            Dim szMessageOut As String = String.Empty

            If Not TheModCntr.ObjectCash.ContainsKey("Valassis_Progressive") Then TheModCntr.ObjectCash.Add("Valassis_Progressive", taobj.lactTaNmbr.ToString + ".1")

            Dim lProgressive As Integer = 1
            If TheModCntr.ObjectCash("Valassis_Progressive").ToString.Split(".").First = taobj.lactTaNmbr.ToString Then
                Int32.TryParse(TheModCntr.ObjectCash("Valassis_Progressive").ToString.Split(".").Last, lProgressive)
            Else
                TheModCntr.ObjectCash("Valassis_Progressive") = taobj.lactTaNmbr.ToString + ".1"
            End If


            Dim szNFCData As String = String.Empty
            Dim szWalletID As String = String.Empty
            Dim szWorkstationNmbr As String = String.Empty
            Dim szOperationID As String = taobj.lActOperatorID.ToString
            Dim szTaNmbr As String = taobj.lactTaNmbr.ToString
            Dim szCouponCode As String = String.Empty
            Dim szCustomerID As String = String.Empty

            If MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA Then
                szCouponCode = MyCurrentRecord.GetPropertybyName("szBarcode")
            Else
                Dim MyTaCustRec As TPDotnet.Pos.TaCustomerRec = CType(taobj.GetTALine(taobj.getCustRecNr), TPDotnet.Pos.TaCustomerRec)
                szCustomerID = MyTaCustRec.CUSTinCustomer.szCustomerID
            End If

            response.ReturnCode = ArgenteaCOMObject.ValidationValassis(lProgressive, szCouponCode, szNFCData, szWalletID, szWorkstationNmbr, szOperationID, szTaNmbr, szCustomerID, szMessageOut)


            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                szMessageOut = "KO;;" + szMessageOut + ";;;;"
            Else
                Dim szFileName As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Valassis.ReponseFile").Trim()
                szMessageOut = String.Empty
                If File.Exists(szFileName) Then
                    Dim lines() As String = File.ReadAllLines(szFileName)
                    For Each line As String In lines
                        szMessageOut += line
                    Next
                End If
                szMessageOut = System.Net.WebUtility.HtmlDecode(szMessageOut)

            End If

            LOG_Error(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". CouponCode: " & szCouponCode & ", Customer: " & szCustomerID & ". Output: " & szMessageOut)

            TheModCntr.ObjectCash("Valassis_Progressive") = taobj.lactTaNmbr.ToString + "." + (lProgressive + 1).ToString

            response.MessageOut = szMessageOut
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.ValidationValassis

            paramArg.Copies = 0
            paramArg.PrintWithinTA = True
            'paramArg.MultiMedia = True

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response
    End Function

    Public Function NotificationValassis(ByRef ArgenteaCOMObject As ARGLIB.argpay,
                            ByRef taobj As TPDotnet.Pos.TA,
                            ByRef TheModCntr As TPDotnet.Pos.ModCntr,
                            ByRef MyCurrentRecord As TPDotnet.Pos.TaBaseRec,
                            ByRef paramArg As ArgenteaParameters) As ArgenteaResponse
        Dim funcName As String = "NotificationValassis"
        Dim response As New ArgenteaResponse

        Try
            LOG_Debug(getLocationString(funcName), "We are entered in Argentea IValassisCouponNotification function")
            ' collect the input parameters
            LOG_Debug(getLocationString(funcName), "LoadCommonFunctionParameter")

            Dim szMessageOut As String = String.Empty
            If Not TheModCntr.ObjectCash.ContainsKey("Valassis_Progressive") Then TheModCntr.ObjectCash.Add("Valassis_Progressive", taobj.lactTaNmbr.ToString + ".1")

            Dim lProgressive As Integer = 1
            If TheModCntr.ObjectCash("Valassis_Progressive").ToString.Split(".").First = taobj.lactTaNmbr.ToString Then
                Int32.TryParse(TheModCntr.ObjectCash("Valassis_Progressive").ToString.Split(".").Last, lProgressive)
            Else
                TheModCntr.ObjectCash("Valassis_Progressive") = taobj.lactTaNmbr.ToString + ".1"
            End If

            Dim szTaNmbr As String = taobj.lactTaNmbr.ToString
            Dim szNodeXML As String = String.Empty




            Dim szRequestNumberID As String = "0"


            If MyCurrentRecord.sid = TPDotnet.Pos.TARecTypes.iTA_MEDIA Then
                Dim myRecMedia As TPDotnet.Pos.TaMediaRec = CType(MyCurrentRecord, TPDotnet.Pos.TaMediaRec)
                If myRecMedia.dTaPaidTotal < 0 Then
                    For Each base As TPDotnet.Pos.TaBaseRec In taobj.taCollection
                        If base.sid = TPDotnet.IT.Common.Pos.Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE AndAlso base.theHdr.lTaRefToCreateNmbr = MyCurrentRecord.theHdr.lTaCreateNmbr Then
                            If Not base.ExistField("szCouponCancelReason") Then base.AddField("szCouponCancelReason", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
                            base.setPropertybyName("szCouponCancelReason", CInt(IValassisNotificationCancelReasonCode.TRXSUSPEND).ToString)
                            szRequestNumberID = base.GetPropertybyName("szTransactionID")
                        End If

                    Next
                End If
            End If


            If MyCurrentRecord.ExistField("szTransactionID") Then
                szRequestNumberID = MyCurrentRecord.GetPropertybyName("szTransactionID")
            End If
            Dim myTa As New TPDotnet.Pos.TA
            myTa.Clone(taobj)
            For Each base As TPDotnet.Pos.TaBaseRec In taobj.taCollection
                If base.sid = TPDotnet.Pos.TARecTypes.iTA_ART_SALE Then
                    myTa.Add(base)
                ElseIf base.sid = TPDotnet.IT.Common.Pos.Italy_PosDef.TARecTypes.iTA_EXTERNAL_SERVICE Then
                    Dim myExtService As TPDotnet.IT.Common.Pos.TaExternalServiceRec = CType(base, TPDotnet.IT.Common.Pos.TaExternalServiceRec)
                    If (myExtService.ExistField("szTransactionID") AndAlso szRequestNumberID = myExtService.GetPropertybyName("szTransactionID")) Then
                        myTa.Add(base)
                    End If
                End If
            Next

            If Not Common.ApplyFilterStyleSheet(TheModCntr, myTa, "NotificationValassis.xslt", szNodeXML) Then

            End If
            LOG_Error(getLocationString(funcName), "Request: Progressive: " + lProgressive.ToString + ", szTaNmbr: " + szTaNmbr.ToString + ", szRequestNumberID: " + szRequestNumberID.ToString + ",szNodeXML :" + szNodeXML.ToString)

            response.ReturnCode = ArgenteaCOMObject.NotificationValassis(lProgressive, szTaNmbr, szRequestNumberID, szNodeXML, szMessageOut)

            If response.ReturnCode <> ArgenteaFunctionsReturnCode.OK Then
                szMessageOut = "KO;;" + szMessageOut + ";;;;"
            Else
                Dim szFileName As String = TheModCntr.getParam(PARAMETER_DLL_NAME + ".Valassis.ReponseFile").Trim()
                szMessageOut = String.Empty
                If File.Exists(szFileName) Then
                    Dim lines() As String = File.ReadAllLines(szFileName)
                    For Each line As String In lines
                        szMessageOut += line
                    Next
                End If
                szMessageOut = System.Net.WebUtility.HtmlDecode(szMessageOut)
            End If

            LOG_Error(getLocationString(funcName), "ReturnCode: " & response.ReturnCode.ToString & ". Output: " & szMessageOut)

            TheModCntr.ObjectCash("Valassis_Progressive") = taobj.lactTaNmbr.ToString + "." + (lProgressive + 1).ToString

            response.MessageOut = szMessageOut
            response.CharSeparator = CharSeparator.Semicolon
            response.FunctionType = InternalArgenteaFunctionTypes.NotificationValassis
            response.TransactionID = szRequestNumberID

            paramArg.Copies = 0
            paramArg.PrintWithinTA = False

        Catch ex As Exception
            LOG_Error(getLocationString(funcName), ex.Message)
            response.ReturnCode = ArgenteaFunctionsReturnCode.KO
        Finally
        End Try
        Return response
    End Function
#End Region
#Region "Overridable"

    Protected Overridable Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region
End Class
