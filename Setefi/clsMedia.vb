Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.IT.Common.Pos

Public Class clsMedia
    Inherits TPDotnet.IT.Common.Pos.clsMedia

#Region "Overridden methods"

    Protected Overrides Function DoSpecialHandling4CreditCardsOnline(ByRef taobj As TA, ByRef TheModCntr As ModCntr, ByRef MyTaMediaRec As TaMediaRec, ByRef MyTaMediaMemberDetailRec As TaMediaMemberDetailRec) As Boolean

        Dim Ret As Integer
        Dim MyTaMediaSwap As TaMediaMemberSwap
        Dim bIsVoidReceipt As Boolean
        Dim bIsMediaCorrect As Boolean

        DoSpecialHandling4CreditCardsOnline = False

        Try
            LOG_FuncStart(getLocationString("DoSpecialHandling4CreditCardsOnline"))

            If MyTaMediaRec.dTaPaidTotal < 0.0 Or _
                 MyTaMediaRec.dReturn < 0.0 Then
                'for performance reasons 
                For Ret = 1 To taobj.GetNmbrofRecs
                    If taobj.GetTALine(Ret).sid = PosDef.TARecTypes.iTA_VOID_RECEIPT Then
                        bIsVoidReceipt = True
                        LOG_Debug(getLocationString("DoSpecialHandling4CreditCardsOnline"), "We are in a void receipt")
                        Exit For
                    End If
                    If taobj.GetTALine(Ret).sid = PosDef.TARecTypes.iTA_MEDIAMEMBER_SWAP Then
                        MyTaMediaSwap = taobj.GetTALine(Ret)
                        If MyTaMediaSwap.sFunction = PosDef.TATaMediaMemberSwap.iCORRECT Then
                            bIsMediaCorrect = True
                            LOG_Debug(getLocationString("DoSpecialHandling4CreditCardsOnline"), "We are in a media correction")
                            MyTaMediaSwap = Nothing
                        End If
                        Exit For
                    End If
                Next Ret
            End If

            'set the ModCtrl to thePosEFT
            'set the taobj to thePosEFT
            TheModCntr.thePosEFT.Initialize(taobj, TheModCntr)

            If MyTaMediaRec.dTaPaidTotal > 0.0# And MyTaMediaRec.dReturn = 0.0# Then
                ' payment
                Ret = TheModCntr.thePosEFT.PostEFTDevice(taobj, MyTaMediaRec)
                If Ret = 0 Then
                    DoSpecialHandling4CreditCardsOnline = True
                End If
            ElseIf bIsVoidReceipt Or bIsMediaCorrect Then
                ' void/swap media receipt
                LOG_Debug(getLocationString("DoSpecialHandling4CreditCardsOnline"), "before calling VoidEftDevice")
                Ret = TheModCntr.thePosEFT.VoidEftDevice(taobj, MyTaMediaRec)
                If Ret = 0 Then
                    DoSpecialHandling4CreditCardsOnline = True
                End If
            ElseIf MyTaMediaRec.dTaPaidTotal < 0 Then
                ' line/immediate void
                LOG_Debug(getLocationString("DoSpecialHandling4CreditCardsOnline"), "before calling VoidEftDevice")
                Ret = TheModCntr.thePosEFT.VoidEftDevice(taobj, MyTaMediaRec)
                If Ret = 0 Then
                    DoSpecialHandling4CreditCardsOnline = True
                End If
            End If

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DoSpecialHandling4CreditCardsOnline"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DoSpecialHandling4CreditCardsOnline"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DoSpecialHandling4CreditCardsOnline"), "Function DoSpecialHandling4CreditCardsOnline returns " & DoSpecialHandling4CreditCardsOnline.ToString)
        End Try

    End Function

#End Region

#Region "internal methods"
    Protected Overridable Sub SetAccountAndBank(ByVal taobj As TA, ByVal TheModCntr As ModCntr, ByVal MyTaMediaRec As TaMediaRec)

        Try

            ' todo : understand if is needed

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("SetAccountAndBank"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("SetAccountAndBank"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("SetAccountAndBank"), "Function SetAccountAndBank ")
        End Try

    End Sub

#End Region
End Class
