Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.Compatibility
Imports TPDotnet.Pos
Imports TPDotnet.Italy.Common.Pos

Public Class clsMedia_DaSistemi
    Inherits clsMedia

#Region "Overridden methods"
    Protected Overrides Function DoSpecialHandling4CreditCardsOnline(ByRef taobj As TA, ByRef TheModCntr As ModCntr, ByRef MyTaMediaRec As TaMediaRec, ByRef MyTaMediaMemberDetailRec As TaMediaMemberDetailRec) As Boolean

        Try

            DoSpecialHandling4CreditCardsOnline = MyBase.DoSpecialHandling4CreditCardsOnline(taobj, TheModCntr, MyTaMediaRec, MyTaMediaMemberDetailRec)

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
