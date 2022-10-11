Option Strict Off
Option Explicit On
Imports System
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class TaArtReturnRec
    Inherits TPDotnet.Pos.TaArtReturnRec



#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaArtReturnRec
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2011, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Overwritten functionality"

    Public Overrides Function GetDiscountableTotal() As Decimal
        Try
            LOG_FuncStart(getLocationString("GetDiscountableTotal"))

            GetDiscountableTotal = 0

            If m_Hdr.bIsVoided <> TAdefine.TaAllHdrTypes.IS_VOIDED AndAlso m_Hdr.bTaValid <> 0 Then
                If m_ARTinArtReturn.bTotalDiscountAllowed <> 0 Then
                    'GetDiscountableTotal = m.Fields_Value("dTaTotal") - (m.Fields_Value("dTaDiscount") * m.Fields_Value("dTaQty"))
                End If
            End If

            Exit Function


        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetDiscountableTotal"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetDiscountableTotal"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("GetDiscountableTotal"), "Function GetDiscountableTotal returns " & GetDiscountableTotal.ToString)
        End Try
    End Function

    'Public Overrides Function GetDiscount(ByRef index As Short) As Decimal
    '    GetDiscount = MyBase.GetDiscount(index)
    'End Function

    'Public Overrides Function GetFiscalTotal() As Decimal
    '    GetFiscalTotal = MyBase.GetFiscalTotal()
    'End Function

    'Public Overrides Function GetTotal() As Decimal
    '    GetTotal = MyBase.GetTotal()
    'End Function

    Protected Overrides Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class
