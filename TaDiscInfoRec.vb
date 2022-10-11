Option Strict Off
Option Explicit On
Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports TPDotnet.Pos

Public Class TaDiscInfoRec
    Inherits TPDotnet.Pos.TaDiscInfoRec


#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaDiscInfoRec
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2009, All rights reserved.
    ' -----------------------------------
#End Region

    Public Overrides Function GetPresentation(ByRef theDecive As Short, ByRef thegReceipt As TPDotnet.Pos.gReceipt, ByRef bTrainingMode As Integer) As String

        GetPresentation = ""

        Try
            LOG_FuncStart(getLocationString("GetPresentation"))

            If m.Fields_Value("dDiscValue") = 0 AndAlso _
                m.Fields_Value("dTotalDiscount") = 0 Then
                Exit Function
            End If

            If m.Fields_Value("lDiscExtNmbr") = PosDef.DiscountTypes.iLINE_DISCOUNT AndAlso _
                m.Fields_Value("lDiscListType") = PosDef.DiscountTypes.iLINE_DISCOUNT Then
                Me.lPresentation = 1000
            End If

            Return MyBase.GetPresentation(theDecive, thegReceipt, bTrainingMode)

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("GetPresentation"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("GetPresentation"), InnerEx)
            End Try
        Finally
            If GetPresentation Is Nothing Then
                LOG_FuncExit(getLocationString("GetPresentation"), " returns nothing")
            Else
                LOG_FuncExit(getLocationString("GetPresentation"), "Function GetPresentation returns " & GetPresentation.ToString)
            End If
        End Try

    End Function

End Class
