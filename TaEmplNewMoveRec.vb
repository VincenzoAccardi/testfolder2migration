Option Strict Off
Option Explicit On
Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports TPDotnet.Pos


Public Class TaEmplNewMoveRec
    Inherits TPDotnet.Pos.TaEmplNewMoveRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaEmplNewMoveRec
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2011, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Overwritten functionality"

    Public Overrides Function Fill_NEW(ByRef ActCon As TPDotnet.Pos.ADODB_Connection, ByRef lFromMediaMember As Long, ByVal szMode As String, ByVal lOperatorID As Long) As Short

        Try
            LOG_FuncStart(getLocationString("Fill_NEW"))

            Fill_NEW = MyBase.Fill_NEW(ActCon, lFromMediaMember, szMode, lOperatorID)
            If Fill_NEW > 0 Then
                ' Invalid transactions because szSessionID is missing
                ' In this case the base function already uses the EMPLMOVE/POSMOVE field to fill the szSessionID
                ' So we can just log this problem
                If String.IsNullOrEmpty(m.Fields_Value("EMPLMOVEEXT.szSessionID")) Then
                    LOG_Error(getLocationString("Fill_NEW"), "szSessionID is Null Or Empty")
                End If

            End If

            Exit Function

        Catch ex As Exception
            LOG_Error(getLocationString("Fill_NEW"), ex)
        Finally
            LOG_FuncExit(getLocationString("Fill_NEW"), "Function Fill_NEW returns " & Fill_NEW.ToString)
        End Try

    End Function

#End Region

End Class
