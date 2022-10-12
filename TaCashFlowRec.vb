Option Strict Off
Option Explicit On
Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports TPDotnet.Pos

Public Class TaCashFlowRec
    Inherits TPDotnet.Pos.TaCashFlowRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaCashFlowRec
    ' ---------- ---------- ---------- ----------
    ' Add : amounts already picked up in dValPickup
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2009, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Properties"

    Public Property dValPickUp() As Double
        Get
            dValPickUp = m.Fields_value("dValPickUp")
        End Get
        Set(ByVal value As Double)
            m.Fields_value("dValPickUp") = value
        End Set
    End Property

#End Region

#Region "Overwritten functionality"

    Protected Overrides Sub DefineFields()

        Try
            LOG_Info(getLocationString("DefineFields"), "starting")

            MyBase.DefineFields()

            m.Append("dValPickUp", DataField.FIELD_TYPES.FIELD_TYPE_DECIMAL)

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DefineFields"), ex)
            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DefineFields"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DefineFields"), "")
        End Try

    End Sub

    Protected Overrides Sub InitFields()

        Try
            LOG_Info(getLocationString("InitFields"), "starting")

            MyBase.InitFields()

            'Alpha numeric fields are initialized by '' and numeric fields are initialized by 0.

        Catch ex As Exception
            Try
                LOG_Error(getLocationString("InitFields"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("InitFields"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("InitFields"), "")
        End Try

    End Sub

    Protected Overrides Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

    Public Overrides Function UpdEmplMove(ByRef EMPLMOVE As ADODB_Recordset, _
                                          ByRef EMPLMOVEEXT As ADODB_Recordset) As Boolean

        Dim recordfound As Boolean

        Try
            LOG_FuncStart(getLocationString("UpdEmplMove"))

            UpdEmplMove = MyBase.UpdEmplMove(EMPLMOVE, EMPLMOVEEXT)
            If UpdEmplMove Then

                recordfound = False
                EMPLMOVEEXT.MoveFirst()
                Do While EMPLMOVEEXT.EOF = False
                    If EMPLMOVEEXT.Fields_value("lMediaMember") = m.Fields_Value("lMediaMember") Then
                        recordfound = True
                        Exit Do
                    End If
                    EMPLMOVEEXT.MoveNext()
                Loop

                If recordfound Then
                    ' Invalid transactions because szSessionID is missing
                    ' In this case the base function uses the EMPLMOVE/POSMOVE field to fill the szSessionID if this is a new record
                    ' So we can test the field and replace it with the EMPLMOVE/POSMOVE field
                    If String.IsNullOrEmpty(EMPLMOVEEXT.Fields_value("szSessionID")) Then
                        LOG_Error(getLocationString("UpdEmplMove"), "EMPLMOVEEXT.szSessionID is Null Or Empty")
                        If String.IsNullOrEmpty(EMPLMOVE.Fields_value("szSessionID")) Then
                            LOG_Error(getLocationString("UpdEmplMove"), "EMPLMOVE.szSessionID is Null Or Empty also!")
                        Else
                            EMPLMOVEEXT.Fields_value("szSessionID") = EMPLMOVE.Fields_value("szSessionID")
                            EMPLMOVEEXT.Update()
                        End If
                    End If
                End If

            End If

            Exit Function

        Catch ex As Exception
            LOG_Error(getLocationString("UpdEmplMove"), ex)
        Finally
            LOG_FuncExit(getLocationString("UpdEmplMove"), "Function UpdEmplMove returns " & UpdEmplMove.ToString)
        End Try

    End Function

#End Region

End Class
