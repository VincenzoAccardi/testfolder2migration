Option Strict Off
Option Explicit On
Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports TPDotnet.Pos

Public Class TaFtrRec
    Inherits TPDotnet.Pos.TaFtrRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaFtrRec
    ' ---------- ---------- ---------- ----------
    ' Add : Counter of items sold
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2009, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Properties"

    ''' <summary>
    ''' gets / sets the sold items
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property lSoldItems() As Integer
        Get
            lSoldItems = m.Fields_value("lSoldItems")
        End Get
        Set(ByVal Value As Integer)
            m.Fields_Value("lSoldItems") = Value
        End Set
    End Property

#End Region

#Region "Overwritten functionality"

    Protected Overrides Sub DefineFields()

        Try
            LOG_Info(getLocationString("DefineFields"), "starting")

            MyBase.DefineFields()

            m.Append("lSoldItems", DataField.FIELD_TYPES.FIELD_TYPE_INTEGER)

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
#End Region

End Class
