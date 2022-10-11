Option Strict Off
Option Explicit On
Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports TPDotnet.Pos


Public Class TaMediaRec
    Inherits TPDotnet.Pos.TaMediaRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaMediaRec
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2011, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Overwritten functionality"

    Public Overrides Function UpdEmplMove(ByRef EMPLMOVE As ADODB_Recordset, _
                                      ByRef EMPLMOVEEXT As ADODB_Recordset) As Boolean

        Dim bRecordFound As Boolean

        Try
            LOG_FuncStart(getLocationString("UpdEmplMove"))

            UpdEmplMove = MyBase.UpdEmplMove(EMPLMOVE, EMPLMOVEEXT)
            If UpdEmplMove AndAlso (m_Hdr.bTaValid <> False) Then

                bRecordFound = False
                EMPLMOVEEXT.MoveFirst()
                Do While EMPLMOVEEXT.EOF = False
                    If EMPLMOVEEXT.Fields_value("lMediaMember") = m_PAYMENTinMedia.lMediaMember Then
                        bRecordFound = True
                        Exit Do
                    End If
                    EMPLMOVEEXT.MoveNext()
                Loop

                If bRecordFound Then
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
            Try
                LOG_Error(getLocationString("UpdEmplMove"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("UpdEmplMove"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("UpdEmplMove"), "Function UpdEmplMove returns " & UpdEmplMove.ToString)
        End Try

    End Function

#End Region

End Class
