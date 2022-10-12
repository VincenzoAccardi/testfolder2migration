Option Strict Off
Option Explicit On 
Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports TPDotnet.Pos

Public Class TaArtSaleRec : Inherits TPDotnet.Pos.TaArtSaleRec

#Region "Documentation"
    ' ********** ********** ********** **********
    ' TaArtSaleRec
    ' ---------- ---------- ---------- ----------
    ' Add : ScaleBag Handling
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Assago, 2008, All rights reserved.
    ' -----------------------------------
#End Region


#Region "Properties"
   
#End Region
#Region "Overwritten functionality"

    Protected Overrides Sub DefineFields()

        Try
            LOG_Info(getLocationString("DefineFields"), "starting")

            MyBase.DefineFields()

            ' Standard fields
            ' ---------------

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

    Public Overrides Function GetPresentation(ByRef theDevice As Short, _
                                              ByRef thegReceipt As gReceipt, _
                                              ByRef bTrainingMode As Integer) As String
        Dim PresentationKey As String = ""

        GetPresentation = ""

        Try

            GetPresentation = MyBase.GetPresentation(theDevice, thegReceipt, bTrainingMode)

            If (Me.lArtRef <> 0) Then

                ' presentation for the subarticles
                PresentationKey = PosDef.TARecTypes.iTA_ART_SALE & ".27"
                GetPresentation = GetTheLines(theDevice, Me, thegReceipt, bTrainingMode, PresentationKey)
            ElseIf Me.ARTinArtSale.szItemCategoryTypeCode = "BAGYY" Or _
                        Me.ARTinArtSale.szItemCategoryTypeCode = "BAGYN" Or _
                        Me.ARTinArtSale.szItemCategoryTypeCode = "BAGNN" Or _
                        Me.ARTinArtSale.szItemCategoryTypeCode = "BAGNY" Then
                ' presentation for the header articles without turnover
                If Me.dTaTotal = 0 Then
                    PresentationKey = PosDef.TARecTypes.iTA_ART_SALE & ".26"
                    GetPresentation = GetTheLines(theDevice, Me, thegReceipt, bTrainingMode, PresentationKey)
                End If
            End If

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

    Public Overrides Function DBRead(ByRef ActCon As ADODB_Connection, _
                                    ByRef ArtKey As String, _
                                    ByVal lRetailStoreID As Integer) As Short

        DBRead = 0

        Dim RetRead As Short

        Try
            LOG_FuncStart(getLocationString("DBRead"))



            RetRead = ARTinArtSale.DBRead(ActCon, ArtKey, lRetailStoreID)
            m.Fields_Value("szItemLookupCode") = ArtKey
            DBRead = RetRead

            Exit Function


        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DBRead"), ex)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DBRead"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DBRead"), "Function DBRead returns " & DBRead.ToString)
        End Try

    End Function

    Protected Overrides Function getLocationString(ByRef actMethode As String) As String
        getLocationString = TypeName(Me) & "." & actMethode & " "
    End Function

#End Region

End Class