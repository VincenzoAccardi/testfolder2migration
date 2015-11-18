Option Strict Off
Option Explicit On
Imports System
Imports System.Text
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos

Public Class ART : Inherits TPDotnet.Pos.ART

#Region "Documentation"
    ' ART
    ' ---------- ---------- ---------- ----------
    ' the class implements one Article
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, bASIGLIO, 2014, All rights reserved.
    ' -----------------------------------
#End Region

#Region "Properties"

    Public Overridable Property szITWeightTemplate() As String
        Get
            szITWeightTemplate = m.Fields_Value("szITWeightTemplate")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szITWeightTemplate") = Value
        End Set
    End Property

    Public Overridable Property szITSpecialItemType() As String
        Get
            szITSpecialItemType = m.Fields_Value("szITSpecialItemType")
        End Get
        Set(value As String)
            m.Fields_Value("szITSpecialItemType") = value
        End Set
    End Property

    Public Overridable Property szITSpecialItemType1() As String
        Get
            szITSpecialItemType1 = m.Fields_Value("szITSpecialItemType1")
        End Get
        Set(value As String)
            m.Fields_Value("szITSpecialItemType1") = value
        End Set
    End Property

#End Region

#Region "Overrides"

    Protected Overrides Sub DefineFields()

        MyBase.DefineFields()
        m.Append("szITWeightTemplate", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
        m.Append("szITSpecialItemType", DataField.FIELD_TYPES.FIELD_TYPE_STRING)
        m.Append("szITSpecialItemType1", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

    End Sub

    Public Overrides Function DBRead(ByRef ActCon As ADODB_Connection, _
                               ByRef ArtKey As String, _
                               ByVal lRetailStoreID As Integer) As Short

        ' fields for check the livetime
        Dim szTmpDateValidFrom As String
        Dim szTmpTimeValidFrom As String
        Dim szTmpDateValidTo As String
        Dim szTmpTimeValidTo As String

        Dim ARTRECSET As New ADODB_Recordset
        Dim bRet As Boolean

        DBRead = 0

        Try

            LOG_FuncStart(getLocationString("DBRead"))


            ' to do : improve the dbread 
            SqlString = "SELECT * FROM ItemLookupCode " & _
                            "WHERE szItemLookupCode = '" & ArtKey.Replace("'", "''") & "'" & _
                             " AND lRetailStoreID = " & lRetailStoreID & _
                             " AND bLocked = 0"
            ARTRECSET.Close()
            LOG_Debug(getLocationString("DBRead"), SqlString)
            ARTRECSET.Open(SqlString, ActCon, ADODB_CursorTypeEnum.adOpenForwardOnly, ADODB_LockTypeEnum.adLockReadOnly)
            Dim szItemLookupCode As String = String.Empty

            If Not ARTRECSET.EOF Then
                If Not IsDBNull(ARTRECSET.Fields_value("szItemLookupCode")) Then
                    szItemLookupCode = ARTRECSET.Fields_value("szItemLookupCode")
                End If

                If Not IsDBNull(ARTRECSET.Fields_value("szSizeCode")) Then
                    Me.szSizeCode = ARTRECSET.Fields_value("szSizeCode")
                End If
                If Not IsDBNull(ARTRECSET.Fields_value("szColorCode")) Then
                    Me.szColorCode = ARTRECSET.Fields_value("szColorCode")
                End If

                ArtKey = ARTRECSET.Fields_value("szPOSItemID") ' with this key now to Art
            End If

            ' with the ArtNmbr from ArtEan , we read in Art
            ARTRECSET.Close()
            SqlString = "SELECT * , POSIdentity.bLocked as PosIdentitybLocked, " & _
                         "POSIdentity.szITWeightTemplate as szITWeightTemplate, POSIdentity.szITSpecialItemType as szITSpecialItemType , POSIdentity.szITSpecialItemType1 as szITSpecialItemType1  FROM Item " & _
                         "INNER JOIN POSIdentity ON Item.szItemID = POSIdentity.szItemId " & _
                         "WHERE szPOSItemID = '" & ArtKey.Replace("'", "''") & "'" & _
                          " AND lRetailStoreID = " & lRetailStoreID
            LOG_Debug(getLocationString("DBRead"), SqlString)
            ARTRECSET.Open(SqlString, ActCon, ADODB_CursorTypeEnum.adOpenForwardOnly, ADODB_LockTypeEnum.adLockReadOnly)
            If ARTRECSET.EOF Then
                DBRead = 0 ' ScanNumber also not in Art
                Exit Function
            Else
                ' ok , we have found the article in table Art
                If Not IsDBNull(ARTRECSET.Fields_value("szITMasterEan")) AndAlso String.IsNullOrEmpty(szItemLookupCode) Then
                    szItemLookupCode = ARTRECSET.Fields_value("szITMasterEan")
                End If
            End If
            ArtKey = szItemLookupCode
            ' to do : improve the dbread 
            'SqlString = "SELECT * , POSIdentity.bLocked as PosIdentitybLocked, " & _
            '             "POSIdentity.szITWeightTemplate as szITWeightTemplate, POSIdentity.szITSpecialItemType as szITSpecialItemType, POSIdentity.szITSpecialItemType1 as szITSpecialItemType1,POSIdentity.szITMasterEan as szITMasterEan FROM Item " & _
            '             "INNER JOIN POSIdentity ON Item.szItemID = POSIdentity.szItemId " & _
            '             "WHERE szPOSItemID = '" & ArtKey.Replace("'", "''") & "'" & _
            '              " AND lRetailStoreID = " & lRetailStoreID
            'LOG_Debug(getLocationString("DBRead"), SqlString)
            'ARTRECSET.Open(SqlString, ActCon, ADODB_CursorTypeEnum.adOpenForwardOnly, ADODB_LockTypeEnum.adLockReadOnly)
            '' here we init the values coming only from ArtEan

            'If ARTRECSET.EOF Then ' no Article found, we try ItemLookupCode
            '    SqlString = "SELECT * FROM ItemLookupCode " & _
            '                 "WHERE szItemLookupCode = '" & ArtKey.Replace("'", "''") & "'" & _
            '                  " AND lRetailStoreID = " & lRetailStoreID & _
            '                  " AND bLocked = 0"
            '    ARTRECSET.Close()
            '    LOG_Debug(getLocationString("DBRead"), SqlString)
            '    ARTRECSET.Open(SqlString, ActCon, ADODB_CursorTypeEnum.adOpenForwardOnly, ADODB_LockTypeEnum.adLockReadOnly)
            '    If ARTRECSET.EOF Then
            '        DBRead = 0 ' also nothing found in ArtEan
            '        Exit Function
            '    End If

            '    Dim szItemLookupCode As String = String.Empty

            '    If Not IsDBNull(ARTRECSET.Fields_value("szItemLookupCode")) Then
            '        szItemLookupCode = ARTRECSET.Fields_value("szItemLookupCode")
            '    End If

            '    If Not IsDBNull(ARTRECSET.Fields_value("szSizeCode")) Then
            '        Me.szSizeCode = ARTRECSET.Fields_value("szSizeCode")
            '    End If
            '    If Not IsDBNull(ARTRECSET.Fields_value("szColorCode")) Then
            '        Me.szColorCode = ARTRECSET.Fields_value("szColorCode")
            '    End If

            '    ArtKey = ARTRECSET.Fields_value("szPOSItemID") ' with this key now to Art
            '    ' with the ArtNmbr from ArtEan , we read in Art
            '    ARTRECSET.Close()
            '    SqlString = "SELECT * , POSIdentity.bLocked as PosIdentitybLocked, " & _
            '                 "POSIdentity.szITWeightTemplate as szITWeightTemplate, POSIdentity.szITSpecialItemType as szITSpecialItemType , POSIdentity.szITSpecialItemType1 as szITSpecialItemType1  FROM Item " & _
            '                 "INNER JOIN POSIdentity ON Item.szItemID = POSIdentity.szItemId " & _
            '                 "WHERE szPOSItemID = '" & ArtKey.Replace("'", "''") & "'" & _
            '                  " AND lRetailStoreID = " & lRetailStoreID
            '    LOG_Debug(getLocationString("DBRead"), SqlString)
            '    ARTRECSET.Open(SqlString, ActCon, ADODB_CursorTypeEnum.adOpenForwardOnly, ADODB_LockTypeEnum.adLockReadOnly)
            '    If ARTRECSET.EOF Then
            '        DBRead = 0 ' ScanNumber also not in Art
            '        Exit Function
            '    End If
            '    ArtKey = szItemLookupCode
            'Else
            '    ' ok , we have found the article in table Art
            '    If Not IsDBNull(ARTRECSET.Fields_value("szITMasterEan")) Then
            '        ArtKey = ARTRECSET.Fields_value("szITMasterEan")
            '    End If
            'End If


            ' check the livetime of this article
            ' ----------------------------------
            If IsDBNull(ARTRECSET.Fields_value("szDateValidFrom")) Then
                m.Fields_Value("szDateValidFrom") = " "
                szTmpDateValidFrom = ""
            Else
                szTmpDateValidFrom = ARTRECSET.Fields_value("szDateValidFrom")
                m.Fields_Value("szDateValidFrom") = ARTRECSET.Fields_value("szDateValidFrom")
            End If
            If IsDBNull(ARTRECSET.Fields_value("szTimeValidFrom")) Then
                m.Fields_Value("szTimeValidFrom") = " "
                szTmpTimeValidFrom = ""
            Else
                szTmpTimeValidFrom = ARTRECSET.Fields_value("szTimeValidFrom")
                m.Fields_Value("szTimeValidFrom") = ARTRECSET.Fields_value("szTimeValidFrom")
            End If
            If IsDBNull(ARTRECSET.Fields_value("szDateValidTo")) Then
                m.Fields_Value("szDateValidTo") = " "
                szTmpDateValidTo = ""
            Else
                szTmpDateValidTo = ARTRECSET.Fields_value("szDateValidTo")
                m.Fields_Value("szDateValidTo") = ARTRECSET.Fields_value("szDateValidTo")
            End If
            If IsDBNull(ARTRECSET.Fields_value("szTimeValidTo")) Then
                m.Fields_Value("szTimeValidTo") = " "
                szTmpTimeValidTo = ""
            Else
                szTmpTimeValidTo = ARTRECSET.Fields_value("szTimeValidTo")
                m.Fields_Value("szTimeValidTo") = ARTRECSET.Fields_value("szTimeValidTo")
            End If

            ' for compatibility with older versions we check szTmpDateValidFrom
            bRet = True
            If szTmpDateValidFrom <> "" OrElse szTmpDateValidTo <> "" Then
                bRet = CheckTimeRange(szTmpDateValidFrom, szTmpTimeValidFrom, szTmpDateValidTo, szTmpTimeValidTo)
            End If
            If bRet = False Then
                LOG_Error(getLocationString("DBRead"), "TimeRange of the article " & ArtKey & " not valid")
                Exit Function
            End If

            Try
                ' ignore when ProdRangeId is not available
                If IsDBNull(ARTRECSET.Fields_value("lProdRangeID")) Then
                    lProdRangeID = 0
                Else
                    lProdRangeID = ARTRECSET.Fields_value("lProdRangeID")

                    ' we look if there is a ProdRangeStoreGroupPosIdentityAffiliation present
                    ' -----------------------------------------------------------------------
                    bRet = DBReadProductRangePrice(ActCon, (ARTRECSET.Fields_value("szPOSItemID")), lProdRangeID, lRetailStoreID)
                End If
            Catch ex As Exception
                ' ignore, do nothing
            End Try


            ' put in the new values for article in the object
            ' -----------------------------------------------
            FillFields(ARTRECSET)


            ARTRECSET.Close()
            ARTRECSET = Nothing
            DBRead = 1 ' returnvalue ok
            Exit Function


        Catch ex As Exception
            Try
                LOG_Error(getLocationString("DBRead"), ex)
                LOG_Error(getLocationString("DBRead"), " SqlString: " & SqlString)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("DBRead"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("DBRead"), "Function DBRead returns " & DBRead.ToString)
        End Try

    End Function

    Public Overrides Sub FillFields(ARTRECSET As TPDotnet.Pos.ADODB_Recordset)

        Try
            MyBase.FillFields(ARTRECSET)

            If IsDBNull(ARTRECSET.Fields_value("szITWeightTemplate")) Then
                m.Fields_Value("szITWeightTemplate") = ""
            Else
                m.Fields_Value("szITWeightTemplate") = ARTRECSET.Fields_value("szITWeightTemplate")
            End If

            If IsDBNull(ARTRECSET.Fields_value("szITSpecialItemType")) Then
                m.Fields_Value("szITSpecialItemType") = ""
            Else
                m.Fields_Value("szITSpecialItemType") = ARTRECSET.Fields_value("szITSpecialItemType")
            End If

            If IsDBNull(ARTRECSET.Fields_value("szITSpecialItemType1")) Then
                m.Fields_Value("szITSpecialItemType1") = ""
            Else
                m.Fields_Value("szITSpecialItemType1") = ARTRECSET.Fields_value("szITSpecialItemType1")
            End If



        Catch ex As Exception
            Try
                LOG_Error(getLocationString("FillFields"), ex)
                LOG_Error(getLocationString("FillFields"), " SqlString: " & SqlString)

            Catch InnerEx As Exception
                LOG_ErrorInTry(getLocationString("FillFields"), InnerEx)
            End Try
        Finally
            LOG_FuncExit(getLocationString("FillFields"), "Function FillFields returns ")
        End Try

    End Sub

#End Region

End Class
