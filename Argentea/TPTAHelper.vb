Imports System
Imports TPDotnet.Pos
Imports System.IO



Public Class TPTAHelper
    Implements ITPTAMediaSwapHelper

#Region "Documentation"
    ' ********** ********** ********** **********
    ' Argentea EFT
    ' ---------- ---------- ---------- ----------
    ' Author : Emanuele Gualtierotti
    ' Wincor Nixdorf Retail Consulting
    ' -----------------------------------
    ' Copyright by Wincor Nixdorf Retail Consulting
    ' 20090, Basiglio, 2014, All rights reserved.
    ' -----------------------------------
#End Region


    Public Function CreateTA(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef MyTaRec As TPDotnet.Pos.TaBaseRec, ByRef updTANmbr As Boolean) As TPDotnet.Pos.TA
        CreateTA = Nothing

        Try
            If MyTaRec Is Nothing Then
                Exit Function
            End If

            CreateTA = createTheTA(TheModCntr, taobj.colObjects)
            Dim szMode As String = GetOnlyMode(TheModCntr)
            'initialize the new transaction with the values used to show on display
            '====================================================================== 
            If CreateTA.iTaStatus = PosDef.TAStatus.iTASTATUS_NOT_STARTED Then
                'Initialize the new Transaction
                CreateTA.szSignOnName = TheModCntr.szSignOnName
                CreateTA.szPrintCode = TheModCntr.szPrintCode
                CreateTA.szActEmployeeName = TheModCntr.szActEmployeeName
                CreateTA.lActOperatorID = TheModCntr.lActOperatorID
                CreateTA.lRetailStoreID = TheModCntr.lRetailStoreID
                CreateTA.szWorkstationGroupID = TheModCntr.szWorkstationGroupID
                CreateTA.szWorkstationID = TheModCntr.szWorkstationID
                CreateTA.lWorkStationNmbr = TheModCntr.lWorkstationNmbr
                CreateTA.bTrainingMode = TheModCntr.bTrainingMode
                CreateTA.szMode = szMode
                CreateTA.iExactness = TheModCntr.iEXACTNESS_IN_DIGITS
                CreateTA.TAStart(TheModCntr.contxt, TheModCntr.con)

                ' here we reload the parameters from table parameter
                TheModCntr.ReadParameter()

                ' fill the headerlines from globstoreval
                fillHeaderLines(TheModCntr, CreateTA)
            End If

            CreateTA.bPrintReceipt = False
            CreateTA.bTAtoFile = True
            CreateTA.bDelete = True ' ok , we will delete this TA

            ' add out eft record
            CreateTA.Add(MyTaRec)
            CreateTA.TAEnd(fillFooterLines(TheModCntr.con, CreateTA, TheModCntr))

            If updTANmbr AndAlso CreateTA.lactTaNmbr = taobj.lactTaNmbr Then

                For i As Integer = 1 To taobj.GetNmbrofRecs
                    'first remove the Header
                    If taobj.GetsId(i) = PosDef.TARecTypes.iTA_HEADER Then
                        taobj.Remove(i)
                        ' to get really a new number, store in all cases the one of the above new created one
                        CreateTA.NewTa2Reg()

                        ' now assign a new Transactionnumber from TxControlTransactionNumber
                        taobj.AssignRegValues()

                        ' finally build a new header and put it at Pos 1 in TA
                        Dim MyHeaderRec As TaHdrRec = taobj.CreateTaObject(PosDef.TARecTypes.iTA_HEADER)
                        MyHeaderRec.Fill(taobj)

                        taobj.AddObject(MyHeaderRec, 0)
                        fillHeaderLines(TheModCntr, taobj)
                        Exit For
                    End If
                Next i

            End If

        Catch ex As Exception
            CreateTA = Nothing
        End Try

    End Function

    Public Function WriteTA(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        WriteTA = False
        Dim funcName As String = "PrintEFTReceipt"
        Dim WriteObject As clsEndTAHandling

        Try
            WriteObject = createPosObject(Of clsEndTAHandling)(TheModCntr, "clsEndTAHandling", 0)
            If Not WriteObject Is Nothing Then
                taobj.bDelete = True
                taobj.bPrintReceipt = False
                taobj.bTAtoFile = False
                WriteTA = WriteObject.EndTA(taobj, TheModCntr)
            End If
            WriteObject = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Function PrintReceipt(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr) As Boolean
        PrintReceipt = False
        Dim funcName As String = "PrintEReceipt"
        Dim PrintObject As clsEndTAHandling

        Try

            PrintObject = createPosObject(Of clsEndTAHandling)(TheModCntr, "clsEndTAHandling", 0)
            If Not PrintObject Is Nothing Then
                taobj.bPrintReceipt = True
                PrintObject.PrintTa(taobj, TheModCntr)
                If PrintObject.IsSomethingPrinted Then
                    If TheModCntr.OposPrinter_work_with Then
                        TheModCntr.OposCutPaper()
                    End If
                    PrintReceipt = True
                End If

            End If
            PrintObject = Nothing

        Catch ex As Exception

        Finally

        End Try

    End Function


    Public Function GetMediaMemberByCardType(ByRef CardType As String) As Integer Implements ITPTAMediaSwapHelper.GetMediaMemberByCardType
        GetMediaMemberByCardType = 0 ' default: no mapping found
        Dim funcName As String = "GetMediaMemberByCardType"
        Dim szFileName As String = ""

        Try

            'read the swap table
            'format of each swap file line is Type=MediaMember (eg. 00=400)
            szFileName = getPosConfigurationPath().TrimEnd("\") + "\EftSwap.txt"
            If File.Exists(szFileName) Then
                Dim lines() As String = File.ReadAllLines(szFileName)
                For Each line As String In lines
                    Dim Type As String = "", MediaMember As String = ""
                    Type = line.Split("=")(0)
                    MediaMember = line.Split("=")(1)
                    If Not Type Is Nothing AndAlso Type <> "" AndAlso Not MediaMember Is Nothing AndAlso MediaMember <> "" Then
                        If Type = CardType Then
                            GetMediaMemberByCardType = Integer.Parse(MediaMember)
                            Exit For
                        End If
                    End If
                Next
            End If

        Catch ex As Exception

        End Try

    End Function

    Public Sub SwapElectronicMedia(ByRef taobj As TPDotnet.Pos.TA, ByRef TheModCntr As TPDotnet.Pos.ModCntr, ByRef TheMediaRec As TPDotnet.Pos.TaMediaRec, ByRef CardType As String) Implements ITPTAMediaSwapHelper.SwapElectronicMedia
        Dim funcName As String = "SwapElectronicMedia"
        Dim clsSelectMedia As clsSelectMedia
        Dim lMediaMemberToAssign As Integer = 0
        Dim bRet As Boolean

        Try

            'get an instance of the class with reads the DB-table for the Media
            clsSelectMedia = createPosModelObject(Of clsSelectMedia)(TheModCntr, "clsSelectMedia", 0, True)
            If clsSelectMedia Is Nothing Then
                ' no media Record for this extension in database present
                LOG_Error(funcName, "configuration error: module not found: clsSelectMedia")
                Exit Sub
            End If

            lMediaMemberToAssign = GetMediaMemberByCardType(CardType)
            If lMediaMemberToAssign > 0 AndAlso lMediaMemberToAssign <> TheMediaRec.PAYMENTinMedia.lMediaMember Then
                If Not TheMediaRec Is Nothing Then
                    bRet = clsSelectMedia.FillPaymentDataFromID(TheModCntr, TheMediaRec.PAYMENTinMedia, _
                                                             lMediaMemberToAssign, taobj, taobj.colObjects)
                    If bRet = False Then
                        ' no media Record for this extension in database present
                        LOG_Error(funcName, "MediaMember " & lMediaMemberToAssign.ToString & " not found in database")
                        clsSelectMedia = Nothing
                        TheMediaRec = Nothing
                        Exit Sub
                    Else
                        'AddDetailToMedia(taobj, TheMediaRec)
                    End If
                End If
            End If

            clsSelectMedia = Nothing

        Catch ex As Exception

        Finally

        End Try

    End Sub

End Class
