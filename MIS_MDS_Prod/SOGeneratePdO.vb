Module SOGeneratePdO

    'Public WithEvents oComp As SAPbobsCOM.Company = Nothing

    Sub GeneratePdOFromSO(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler

        Dim oSOToMFGGrid As SAPbouiCOM.Grid

        Dim idx As Long


        Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
        Dim oSalesOrderLines As SAPbobsCOM.Document_Lines = Nothing

        Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
        Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines = Nothing

        Dim errConnect As String = ""

        Dim oPdODocSeriesRec As SAPbobsCOM.Recordset

        Dim strQry As String = ""
        Dim oPdODocSeriesOrder As String = ""
        Dim oPdODocSeriesJasa As String = ""

        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific



        'GRID - Order by column checkbox
        oSOToMFGGrid.Columns.Item("Release PdO").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


        'Get PdO Doc. Series 
        'oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPdODocSeriesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ' FG-002: PENJUALAN JASA (SERIENAME:2011JS), FG-001: PENJUALAN ORDER (SERIENAME:2011) 


        strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) = 'JS' AND Indicator = YEAR(GETDATE()) "

        oPdODocSeriesRec.DoQuery(strQry)
        '??? 
        If oPdODocSeriesRec.RecordCount <> 0 Then
            oPdODocSeriesJasa = oPdODocSeriesRec.Fields.Item("Series").Value
        Else
            MsgBox("Production Order Document Series Jasa Tidak ada, Mohon Setup PdO Document Series!")
            Exit Sub
        End If

        strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) <> 'JS' AND Indicator = YEAR(GETDATE()) "
        oPdODocSeriesRec.DoQuery(strQry)
        '??? 
        If oPdODocSeriesRec.RecordCount <> 0 Then
            oPdODocSeriesOrder = oPdODocSeriesRec.Fields.Item("Series").Value
        Else
            MsgBox("Production Order Document Series Kaca Order Tidak ada, Mohon Setup PdO Document Series!")
            Exit Sub
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPdODocSeriesRec)
        oPdODocSeriesRec = Nothing
        GC.Collect()

        If oPdODocSeriesOrder <> "" And oPdODocSeriesJasa <> "" Then

            'If oSOToMFGGrid.Rows.Count > 5 Then
            '    SBO_Application.MessageBox("Minimal 5 To Generate So", 1, "OK")
            'Else
            'Loop only selected/checked in grid rows and exit.
            For idx = oSOToMFGGrid.Rows.Count - 1 To 0 Step -1
                oForm.Items.Item("SoNumber").Click()
                'SBO_Application.SetStatusBarMessage("Generating PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                oApp.SetStatusBarMessage("Generating PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                If oSOToMFGGrid.DataTable.GetValue(1, oSOToMFGGrid.GetDataTableRowIndex(idx)) = "Y" Then

                    oForm.Items.Item("SoNumber").Click()


                    If Not oCompany.InTransaction Then
                        oCompany.StartTransaction()
                    End If


                    Dim oRS As SAPbobsCOM.Recordset

                    'vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strQry = "SELECT DocNum FROM OWOR WHERE Status <> 'C' AND OriginNum =  " & oSOToMFGGrid.DataTable.GetValue(4, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
                        & " AND ItemCode = '" & oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) & "' "
                    oRS.DoQuery(strQry)
                    'oRS.DoQuery("UPDATE RDR1 SET U_bacthNum = 'b321' where docentry = 249 and linenum = 0")

                    oForm.Items.Item("BPCardCode").Click()


                    'If oRS.RecordCount = 0 Then -- if duplicate don't insert PdO
                    If oRS.RecordCount <> 0 Or oRS.RecordCount = 0 Then

                        oProd1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)


                        ' Series PdO JS = 202 (NNM1) objectCode = 202 (OWOR PdO) series id = 45

                        ' IMPORTANT !!!
                        ' PdO SERIES YEAR 2011, 2011JS PdO JASA, SERIES# = 45
                        ' PdO SERIES YEAR 2011, 2011   PdO KACA SERIES# = 27 

                        If oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) = "FG-001" Then
                            'oProd1.Series = 27
                            oProd1.Series = oPdODocSeriesOrder
                        Else
                            'oProd1.Series = 45
                            oProd1.Series = oPdODocSeriesJasa
                        End If

                        'oProd1.ItemNo = "KTF12CLXX589"
                        oProd1.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        oProd1.PlannedQuantity = oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        ''oProdOrder.DueDate = oSOToMFGGrid.DataTable.GetValue(13, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        'PdO Posting Date = SO Posting Date
                        'oProd1.PostingDate = oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        oProd1.PostingDate = Format(Now, "yyyy-MM-dd")

                        Dim dueDt As DateTime
                        Dim sodt As DateTime
                        Dim sodelivdt As DateTime
                        Dim dtdiff As Integer

                        sodt = oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        sodelivdt = oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
                        dtdiff = DateDiff(DateInterval.Day, CDate(oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)), CDate(oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)))
                        'sodelivdt = ""
                        dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
                        dueDt = DateAdd(DateInterval.Day, IIf(dtdiff < 0, 0, dtdiff), Now)

                        'PdO Due Date = SO Deliv. Date
                        'oProd1.DueDate = Today + n days (so date - so deliv date)
                        ''oProd1.DueDate = oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        oProd1.DueDate = dueDt

                        'oprod1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                        oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                        oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                        'oprod1.Warehouse = "01"
                        'oProd1.Warehouse = "FG-001"

                        oProd1.Warehouse = oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        oProd1.CustomerCode = oSOToMFGGrid.DataTable.GetValue(7, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
                        ' so docnum
                        oProd1.ProductionOrderOriginEntry = oSOToMFGGrid.DataTable.GetValue(3, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        oProd1.UserFields.Fields.Item("U_PoD_Pcm").Value = oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString
                        oProd1.UserFields.Fields.Item("U_PdO_Lcm").Value = oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString
                        oProd1.UserFields.Fields.Item("U_PdO_Bentuk").Value = oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        'Dim QTY_LUASM2 As Double
                        'QTY_LUASM2 = _
                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString))


                        oProd1.UserFields.Fields.Item("U_SO_Luas_M2").Value = _
                        Left(CStr( _
                            Math.Round( _
                              (IIf(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
                              IIf(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
                              IIf(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString)) / 10000) _
                              , 4) _
                            ) _
                        , 10)

                        'oprod1.UserFields.Fields.Item("U_PoD_Pcm").Value = "p100cm"
                        'oprod1.UserFields.Fields.Item("U_PdO_Lcm").Value = "L90cm"
                        'oprod1.UserFields.Fields.Item("U_PdO_Bentuk").Value = "segi"
                        'oprod1.UserFields.Fields.Item("U_NBS_OnHoldReason").Value = "test123"

                        oProd1.UserFields.Fields.Item("U_ORDRDocEntry").Value = _
                        oSOToMFGGrid.DataTable.GetValue(3, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString()
                        oProd1.UserFields.Fields.Item("U_ORDRLineNum").Value = _
                        oSOToMFGGrid.DataTable.GetValue(19, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString()

                        oProdLine1 = oProd1.Lines

                        ' Generate one line - Dummy item
                        oProdLine1.ItemNo = "XDUMMY"
                        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                        oProdLine1.Warehouse = "SRV-DL"


                        'MsgBox(GC.GetTotalMemory(True))

                        lRetCode = oProd1.Add()


                        If lRetCode <> 0 Then
                            oCompany.GetLastError(lErrCode, sErrMsg)
                            'vCompany.GetLastError(lErrCode, sErrMsg)
                            'SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                            oApp.MessageBox(lErrCode & ": " & sErrMsg)
                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                            If oCompany.InTransaction Then
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        Else

                            ''vCompany.GetNewObjectCode(tmpKey)
                            ''vCompany.GetNewObjectCode(PdOno)
                            'oCompany.GetNewObjectCode(PdOno)
                            'tmpKey = Convert.ToInt32(PdOno)

                            ' !!!! Make sure before create another object type-> clear previous/current object type.
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
                            oProdLine1 = Nothing

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
                            oProd1 = Nothing

                            GC.Collect()

                            oForm.Items.Item("SoNumber").Click()



                            If oCompany.InTransaction Then
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If

                        End If

                        'SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        oApp.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                    oRS = Nothing

                    ' by Toin 2011-02-09 Check Duplicate PdO before Generate PdO


                    GC.Collect()
                    'MsgBox(GC.GetTotalMemory(True))

                    'MsgBox("generating... PdO; DONE!!!")

                Else
                    Exit For
                End If
            Next

        End If  ' Checking PdO Series

        oApp.MessageBox("Generating PdO.... Finished !!! ", 1, "Ok")

        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

    Sub SOToMFGEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim SOToMFGQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oSOToMFGGrid As SAPbouiCOM.Grid

        Try
            'oForm = SBO_Application.Forms.Item("mds_p1")
            'SBO_Application.MessageBox("Form Already Open")
            oForm = oApp.Forms.Item("mds_p1")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            'fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds"
            fcp.UniqueID = "mds_p1"

            'fcp.XmlData = LoadFromXML("sotomfg.srf")
            'fcp.XmlData = MenuCreation.LoadFromXML("sotomfg.srf")
            fcp.XmlData = LoadFromXML("sotomfg.srf")

            'oForm = SBO_Application.Forms.AddEx(fcp)
            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("SODateFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("SODateTo", SAPbouiCOM.BoDataType.dt_DATE)


            'Default value for SO Date
            oForm.DataSources.UserDataSources.Item("SODateFrom").Value = DateTime.Today.ToString("yyyyMMdd")

            oForm.DataSources.UserDataSources.Item("SODateTo").Value = DateTime.Today.ToString("yyyyMMdd")

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("SoNumber", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            ''Dim bpCFL01 As MISToolbox
            ''bpCFL01 = New MISToolbox

            ''bpCFL01.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, "SOBPCFL1", "SOBPCFL2", _
            ''                "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C")

            ''bpCFL01 = Nothing
            ''GC.Collect()



            oEditText = oForm.Items.Item("BPCardCode").Specific
            oButton = oForm.Items.Item("BPButton").Specific

            oEditText.DataBind.SetBound(True, "", "BPDS")


            ''oEditText.ChooseFromListUID = "SOBPCFL1"
            ''oEditText.ChooseFromListAlias = "CardCode"
            ''oButton.ChooseFromListUID = "SOBPCFL2"

            oEditText = oForm.Items.Item("SoNumber").Specific

            oForm.Items.Item("SODateFrom").Width = 100
            oEditText = oForm.Items.Item("SODateFrom").Specific
            oEditText.DataBind.SetBound(True, "", "SODateFrom")

            oForm.Items.Item("SODateTo").Width = 100
            oEditText = oForm.Items.Item("SODateTo").Specific
            oEditText.DataBind.SetBound(True, "", "SODateTo")



            oItem = oForm.Items.Item("myGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200


            oSOToMFGGrid = oItem.Specific

            oSOToMFGGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)


            SOToMFGQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, " _
            & " 'Y' [Release PdO], " _
            & " CONVERT(VARCHAR(10), T0.DocDate, 102) [SO Date], T0.DocEntry [DocEntry], T0.DocNum [DocNum], VisOrder + 1 [SO Line]," _
            & " T3.SlpName [Sales Rep.], T0.CardCode [Cust. Code], T0.CardName [Customer Name], " _
            & " T1.ItemCode FG,T1.Dscription FGName, Quantity, " _
            & " T1.WhsCode, T2.InvntryUom UOM, T0.DocDueDate [Exp Delivery Date], " _
            & " T1.U_SO_Pcm PanjangInCm, T1.U_SO_Lcm LebarInCm, " _
            & " Case " _
            & " when T1.[U_SO_Bentuk] ='S' then 'Segi' " _
            & " when T1.[U_SO_Bentuk] ='J' then 'Jenjang' " _
            & " when T1.[U_SO_Bentuk] ='O' then 'Oval' " _
            & " when T1.[U_SO_Bentuk] ='B' then 'Bulat' " _
            & " when T1.[U_SO_Bentuk] ='G' then 'Bending' " _
            & " when T1.[U_SO_Bentuk] ='M' then 'Mal' " _
            & " when T1.[U_SO_Bentuk] ='X' then 'Others' " _
            & " Else 'Undefined' " _
            & " End SO_Bentuk, " _
            & " Case " _
            & " when T0.U_SOApprovalStatus ='A' then 'SO Approved' " _
            & " when T0.U_SOApprovalStatus ='D' then 'SO Draft' " _
            & " when T0.U_SOApprovalStatus ='O' then 'SO Reguler' " _
            & " Else T0.U_SOApprovalStatus " _
            & " End [SO Approval Status], " _
            & " T1.LineNum [SO LineNum] " _
            & " FROM ORDR T0 " _
            & " LEFT JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " _
            & " LEFT JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " _
            & " LEFT JOIN OSLP T3 ON T1.SlpCode = T3.SlpCode " _
            & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
            & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
            & " AND T1.LineStatus = 'O' " _
            & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _
            & " AND (T1.U_MIS_ReleasePdOFlag = '' OR T1.U_MIS_ReleasePdOFlag IS NULL) " _
            & " AND T0.U_SOApprovalStatus <> 'D' " _
            & " ORDER BY T0.DocDate, T0.DocNum, VisOrder DESC "

            '& " AND T1.WhsCode = 'FG-002' " _
            '            & " AND T1.U_MIS_ReleasePdOFlag = '' AND T1.U_MIS_SupplyWith = 'M' "



            ' Grid #: 1
            oForm.DataSources.DataTables.Add("SOToMFGLst")
            oForm.DataSources.DataTables.Item("SOToMFGLst").ExecuteQuery(SOToMFGQuery)
            oSOToMFGGrid.DataTable = oForm.DataSources.DataTables.Item("SOToMFGLst")


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oSOToMFGGrid = Nothing

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try

        ''oForm.Top = 150
        ''oForm.Left = 330
        ''oForm.Width = 900


        SOToMFGQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, " _
            & " 'Y' [Release PdO], " _
            & " CONVERT(VARCHAR(10), T0.DocDate, 102) [SO Date], T0.DocEntry [DocEntry], T0.DocNum [DocNum], VisOrder + 1 [SO Line]," _
            & " T3.SlpName [Sales Rep.], T0.CardCode [Cust. Code], T0.CardName [Customer Name], " _
            & " T1.ItemCode FG,T1.Dscription FGName, Quantity, " _
            & " T1.WhsCode, T2.InvntryUom UOM, T0.DocDueDate [Exp Delivery Date], " _
            & " T1.U_SO_Pcm PanjangInCm, T1.U_SO_Lcm LebarInCm, " _
            & " Case " _
            & " when T1.[U_SO_Bentuk] ='S' then 'Segi' " _
            & " when T1.[U_SO_Bentuk] ='J' then 'Jenjang' " _
            & " when T1.[U_SO_Bentuk] ='O' then 'Oval' " _
            & " when T1.[U_SO_Bentuk] ='B' then 'Bulat' " _
            & " when T1.[U_SO_Bentuk] ='G' then 'Bending' " _
            & " when T1.[U_SO_Bentuk] ='M' then 'Mal' " _
            & " when T1.[U_SO_Bentuk] ='X' then 'Others' " _
            & " Else 'Undefined' " _
            & " End SO_Bentuk, " _
            & " Case " _
            & " when T0.U_SOApprovalStatus ='A' then 'SO Approved' " _
            & " when T0.U_SOApprovalStatus ='D' then 'SO Draft' " _
            & " when T0.U_SOApprovalStatus ='O' then 'SO Reguler' " _
            & " Else T0.U_SOApprovalStatus " _
            & " End [SO Approval Status], " _
            & " T1.LineNum [SO LineNum] " _
            & " FROM ORDR T0 " _
            & " LEFT JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " _
            & " LEFT JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " _
            & " LEFT JOIN OSLP T3 ON T1.SlpCode = T3.SlpCode " _
            & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
            & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
            & " AND T1.LineStatus = 'O' " _
            & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _
            & " AND (T1.U_MIS_ReleasePdOFlag = '' OR T1.U_MIS_ReleasePdOFlag IS NULL) " _
            & " AND T0.U_SOApprovalStatus <> 'D' " _
            & " ORDER BY T0.DocDate, T0.DocNum, VisOrder DESC "
        '& " AND T1.WhsCode = 'FG-002' " _
        '    & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _

        '        & " JOIN MARUNI_SOTRIAL..mis_sofg002 T4 ON T0.DocNum = T4.[Document Number] AND T1.LineNum = T4.[Row Number] " _
        '            & " AND T1.U_MIS_ReleasePdOFlag = '' AND T1.U_MIS_SupplyWith = 'M' "


        oForm.DataSources.DataTables.Item(0).ExecuteQuery(SOToMFGQuery)

        oForm.Items.Item("BPCardCode").Click()



        RearrangeGrid(oForm)


        oForm.Visible = True

    End Sub


    Private Sub RearrangeGrid(ByVal oForm As SAPbouiCOM.Form)


        Dim oColumn As SAPbouiCOM.EditTextColumn

        Dim oSOToMFGGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)

        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

        oSOToMFGGrid.RowHeaders.Width = 50

        'Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        oColumn = oSOToMFGGrid.Columns.Item("Cust. Code")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner '"2" -> BP Master
        oColumn.Editable = False

        oColumn = oSOToMFGGrid.Columns.Item("DocEntry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
        oColumn.Editable = False

        oColumn = oSOToMFGGrid.Columns.Item("FG")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items '"2"
        oColumn.Editable = False


        oSOToMFGGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oSOToMFGGrid.Columns.Item("Release PdO").TitleObject.Sortable = True

        oSOToMFGGrid.Columns.Item("DocEntry").Width = 60
        oSOToMFGGrid.Columns.Item("Cust. Code").Width = 130


        oColumn = oSOToMFGGrid.Columns.Item("Customer Name")
        oColumn.Editable = False

        oSOToMFGGrid.Columns.Item("SO Date").Width = 80
        oSOToMFGGrid.Columns.Item("SO Date").TitleObject.Sortable = True


        oSOToMFGGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        ' Set Total Row count in colum title/header
        'oSOToMFGGrid.Columns.Item(0).TitleObject.Caption = oSOToMFGGrid.Rows.Count.ToString


        If oForm.DataSources.DataTables.Item(0).Rows.Count <> 0 _
        And oSOToMFGGrid.DataTable.GetValue(0, 0) <> "" Then
            oForm.Items.Item("cmdGenPdO").Enabled = True
        Else
            oForm.Items.Item("cmdGenPdO").Enabled = False
        End If

        oSOToMFGGrid.Columns.Item("#").Editable = False
        oSOToMFGGrid.Columns.Item(1).Editable = True
        oSOToMFGGrid.Columns.Item("SO Date").Editable = False
        oSOToMFGGrid.Columns.Item("DocEntry").Editable = False
        oSOToMFGGrid.Columns.Item("DocNum").Editable = False
        oSOToMFGGrid.Columns.Item("SO Line").Editable = False
        oSOToMFGGrid.Columns.Item("Sales Rep.").Editable = False
        oSOToMFGGrid.Columns.Item("Cust. Code").Editable = False
        oSOToMFGGrid.Columns.Item("FG").Editable = False
        oSOToMFGGrid.Columns.Item("FGName").Editable = False
        oSOToMFGGrid.Columns.Item("Quantity").Editable = False
        oSOToMFGGrid.Columns.Item("UOM").Editable = False
        oSOToMFGGrid.Columns.Item("Exp Delivery Date").Editable = False

        oSOToMFGGrid.Columns.Item("WhsCode").Editable = False
        oSOToMFGGrid.Columns.Item("PanjangInCm").Editable = False
        oSOToMFGGrid.Columns.Item("LebarInCm").Editable = False
        oSOToMFGGrid.Columns.Item("SO_Bentuk").Editable = False

        oSOToMFGGrid.Columns.Item("SO Approval Status").Editable = False
        oSOToMFGGrid.Columns.Item("SO LineNum").Editable = False


        oSOToMFGGrid.RowHeaders.Width = 20
        oSOToMFGGrid.Columns.Item("#").Width = 30
        oSOToMFGGrid.Columns.Item(1).Width = 20
        oSOToMFGGrid.Columns.Item("SO Date").Width = 60
        oSOToMFGGrid.Columns.Item("DocEntry").Width = 60
        oSOToMFGGrid.Columns.Item("DocNum").Width = 60
        oSOToMFGGrid.Columns.Item("SO Line").Width = 30
        oSOToMFGGrid.Columns.Item("Cust. Code").Width = 80
        oSOToMFGGrid.Columns.Item("FG").Width = 100
        oSOToMFGGrid.Columns.Item("Exp Delivery Date").Width = 80
        oSOToMFGGrid.Columns.Item("WhsCode").Width = 50
        oSOToMFGGrid.Columns.Item("PanjangInCm").Width = 50
        oSOToMFGGrid.Columns.Item("LebarInCm").Width = 50
        oSOToMFGGrid.Columns.Item("SO_Bentuk").Width = 80


        'Dim oEditText As SAPbouiCOM.EditText
        'Dim oStaticText As SAPbouiCOM.StaticText
        'Dim oButton As SAPbouiCOM.Button

        oForm.Items.Item("1").Top = 400


        Dim sboDate As String
        Dim dDate As DateTime

        'dDate = DateTime.Now

        'sbo formatdate
        'sboDate = oMIS_Utils.fctFormatDate(dDate, oCompany)
        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oSOToMFGGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Sub LoadSO(ByVal oForm As SAPbouiCOM.Form)
        Dim SOToMFGQuery As String

        If oForm.Items.Item("BPCardCode").Specific.string = "" Then
            oApp.SetStatusBarMessage("Customer must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Exit Sub
        End If

        If oForm.Items.Item("SoNumber").Specific.value = "" Then
            oApp.SetStatusBarMessage("So Number Must Be Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Exit Sub
        End If

        If oForm.Items.Item("SODateTo").Specific.string = "" Then
            oForm.Items.Item("SODateTo").Specific.string = oForm.Items.Item("SODateFrom").Specific.string
        End If

        SOToMFGQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, " _
            & " 'Y' [Release PdO], " _
            & " CONVERT(VARCHAR(10), T0.DocDate, 102) [SO Date], T0.DocEntry [DocEntry], T0.DocNum [DocNum], VisOrder + 1 [SO Line]," _
            & " T3.SlpName [Sales Rep.], T0.CardCode [Cust. Code], T0.CardName [Customer Name], " _
            & " T1.ItemCode FG,T1.Dscription FGName, Quantity, " _
            & " T1.WhsCode, T2.InvntryUom UOM, T0.DocDueDate [Exp Delivery Date], " _
            & " T1.U_SO_Pcm PanjangInCm, T1.U_SO_Lcm LebarInCm, " _
            & " Case " _
            & " when T1.[U_SO_Bentuk] ='S' then 'Segi' " _
            & " when T1.[U_SO_Bentuk] ='J' then 'Jenjang' " _
            & " when T1.[U_SO_Bentuk] ='O' then 'Oval' " _
            & " when T1.[U_SO_Bentuk] ='B' then 'Bulat' " _
            & " when T1.[U_SO_Bentuk] ='G' then 'Bending' " _
            & " when T1.[U_SO_Bentuk] ='M' then 'Mal' " _
            & " when T1.[U_SO_Bentuk] ='X' then 'Others' " _
            & " Else 'Undefined' " _
            & " End SO_Bentuk, " _
            & " Case " _
            & " when T0.U_SOApprovalStatus ='A' then 'SO Approved' " _
            & " when T0.U_SOApprovalStatus ='D' then 'SO Draft' " _
            & " when T0.U_SOApprovalStatus ='O' then 'SO Reguler' " _
            & " Else T0.U_SOApprovalStatus " _
            & " End [SO Approval Status], " _
            & " T1.LineNum [SO LineNum] " _
            & " FROM ORDR T0 " _
            & " LEFT JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " _
            & " LEFT JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " _
            & " LEFT JOIN OSLP T3 ON T1.SlpCode = T3.SlpCode " _
            & " LEFT JOIN OWOR T4 ON T4.OriginNum = T0.DocNum AND T4.ItemCode = T1.ItemCode " _
            & "     AND T4.U_ORDRDocEntry = T0.DocEntry AND T4.U_ORDRLineNum = T1.LineNum " _
        & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
            & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
            & " AND T1.LineStatus = 'O' " _
            & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _
            & " AND (T1.U_MIS_ReleasePdOFlag = '' OR T1.U_MIS_ReleasePdOFlag IS NULL) " _
            & " AND T0.Docnum = " & oForm.Items.Item("SoNumber").Specific.value & " " _
            & " AND ISNULL(T0.U_SOApprovalStatus, '') <> 'D' " _
            & " AND T1.Quantity - ISNULL(T4.PlannedQty, 0) > 0 " _
            & " ORDER BY T0.DocDate, T0.DocNum, VisOrder DESC "
        '& " AND T1.WhsCode = 'FG-002' " _


        '& " JOIN MARUNI_SOTRIAL..mis_sofg002 T4 ON T0.DocNum = T4.[Document Number] AND T1.LineNum = T4.[Row Number] " _
        '    & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _

        '            & " AND T1.U_MIS_ReleasePdOFlag = '' AND T1.U_MIS_SupplyWith = 'M' "


        '            & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _

        '        & " , T1.U_SO_P1, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xJob1, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xDC1," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xIC1," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xFOH1," _
        '& " T1.U_SO_P2, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xJob2, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xDC2," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xIC2," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xFOH2," _
        '& " T1.U_SO_P3, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xJob3, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xDC3," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xIC3," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xFOH3," _
        '& " T1.U_SO_P4, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xJob4, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xDC4," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xIC4," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xFOH4," _
        '& " T1.U_SO_P5, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xJob5, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xDC5," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xIC5," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xFOH5," _
        '& " T1.U_SO_P6, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xJob6, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xDC6," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xIC6," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xFOH6," _
        '& " T1.U_SO_P7, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xJob7, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xDC7," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xIC7," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xFOH7 " _

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(SOToMFGQuery)

        RearrangeGrid(oForm)

    End Sub

End Module
