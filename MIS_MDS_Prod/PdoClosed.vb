Module PdoClosed

    '-------------- Yadi FC ----------------------------
    Sub LoadProductionClosed(ByVal oForm As SAPbouiCOM.Form)
        Dim PCQuery As String
        PCQuery = " SELECT CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY A.DOCNUM)) #, CASE WHEN  ISNULL(A.PlannedQty,0) - ISNULL(A.CmpltQty,0) = 0 THEN 'Y' ELSE 'N' END [Check], A.DocEntry [Pdo Entry] " _
                                  & " , A.DocNum [Pdo #], A.ItemCode [Product No.], B.ItemName [Product Description], ISNULL(A.PlannedQty,0) [Planned Qty] " _
                                  & " , ISNULL(A.CmpltQty,0) [Complete Qty], A.PostDate [Pdo Order Date], A.DueDate [Pdo Due Date], (select DocEntry from ORDR where DocNum = A.OriginNum ) [SO DocEntry], A.OriginNum [SO #] " _
                                  & " , A.CardCode [Customer Code], C.CardName [Customer Name]  " _
                              & " FROM OWOR A " _
                                  & " INNER JOIN OITM B ON A.ItemCode = B.ItemCode " _
                                  & " INNER JOIN OCRD C ON A.CardCode = C.CardCode " _
                              & " WHERE A.Status = 'R' " _
                                  & " AND A.PostDate >= '" & Format(CDate(oForm.Items.Item("txtDate1").Specific.string), "yyyyMMdd") & "' " _
                                  & " AND A.PostDate <= '" & Format(CDate(oForm.Items.Item("txtDate2").Specific.string), "yyyyMMdd") & "' "

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(PCQuery)

        RearrangePCGrid(oForm)

    End Sub
    '-------------- Yadi FC ----------------------------

    '-------------- Yadi FC ----------------------------
    Sub UpdateProductionClosed(ByVal oForm As SAPbouiCOM.Form)
        Dim oProductionOrder As SAPbobsCOM.ProductionOrders = Nothing
        Dim oPCGrid As SAPbouiCOM.Grid
        Dim intI As Integer = 0
        Dim RetVal As Boolean

        oPCGrid = oForm.Items.Item("grdPC").Specific
        oPCGrid.Columns.Item("Check").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        For intI = oPCGrid.Rows.Count - 1 To 0 Step -1
            'SBO_Application.SetStatusBarMessage("Processing.... Start !!! " & intI + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            oApp.SetStatusBarMessage("Processing.... Start !!! " & intI + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            'oForm.Items.Item("Pdo #").Click()

            If oPCGrid.DataTable.GetValue(1, oPCGrid.GetDataTableRowIndex(intI)) = "Y" Then
                oProductionOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                'oForm.Items.Item("Pdo #").Click()

                If Not oCompany.InTransaction Then
                    oCompany.StartTransaction()
                End If

                RetVal = oProductionOrder.GetByKey(oPCGrid.DataTable.GetValue(2, oPCGrid.GetDataTableRowIndex(intI)))

                If RetVal = True Then
                    'oProductionOrder.ProductionOrderOriginEntry = oPCGrid.DataTable.GetValue(2, oPCGrid.GetDataTableRowIndex(intI))
                    oProductionOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposClosed
                    lRetCode = oProductionOrder.Update()

                    If lRetCode <> 0 Then
                        If oCompany.InTransaction Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If

                        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionOrder)
                        'oProductionOrder = Nothing
                    Else

                        If oCompany.InTransaction Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionOrder)
                    oProductionOrder = Nothing

                    GC.Collect()

                End If
            Else
                Exit For
            End If
        Next


        'SBO_Application.SetStatusBarMessage("Processing.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        oApp.SetStatusBarMessage("Processing.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        'SBO_Application.MessageBox("Processing.... Finished !!! ", 1, "Ok")
        oApp.MessageBox("Processing.... Finished !!! ", 1, "Ok")

        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionOrder)
        'oProductionOrder = Nothing
        'GC.Collect()

    End Sub
    '-------------- Yadi FC ----------------------------

    '-------------- Yadi FC ----------------------------
    Sub ProductionClosed()
        Dim oForm As SAPbouiCOM.Form

        Dim PCQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oProductionClosedGrid As SAPbouiCOM.Grid

        Try
            'oForm = SBO_Application.Forms.Item("MDS_P6")
            'SBO_Application.MessageBox("Form Already Open")
            oForm = oApp.Forms.Item("MDS_P6")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            'fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "MDS_P6"
            fcp.UniqueID = "MDS_P6"

            fcp.XmlData = LoadFromXML("ProductionClosed.srf")
            'oForm = SBO_Application.Forms.AddEx(fcp)
            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("txtDate1", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("txtDate2", SAPbouiCOM.BoDataType.dt_DATE)


            'Default value for SO Date
            oForm.DataSources.UserDataSources.Item("txtDate1").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("txtDate2").Value = DateTime.Today.ToString("yyyyMMdd")

            oForm.Items.Item("txtDate1").Width = 100
            oEditText = oForm.Items.Item("txtDate1").Specific
            oEditText.DataBind.SetBound(True, "", "txtDate1")

            oForm.Items.Item("txtDate2").Width = 100
            oEditText = oForm.Items.Item("txtDate2").Specific
            oEditText.DataBind.SetBound(True, "", "txtDate2")



            oItem = oForm.Items.Item("grdPC")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 150


            oProductionClosedGrid = oItem.Specific

            oProductionClosedGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)


            PCQuery = " SELECT CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY A.DOCNUM)) #,  CASE WHEN  ISNULL(A.PlannedQty,0) - ISNULL(A.CmpltQty,0) = 0 THEN 'Y' ELSE 'N' END [Check], A.DocEntry [Pdo Entry]  " _
                          & " , A.DocNum [Pdo #], A.ItemCode [Product No.], B.ItemName [Product Description], ISNULL(A.PlannedQty,0) [Planned Qty] " _
                          & " , ISNULL(A.CmpltQty,0) [Complete Qty], A.PostDate [Pdo Order Date], A.DueDate [Pdo Due Date], (select DocEntry from ORDR where DocNum = A.OriginNum ) [SO DocEntry], A.OriginNum [SO #] " _
                          & " , A.CardCode [Customer Code], C.CardName [Customer Name]  " _
                      & " FROM OWOR A " _
                          & " INNER JOIN OITM B ON A.ItemCode = B.ItemCode " _
                          & " INNER JOIN OCRD C ON A.CardCode = C.CardCode " _
                      & " WHERE A.Status = 'R' " _
                          & " AND A.PostDate >= '" & Format(CDate(oForm.Items.Item("txtDate1").Specific.string), "yyyyMMdd") & "' " _
                          & " AND A.PostDate <= '" & Format(CDate(oForm.Items.Item("txtDate2").Specific.string), "yyyyMMdd") & "' "

            ' Grid #: 1
            oForm.DataSources.DataTables.Add("PCLst")
            oForm.DataSources.DataTables.Item("PCLst").ExecuteQuery(PCQuery)
            oProductionClosedGrid.DataTable = oForm.DataSources.DataTables.Item("PCLst")


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oProductionClosedGrid = Nothing

            GC.Collect()

        End Try


        PCQuery = " SELECT CONVERT(VARCHAR(10), ROW_NUMBER() OVER(ORDER BY A.DOCNUM)) #,  CASE WHEN  ISNULL(A.PlannedQty,0) - ISNULL(A.CmpltQty,0) = 0 THEN 'Y' ELSE 'N' END [Check], A.DocEntry [Pdo Entry] " _
                          & " , A.DocNum [Pdo #], A.ItemCode [Product No.], B.ItemName [Product Description], ISNULL(A.PlannedQty,0) [Planned Qty] " _
                          & " , ISNULL(A.CmpltQty,0) [Complete Qty], A.PostDate [Pdo Order Date], A.DueDate [Pdo Due Date], (select DocEntry from ORDR where DocNum = A.OriginNum ) [SO DocEntry], A.OriginNum [SO #] " _
                          & " , A.CardCode [Customer Code], C.CardName [Customer Name]  " _
                      & " FROM OWOR A " _
                          & " INNER JOIN OITM B ON A.ItemCode = B.ItemCode " _
                          & " INNER JOIN OCRD C ON A.CardCode = C.CardCode " _
                      & " WHERE A.Status = 'R' " _
                          & " AND A.PostDate >= '" & Format(CDate(oForm.Items.Item("txtDate1").Specific.string), "yyyyMMdd") & "' " _
                          & " AND A.PostDate <= '" & Format(CDate(oForm.Items.Item("txtDate2").Specific.string), "yyyyMMdd") & "' "



        oForm.DataSources.DataTables.Item(0).ExecuteQuery(PCQuery)

        oForm.Items.Item("txtDate1").Click()



        RearrangePCGrid(oForm)


        oForm.Visible = True

    End Sub
    '-------------- Yadi FC ----------------------------


    '-------------- Yadi FC ----------------------------
    Sub RearrangePCGrid(ByVal oForm As SAPbouiCOM.Form)

        Dim oColumn As SAPbouiCOM.EditTextColumn
        'Dim oCheck As SAPbouiCOM.CheckBoxColumn
        'Dim idx As Long

        Dim oPCGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)

        oPCGrid = oForm.Items.Item("grdPC").Specific

        oPCGrid.RowHeaders.Width = 50

        'Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        oColumn = oPCGrid.Columns.Item("Pdo Entry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_ProductionOrder

        oColumn = oPCGrid.Columns.Item("SO DocEntry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order

        oColumn = oPCGrid.Columns.Item("Customer Code")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner


        oPCGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oPCGrid.Columns.Item("Check").TitleObject.Sortable = True

        oPCGrid.Columns.Item("#").Width = 30
        oPCGrid.Columns.Item("Check").Width = 40
        oPCGrid.Columns.Item("Pdo Entry").Width = 80
        oPCGrid.Columns.Item("Pdo #").Width = 80
        oPCGrid.Columns.Item("Product No.").Width = 80
        oPCGrid.Columns.Item("Product Description").Width = 150
        oPCGrid.Columns.Item("Planned Qty").Width = 60
        oPCGrid.Columns.Item("Complete Qty").Width = 60
        oPCGrid.Columns.Item("Pdo Order Date").Width = 80
        oPCGrid.Columns.Item("Pdo Due Date").Width = 80
        oPCGrid.Columns.Item("SO DocEntry").Width = 80
        oPCGrid.Columns.Item("SO #").Width = 80
        oPCGrid.Columns.Item("Customer Code").Width = 80
        oPCGrid.Columns.Item("Customer Name").Width = 150



        oPCGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto




        oPCGrid.Columns.Item("#").Editable = False
        oPCGrid.Columns.Item("Check").Editable = True
        oPCGrid.Columns.Item("Pdo Entry").Editable = False
        oPCGrid.Columns.Item("Pdo #").Editable = False
        oPCGrid.Columns.Item("Product No.").Editable = False
        oPCGrid.Columns.Item("Product Description").Editable = False
        oPCGrid.Columns.Item("Planned Qty").Editable = False
        oPCGrid.Columns.Item("Complete Qty").Editable = False
        oPCGrid.Columns.Item("Pdo Order Date").Editable = False
        oPCGrid.Columns.Item("Pdo Due Date").Editable = False
        oPCGrid.Columns.Item("SO DocEntry").Editable = False
        oPCGrid.Columns.Item("SO #").Editable = False
        oPCGrid.Columns.Item("Customer Code").Editable = False
        oPCGrid.Columns.Item("Customer Name").Editable = False

        Dim sboDate As String
        Dim dDate As DateTime

        'dDate = DateTime.Now

        'sbo formatdate
        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oPCGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub
    '-------------- Yadi FC ----------------------------
End Module
