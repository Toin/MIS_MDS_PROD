Module MassUpdatePdOStatus

    'karno Production status
    Sub ProductionStatus()
        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText

        Dim oPDOStatusGrid As SAPbouiCOM.Grid

        Try
            'oForm = SBO_Application.Forms.Item("PDOStatus")
            'SBO_Application.MessageBox("Form Already Open")
            oForm = oApp.Forms.Item("PDOStatus")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            'fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "PDOStatus"
            fcp.UniqueID = "PDOStatus"
            fcp.XmlData = LoadFromXML("PDOStatus.srf")
            'oForm = SBO_Application.Forms.AddEx(fcp)
            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.DataTables.Add("PDOStatusLst")
            oForm.DataSources.UserDataSources.Add("TxtDtfrm", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("TxtDtTo", SAPbouiCOM.BoDataType.dt_DATE)


            'Default value for SO Date
            oForm.DataSources.UserDataSources.Item("TxtDtfrm").Value = oForm.Items.Item("TxtDtFrm").Specific.string

            oForm.DataSources.UserDataSources.Item("TxtDtTo").Value = oForm.Items.Item("TxtDtTo").Specific.string

            'Default setting
            ' add txtbox
            '        oEditText = oForm.Items.Add("SODate", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific


            oEditText = oForm.Items.Item("TxtDtFrm").Specific
            oEditText.DataBind.SetBound(True, "", "TxtDtFrm")
            oEditText = oForm.Items.Item("TxtDtTo").Specific
            oEditText.DataBind.SetBound(True, "", "TxtDtTo")

            '  add a GRID item to the form
            oItem = oForm.Items.Add("myGridPDO", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 80
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200



            'oSOToMFGGrid = oForm.Items.Item("myGrid").Specific
            oPDOStatusGrid = oItem.Specific

            oPDOStatusGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

        End Try
    End Sub

    Sub GeneratePdOStatus(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler

        Dim oPDOStatusGrid As SAPbouiCOM.Grid

        Dim idx As Long

        oPDOStatusGrid = oForm.Items.Item("myGridPDO").Specific

        'GRID - Order by column checkbox
        oPDOStatusGrid.Columns.Item("Release PdO").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        'Loop only selected/checked in grid rows and exit.
        For idx = oPDOStatusGrid.Rows.Count - 1 To 0 Step -1

            If oPDOStatusGrid.DataTable.GetValue(1, oPDOStatusGrid.GetDataTableRowIndex(idx)) = "Y" Then

                Dim oPDOStatus As SAPbobsCOM.ProductionOrders = Nothing

                Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
                'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines

                'Dim vCompany As SAPbobsCOM.Company = Nothing
                'Dim sCookie As String
                'Dim sConnectionContext As String

                'Dim isconnect As Long
                Dim errConnect As String = ""

                'Try
                '    vCompany = New SAPbobsCOM.Company
                '    'Dim sCookie As String = vCompany.GetContextCookie
                '    'Dim sConnectionContext As String
                '    sCookie = vCompany.GetContextCookie
                '    sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
                '    vCompany.SetSboLoginContext(sConnectionContext)
                '    isconnect = vCompany.Connect()

                '    'If vCompany.Connect() <> 0 Then
                '    If isconnect <> 0 Then
                '        End
                '    End If
                'Catch ex As Exception
                '    End
                'End Try

                'vCompany.StartTransaction()

                If Not oCompany.InTransaction Then
                    oCompany.StartTransaction()
                End If

                ''oProd1 = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                'oProd1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                ''oProd1.PlannedQuantity = 2

                'oProd1.ItemNo = oPDOStatusGrid.DataTable.GetValue(7, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                'oProd1.PlannedQuantity = oPDOStatusGrid.DataTable.GetValue(8, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                'oProd1.PostingDate = oPDOStatusGrid.DataTable.GetValue(1, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                'oProd1.DueDate = oPDOStatusGrid.DataTable.GetValue(1, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                'oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                'oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased

                ''oprod1.Warehouse = "01"
                ''oProd1.Warehouse = "FG-001"
                ''oProd1.Warehouse = oPDOStatusGrid.DataTable.GetValue(12, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                'oProd1.CustomerCode = oPDOStatusGrid.DataTable.GetValue(14, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)
                'oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
                '' so docnum
                'oProd1.ProductionOrderOriginEntry = oPDOStatusGrid.DataTable.GetValue(4, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                'oProdLine1 = oProd1.Lines

                'lRetCode = oProd1.Add()

                Dim PdOno As String = ""

                'oPDOStatus = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                oPDOStatus = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                'oPDOStatus.GetByKey(oPDOStatusGrid.DataTable.GetValue(2, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString))
                oPDOStatus.GetByKey(oPDOStatusGrid.DataTable.GetValue(3, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString))


                oPDOStatus.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                'oPDOStatus.UserFields.Fields.Item(12).Value = "12"
                oPDOStatus.UserFields.Fields.Item("U_MIS_Progress").Value = "Released"

                lRetCode = oPDOStatus.Update()


                If lRetCode <> 0 Then
                    'vCompany.GetLastError(lErrCode, sErrMsg)
                    'SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    oApp.MessageBox(lErrCode & ": " & sErrMsg)

                    If oCompany.InTransaction Then
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                Else

                    '' !!!! Make sure before create another object type-> clear previous/current object type.
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
                    'oProdLine1 = Nothing

                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
                    'oProd1 = Nothing


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPDOStatus)
                    oPDOStatus = Nothing

                    'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                    If oCompany.InTransaction Then
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If

                End If


                'vCompany.Disconnect()
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(vCompany)
                'vCompany = Nothing

                GC.Collect()
            Else
                Exit For
            End If
        Next

        'MsgBox("Begin trx: generating... PdO")
        'SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        oApp.SetStatusBarMessage("Generating PdO.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)


        'MsgBox("generating... PdO; DONE!!!")

        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

End Module
