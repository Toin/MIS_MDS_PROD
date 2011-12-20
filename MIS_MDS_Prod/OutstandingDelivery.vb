Module OutstandingDelivery

    Sub OutDelEntry()
        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText

        Dim oDelOutGrid As SAPbouiCOM.Grid

        Try
            'oForm = SBO_Application.Forms.Item("OutDel")
            'SBO_Application.MessageBox("Form Already Open")
            oForm = oApp.Forms.Item("OutDel")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            'fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "OutDel"
            fcp.UniqueID = "OutDel"
            fcp.XmlData = LoadFromXML("OutDel.srf")
            'oForm = SBO_Application.Forms.AddEx(fcp)
            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.DataTables.Add("DelOutLst")
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
            oItem = oForm.Items.Add("myGrid1", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 80
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            'oSOToMFGGrid = oForm.Items.Item("myGrid").Specific
            oDelOutGrid = oItem.Specific



            oDelOutGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

        End Try

    End Sub

End Module
