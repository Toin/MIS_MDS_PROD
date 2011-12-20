Module OptimizationPdO

    Sub OptimizationEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim oCombobox As SAPbouiCOM.ComboBox = Nothing
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim lNextSeriesNumOptimization As Long

        Try
            'oForm = SBO_Application.Forms.Item("mds_p3")
            'SBO_Application.MessageBox("Form Already Open")
            oForm = oApp.Forms.Item("mds_p3")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            'fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds"
            fcp.UniqueID = "mds_p3"
            fcp.ObjectType = "MIS_OPTIM"
            fcp.XmlData = LoadFromXML("Optimization.srf")
            'oForm = SBO_Application.Forms.AddEx(fcp)
            oForm = oApp.Forms.AddEx(fcp)

            'oForm.DataBrowser.BrowseBy = "DocNum"

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("OptimDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("QtyLembar", SAPbouiCOM.BoDataType.dt_QUANTITY)

            'Default value for Optimization Date
            oForm.DataSources.UserDataSources.Item("OptimDate").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("QtyLembar").Value = 2

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.DBDataSources.Add("@MIS_OPTIM")

            oForm.DataSources.DBDataSources.Add("@MIS_OPTIML")

            'oForm.Items.Item("OptimDate").Width = 100
            'oEditText = oForm.Items.Item("OptimDate").Specific
            'oEditText.DataBind.SetBound(True, "", "OptimDate")


            oForm.DataSources.UserDataSources.Add("#", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)


            'Bind Data to Form

            'Combo Series UDO
            oItem = oForm.Items.Add("SeriesOptm", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = 720
            oItem.Top = 10

            'Fill data for combo series
            oCombobox = oItem.Specific
            oCombobox.ValidValues.LoadSeries("MIS_OPTIM", SAPbouiCOM.BoSeriesMode.sf_Add)
            'New Method
            oCombobox.DataBind.SetBound(True, "@MIS_OPTIM", "SERIES")


            'oItem = oForm.Items.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = 300
            'oItem.Top = 10

            oEditText = oForm.Items.Item("DocNum").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "DocNum")

            lNextSeriesNumOptimization = oForm.BusinessObject.GetNextSerialNumber("SERIES")
            oEditText = oForm.Items.Item("DocNum").Specific
            oEditText.String = lNextSeriesNumOptimization

            'oEditText = oForm.Items.Item("DocNum").Specific
            'oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "DocNum")

            oEditText = oForm.Items.Item("OptimDate").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_OptDate")

            oEditText = oForm.Items.Item("OptimRef").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_OptNum")

            oEditText = oForm.Items.Item("ItemCode").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_ItemCode")

            'Dim itemCFL As SBOConnection
            'itemCFL = New SBOConnection


            'itemCFL.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_Items, "ItemCFL1", "ItemCFL2", "ItemCode", _
            '                SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "")

            MIS_Toolkit.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_Items, "ItemCFL1", "ItemCFL2", "ItemCode", _
                            SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "")

            oEditText.ChooseFromListUID = "ItemCFL1"
            oButton = oForm.Items.Item("ItemButton").Specific
            oButton.ChooseFromListUID = "ItemCFL2"


            'itemCFL.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_Items, "ItemCFL3", "ItemCFL4", "ItemCode", _
            '    SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "")

            MIS_Toolkit.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_Items, "ItemCFL3", "ItemCFL4", "ItemCode", _
                SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "")


            oEditText = oForm.Items.Item("ItemKcSisa").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_ItemCdKacaSisa")
            oEditText.ChooseFromListUID = "ItemCFL3"
            oButton = oForm.Items.Item("ItmSisaBtn").Specific
            oButton.ChooseFromListUID = "ItemCFL4"

            'itemCFL = Nothing

            oEditText = oForm.Items.Item("Dscription").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_ItemDesc")

            oEditText = oForm.Items.Item("PnjangKaca").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_Pcm")

            oEditText = oForm.Items.Item("LebarKaca").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_Lcm")

            oEditText = oForm.Items.Item("QtyLembar").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_QtyinLembar")
            oEditText = oForm.Items.Item("LuasKaca").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_LuasM2")
            oEditText = oForm.Items.Item("SisaKcUtuh").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_KcSisaUtuh")
            oEditText = oForm.Items.Item("KacaPakai").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_KacaUsed")
            oEditText = oForm.Items.Item("TotalWaste").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_TotalWaste")
            oEditText = oForm.Items.Item("ByUser").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_User")


            'Set Matrix - add column from PdO & MIS_OPTIML
            'Dim oItem As SAPbouiCOM.Item
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim oColumns As SAPbouiCOM.Columns
            Dim oColumn As SAPbouiCOM.Column

            'oItem = oForm.Items.Item("OptimMtx").Specific
            oItem = oForm.Items.Item("OptimMtx")
            oItem.Width = 980
            oItem.Height = 200

            oForm.Items.Item("1").Top = oForm.ClientHeight - 60
            oForm.Items.Item("2").Top = oForm.ClientHeight - 60

            oForm.Items.Item("43").Top = oForm.ClientHeight - 60
            oForm.Items.Item("44").Top = oForm.ClientHeight - 40
            oForm.Items.Item("45").Top = oForm.ClientHeight - 20
            oForm.Items.Item("GTabc").Top = oForm.ClientHeight - 60
            oForm.Items.Item("GTaloc").Top = oForm.ClientHeight - 40
            oForm.Items.Item("GTplanPdO").Top = oForm.ClientHeight - 20


            oMatrix = oItem.Specific

            oColumns = oMatrix.Columns

            'Add Column to Matrix
            'oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oColumn.TitleObject.Caption = "#"
            'oColumn.Width = 20
            'oColumn.Editable = False

            oColumn = oColumns.Item("#")
            oColumn.TitleObject.Caption = "#"
            oColumn.Width = 30
            oColumn.DataBind.SetBound(True, , "#")
            oColumn.Editable = False

            oColumn = oColumns.Item("LineId")
            oColumn.TitleObject.Caption = "LineId"
            oColumn.Width = 40
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "LineId")


            'Dim mistoolbox As MISToolbox
            'mistoolbox = New MISToolbox
            'mistoolbox.AddChooseFromListForMatrix(oForm, SAPbouiCOM.BoLinkedObject.lf_ProductionOrder, "PdOCFL1", "Status", _
            '                   SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "L")

            'oColumn = oColumns.Add("PdOButton", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oColumn.Width = 20
            'oColumn.ChooseFromListUID = "PdOCFL2"
            'oButton.ChooseFromListUID = "PdOCFL2"

            oColumn = oColumns.Item("PdO#")
            oColumn.TitleObject.Caption = "Pdo No."
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_PdONum")


            oColumn = oColumns.Item("SO#")
            'oColumn = oColumns.Add("SO#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "SO Num"
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_SONum")


            oColumn = oColumns.Item("SOLine")
            oColumn.TitleObject.Caption = "SOLine"
            oColumn.Width = 40
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_SOLineNum")


            oColumn = oColumns.Item("CardCode")
            oColumn.TitleObject.Caption = "Cust. Code"
            oColumn.Width = 50
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_CardCode")

            'oColumn.ChooseFromListUID = "PdOCFL1"
            'oColumn.ChooseFromListAlias = "DocNum"

            'mistoolbox = Nothing



            oColumn = oColumns.Item("CardName")
            oColumn.TitleObject.Caption = "Customer Name"
            oColumn.Width = 120
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_CardName")


            oColumn = oColumns.Item("QtyPotong")
            oColumn.TitleObject.Caption = "Jumlah Potong"
            oColumn.Width = 60
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_QtyPotong")


            oColumn = oColumns.Item("P")
            oColumn.TitleObject.Caption = "Panjang"
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_Pcm")

            oColumn = oColumns.Item("L")
            oColumn.TitleObject.Caption = "Lebar"
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_Lcm")


            oColumn = oColumns.Item("TotalABC")
            oColumn.TitleObject.Caption = "Total A x B x C"
            oColumn.Width = 100
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_TotalABC")

            oColumn = oColumns.Item("AlocWaste")
            oColumn.TitleObject.Caption = "Allocated Waste"
            oColumn.Width = 100
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_AllocatedWaste")

            oColumn = oColumns.Item("PlanPdIsue")
            oColumn.TitleObject.Caption = "Plan PdO Issue"
            oColumn.Width = 100
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_QtPlanPdoIssue")



            oForm.DataBrowser.BrowseBy = "DocNum"
            'oForm.DataBrowser.BrowseBy = "U_MIS_ItemCode"



            oForm.EnableMenu("1292", True) 'Add Row
            oForm.EnableMenu("1293", True) 'Delete Row



            oForm.Freeze(False)


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing


            RearrangeFormOptimEntry(oForm)

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))



        End Try


        GC.Collect()

        RearrangeFormOptimEntry(oForm)
        'RearrangeGridOptimization(oForm)

        oForm.Items.Item("GTabc").Specific.value = 0
        oForm.Items.Item("GTaloc").Specific.value = 0
        oForm.Items.Item("GTplanPdO").Specific.value = 0

        oForm.Visible = True


    End Sub

    Sub RearrangeFormOptimEntry(ByVal oForm As SAPbouiCOM.Form)

        Dim oColumn As SAPbouiCOM.EditTextColumn

        Dim oOptimEntryMatrix As SAPbouiCOM.Matrix

        oForm.Freeze(True)

        oOptimEntryMatrix = oForm.Items.Item("OptimMtx").Specific

        'oOptimEntryMatrix.RowHeaders.Width = 50

        'oOptimEntryMatrix.Columns.Item("Cust. Code").Width = 130


        'oColumn = oOptimEntryMatrix.Columns.Item("Customer Name")
        'oColumn.Editable = False


        'oOptimEntryMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        '' Set Total Row count in colum title/header
        ''oOptimEntryMatrix.Columns.Item(0).TitleObject.Caption = oOptimEntryMatrix.Rows.Count.ToString


        'oOptimEntryMatrix.Columns.Item("#").Editable = False
        'oOptimEntryMatrix.Columns.Item(1).Editable = True
        'oOptimEntryMatrix.Columns.Item("SO Date").Editable = False
        'oOptimEntryMatrix.Columns.Item("DocEntry").Editable = False
        'oOptimEntryMatrix.Columns.Item("DocNum").Editable = False
        'oOptimEntryMatrix.Columns.Item("SO Line").Editable = False
        'oOptimEntryMatrix.Columns.Item("Sales Rep.").Editable = False
        'oOptimEntryMatrix.Columns.Item("Cust. Code").Editable = False
        'oOptimEntryMatrix.Columns.Item("FG").Editable = False
        'oOptimEntryMatrix.Columns.Item("FGName").Editable = False
        'oOptimEntryMatrix.Columns.Item("Quantity").Editable = False
        'oOptimEntryMatrix.Columns.Item("UOM").Editable = False
        'oOptimEntryMatrix.Columns.Item("Exp Delivery Date").Editable = False

        'oOptimEntryMatrix.Columns.Item("WhsCode").Editable = False
        'oOptimEntryMatrix.Columns.Item("PanjangInCm").Editable = False
        'oOptimEntryMatrix.Columns.Item("LebarInCm").Editable = False
        'oOptimEntryMatrix.Columns.Item("SO_Bentuk").Editable = False

        'oOptimEntryMatrix.Columns.Item("SO Approval Status").Editable = False
        'oOptimEntryMatrix.Columns.Item("SO LineNum").Editable = False


        'oOptimEntryMatrix.RowHeaders.Width = 20
        'oOptimEntryMatrix.Columns.Item("#").Width = 30
        'oOptimEntryMatrix.Columns.Item(1).Width = 20
        'oOptimEntryMatrix.Columns.Item("SO Date").Width = 60
        'oOptimEntryMatrix.Columns.Item("DocEntry").Width = 60
        'oOptimEntryMatrix.Columns.Item("DocNum").Width = 60
        'oOptimEntryMatrix.Columns.Item("SO Line").Width = 30
        'oOptimEntryMatrix.Columns.Item("Cust. Code").Width = 80
        'oOptimEntryMatrix.Columns.Item("FG").Width = 100
        'oOptimEntryMatrix.Columns.Item("Exp Delivery Date").Width = 80
        'oOptimEntryMatrix.Columns.Item("WhsCode").Width = 50
        'oOptimEntryMatrix.Columns.Item("PanjangInCm").Width = 50
        'oOptimEntryMatrix.Columns.Item("LebarInCm").Width = 50
        'oOptimEntryMatrix.Columns.Item("SO_Bentuk").Width = 80


        'Dim oEditText As SAPbouiCOM.EditText
        'Dim oStaticText As SAPbouiCOM.StaticText
        'Dim oButton As SAPbouiCOM.Button

        Dim oItem As SAPbouiCOM.Item

        'oItem = oForm.Items.Item("OptimMtx").Specific
        oItem = oForm.Items.Item("OptimMtx")
        'oItem.Height = 200
        oItem.Top = 135
        oItem.Height = oForm.ClientHeight - 200
        oItem.Width = oForm.ClientWidth - 20

        oForm.Items.Item("1").Top = oForm.ClientHeight - 60
        oForm.Items.Item("2").Top = oForm.ClientHeight - 60

        oForm.Items.Item("43").Top = oForm.ClientHeight - 60
        oForm.Items.Item("44").Top = oForm.ClientHeight - 40
        oForm.Items.Item("45").Top = oForm.ClientHeight - 20

        oForm.Items.Item("GTabc").Top = oForm.ClientHeight - 60
        oForm.Items.Item("GTaloc").Top = oForm.ClientHeight - 40
        oForm.Items.Item("GTplanPdO").Top = oForm.ClientHeight - 20

        oForm.Items.Item("43").Left = oForm.ClientWidth - 180
        oForm.Items.Item("44").Left = oForm.ClientWidth - 180
        oForm.Items.Item("45").Left = oForm.ClientWidth - 180

        oForm.Items.Item("GTabc").Left = oForm.ClientWidth - 90
        oForm.Items.Item("GTaloc").Left = oForm.ClientWidth - 90
        oForm.Items.Item("GTplanPdO").Left = oForm.ClientWidth - 90

        oForm.Items.Item("SeriesOptm").Left = oForm.ClientWidth - 280
        oForm.Items.Item("NoOptim").Left = oForm.ClientWidth - 180
        oForm.Items.Item("46").Left = oForm.ClientWidth - 180
        oForm.Items.Item("OptDatelbl").Left = oForm.ClientWidth - 180
        oForm.Items.Item("LuasKcLbl").Left = oForm.ClientWidth - 180
        oForm.Items.Item("qtycm2").Left = oForm.ClientWidth - 120
        oForm.Items.Item("34").Left = oForm.ClientWidth - 180

        oForm.Items.Item("DocNum").Left = oForm.ClientWidth - 90
        oForm.Items.Item("OptimRef").Left = oForm.ClientWidth - 90
        oForm.Items.Item("OptimDate").Left = oForm.ClientWidth - 90
        oForm.Items.Item("LuasKaca").Left = oForm.ClientWidth - 90
        oForm.Items.Item("ByUser").Left = oForm.ClientWidth - 90


        Dim sboDate As String
        Dim dDate As DateTime

        'dDate = DateTime.Now

        'sbo formatdate
        'sboDate = oMIS_Utils.fctFormatDate(dDate, oCompany)
        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oOptimEntryMatrix = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

End Module
