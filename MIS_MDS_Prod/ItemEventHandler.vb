Module ItemEventHandler

    Public WithEvents oApp4ItemEvent As SAPbouiCOM.Application = Nothing

    'Const ProductionIssue_MenuId As String = "4371"
    Const ProductionIssue_FormId As String = "65213"
    Const ProductionIssueUDF_FormId As String = "-65213"
    Dim objFormProductionIssue As SAPbouiCOM.Form
    Dim objFormProductionIssueUDF As SAPbouiCOM.Form
    'Dim intRowProductionIssueDetail As Integer
    ''karno 
    '' Production Issue
    'Const Production_MenuId As String = "4369"
    Const Production_FormId As String = "65211"
    Const ProductionUDF_FormId As String = "-65211"
    Dim objFormProduction As SAPbouiCOM.Form
    Dim objFormProductionUDF As SAPbouiCOM.Form
    'Dim intRowProductionDetail As Integer


    Sub ItemEventHandler(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                     ByRef BubbleEvent As Boolean) Handles oApp4ItemEvent.ItemEvent
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try

            If pVal.BeforeAction = False Then

                If pVal.FormTypeEx = "ListOptimize" Then

                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        'Dim oForm As SAPbouiCOM.Form = Nothing
                        'oForm = SBO_Application.Forms.Item(pVal.FormUID)
                        oForm = oApp.Forms.Item(pVal.FormUID)
                    End If

                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            Select Case pVal.ItemUID
                                Case "OptimNo"
                                    'Dim oForm As SAPbouiCOM.Form = Nothing
                                    'oForm = SBO_Application.Forms.Item("ListOptimize")
                                    oForm = oApp.Forms.Item("ListOptimize")

                                    Dim ListOptimizeQuery As String

                                    If oForm.Items.Item("OptimNo").Specific.string = "" Then
                                        ListOptimizeQuery = "select T0.DocNum [Number Optimize], T2.DocNum [PDO Number], T3.Visorder + 1 [Row No], T3.ItemCode [Item Code],  " & _
                                                            "T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                                            "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                                            "From [@MIS_OPTIM] T0 " & _
                                                            "join [@MIS_OPTIML] T1 " & _
                                                            "ON T0.DocEntry = T1.DocEntry " & _
                                                            "INNER JOIN OWOR T2 " & _
                                                            "ON T2.DocNum = T1.U_MIS_PdONum " & _
                                                            "INNER JOIN WOR1 T3 " & _
                                                            "ON T2.Docentry = T3.Docentry " & _
                                                            "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X'"

                                    Else
                                        ListOptimizeQuery = "select T0.DocNum [Number Optimize], T2.DocNum [PDO Number], T3.Visorder + 1 [Row No], T3.ItemCode [Item Code],  " & _
                                                            "T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                                            "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                                            "From [@MIS_OPTIM] T0 " & _
                                                            "join [@MIS_OPTIML] T1 " & _
                                                            "ON T0.DocEntry = T1.DocEntry " & _
                                                            "INNER JOIN OWOR T2 " & _
                                                            "ON T2.DocNum = T1.U_MIS_PdONum " & _
                                                            "INNER JOIN WOR1 T3 " & _
                                                            "ON T2.Docentry = T3.Docentry " & _
                                                            "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' AND T0.DocNum LIKE '" & oForm.Items.Item("OptimNo").Specific.string & "%'"
                                    End If

                                    oForm.DataSources.DataTables.Item("ListOptim").ExecuteQuery(ListOptimizeQuery)
                                    'oListOptimGrid.DataTable = oForm.DataSources.DataTables.Item("ListOptim")
                            End Select


                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            If pVal.ItemUID = "BtnChoose" Then
                                'Dim oForm As SAPbouiCOM.Form = Nothing
                                'oForm = SBO_Application.Forms.Item("ListOptimize")
                                oForm = oApp.Forms.Item("ListOptimize")
                                Dim oGrid As SAPbouiCOM.Grid = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing
                                Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                                Dim StrSql As String
                                Dim DocNum As String
                                Dim Row As Integer

                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                                oMatrix = objFormProductionIssue.Items.Item("13").Specific
                                oColumns = oMatrix.Columns
                                'karno not yet

                                oGrid = oForm.Items.Item("myGrid2").Specific


                                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 0 To oGrid.Rows.Count - 1
                                    'For i As Integer = 0 To oGrid.Rows.SelectedRows.Count

                                    If oGrid.Rows.IsSelected(i) = True Then
                                        DocNum = oGrid.DataTable.GetValue("PDO Number", oGrid.GetDataTableRowIndex(i)).ToString

                                        'StrSql = "SELECT T0.DocNum DocNum, T0.U_MIS_OPTNUM OrderNo, T0.U_MIS_QtyInLembar Qty FROM [@MIS_OPTIM] T0 INNER JOIN [@MIS_OPTIM] T1 ON T0.DocEntry = T1.DocEntry " & _
                                        '        "INNER JOIN OWOR T2 ON T0.U_MIS_OPTNUM = T2.DocNum WHERE T0.U_MIS_OPTNUM = '" & DocNum & "'"

                                        'StrSql = "select T0.DocNum, T2.Visorder + 1 RowNo, T2.itemcode, T0.U_MIS_QtyInLembar Qty, T0.U_MIS_OPTNUM OrderNo,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                        '                    "From [@MIS_OPTIM] T0 INNER JOIN OWOR T1 ON T0.U_MIS_OPTNUM = T1.Docnum INNER JOIN WOR1 T2 ON T1.Docentry = T2.Docentry " & _
                                        '                    "where T1.STATUS = 'R' AND LEFT(T2.ItemCode,1) <> 'X' AND T0.U_MIS_OPTNUM = '" & DocNum & "' "

                                        StrSql = "select T2.DocNum OrderNo, T3.linenum + 1  RowNo, T3.ItemCode, T1.U_MIS_QtPlanPdoIssue Qty,  " & _
                                                "T0.U_MIS_QtyInLembar Lembar, T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                                "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                                "From [@MIS_OPTIM] T0 " & _
                                                "join [@MIS_OPTIML] T1 " & _
                                                "ON T0.DocEntry = T1.DocEntry " & _
                                                "INNER JOIN OWOR T2 " & _
                                                "ON T2.DocNum = T1.U_MIS_PdONum " & _
                                                "INNER JOIN WOR1 T3 " & _
                                                "ON T2.Docentry = T3.Docentry " & _
                                                "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' AND T2.DocNum = '" & DocNum & "' "




                                        objRecSet.DoQuery(StrSql)

                                        If objRecSet.RecordCount > 0 Then
                                            For Row = 1 To objRecSet.RecordCount
                                                If objRecSet.Fields.Item("Qty").Value = 0.0 Then
                                                    'SBO_Application.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oApp.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oForm.Close()
                                                    Exit Sub
                                                Else
                                                    oColumns.Item("61").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("OrderNo").Value
                                                    oColumns.Item("60").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("RowNo").Value
                                                    oColumns.Item("9").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Qty").Value
                                                End If

                                            Next
                                        End If

                                    End If
                                Next

                                oForm.Close()

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                            End If
                    End Select

                End If

                ' karno optimize production
                If pVal.FormTypeEx = "ListOptimizePro" Then

                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        'Dim oForm As SAPbouiCOM.Form = Nothing
                        'oForm = SBO_Application.Forms.Item(pVal.FormUID)
                        oForm = oApp.Forms.Item(pVal.FormUID)
                    End If

                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            If pVal.ItemUID = "BtnChoose" Then
                                'Dim oForm As SAPbouiCOM.Form = Nothing
                                'oForm = SBO_Application.Forms.Item("ListOptimizePro")
                                oForm = oApp.Forms.Item("ListOptimizePro")

                                Dim oGrid As SAPbouiCOM.Grid = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing
                                Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                                Dim StrSql As String
                                Dim DocNum As String
                                Dim Row As Integer

                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                                oMatrix = objFormProduction.Items.Item("37").Specific
                                oColumns = oMatrix.Columns
                                'karno not yet

                                oGrid = oForm.Items.Item("myGrid2").Specific


                                For i As Integer = 0 To oGrid.Rows.SelectedRows.Count
                                    If oGrid.Rows.IsSelected(i) = True Then
                                        DocNum = oGrid.DataTable.GetValue("U_MIS_OPTNUM", oGrid.GetDataTableRowIndex(i)).ToString

                                        StrSql = "SELECT T0.U_MIS_ItemCode Item, T0.DocNum DocNum, T0.U_MIS_OPTNUM OrderNo, T0.U_MIS_QtyInLembar Qty FROM [@MIS_OPTIM] T0 INNER JOIN [@MIS_OPTIM] T1 ON T0.DocEntry = T1.DocEntry " & _
                                                "INNER JOIN OWOR T2 ON T0.U_MIS_OPTNUM = T2.DocNum WHERE T0.U_MIS_OPTNUM = '" & DocNum & "'"
                                        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        objRecSet.DoQuery(StrSql)

                                        If objRecSet.RecordCount > 0 Then
                                            For Row = 1 To objRecSet.RecordCount

                                                oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("Item").Value
                                            Next
                                        End If
                                    End If
                                Next

                                oForm.Close()

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                            End If
                    End Select

                End If

                If pVal.FormTypeEx = ProductionUDF_FormId Then
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        'objFormProductionUDF = SBO_Application.Forms.Item(pVal.FormUID)
                        objFormProductionUDF = oApp.Forms.Item(pVal.FormUID)
                    End If
                End If


                'If pVal.FormTypeEx = "mds_pdo1" Then
                '    'MsgBox("mds pdo1 popup")

                '    Select Case pVal.EventType
                '        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                '            If pVal.ItemUID = "btnFillOH" Then
                '                'Dim oForm As SAPbouiCOM.Form = Nothing
                '                Dim oFormPdO As SAPbouiCOM.Form = Nothing


                '                Dim oMatrix As SAPbouiCOM.Matrix
                '                Dim oColumns As SAPbouiCOM.Columns

                '                Dim itemcodeX As String
                '                Dim runtimeQty As Double

                '                ''oForm = SBO_Application.Forms.Item("mds_pdo1")
                '                'oForm = SBO_Application.Forms.Item(pVal.FormUID)
                '                oForm = oApp.Forms.Item(pVal.FormUID)
                '                itemcodeX = oForm.Items.Item("ItemCodeX").Specific.string
                '                runtimeQty = oForm.Items.Item("RunTimeQty").Specific.string

                '                oForm.Close()

                '                'oFormPdO = SBO_Application.Forms.Item(Production_FormId)
                '                'oMatrix = oFormPdO.Items.Item("37").Specific

                '                'oFormPdO = objFormProduction
                '                'oFormPdO = SBO_Application.Forms.Item(Production_FormId)
                '                'oMatrix = oFormPdO.Items.Item("37").Specific
                '                oMatrix = objFormProduction.Items.Item("37").Specific
                '                oColumns = oMatrix.Columns


                '                MsgBox("button Fill OH: " & pVal.ItemUID & "; itemcode X: " & itemcodeX & "; Runtime = " & runtimeQty)


                '                'oColumns.Item("4").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value                        End If
                '                oColumns.Item("37").Cells.Item(pVal.Row).Specific.value = itemcodeX


                '                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                '                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                '                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormPdO)
                '                System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)

                '            End If
                '    End Select

                'End If

                ' karno Production
                If pVal.FormTypeEx = Production_FormId Then
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        'objFormProduction = SBO_Application.Forms.Item(pVal.FormUID)
                        objFormProduction = oApp.Forms.Item(pVal.FormUID)
                    End If


                    Select Case pVal.EventType

                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            Select Case pVal.ColUID
                                Case "U_NBS_MatlQty"
                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim oColumns As SAPbouiCOM.Columns = Nothing

                                    oMatrix = objFormProduction.Items.Item("37").Specific
                                    oColumns = oMatrix.Columns

                                    If oColumns.Item("U_MIS_OptNum").Cells.Item(pVal.Row).Specific.value = "" Then
                                        If oColumns.Item("4").Cells.Item(pVal.Row).Specific.value = "" Then
                                            'SBO_Application.SetStatusBarMessage("Please Fill Item Code!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oApp.SetStatusBarMessage("Please Fill Item Code!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        Else

                                            If oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                'SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oApp.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            ElseIf oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value = "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value = "0.0" Then
                                                'SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oApp.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                'SBO_Application.SetStatusBarMessage("Material Quantity Must Blank AND Run Time Must Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oApp.SetStatusBarMessage("Material Quantity Must Blank AND Run Time Must Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                                Exit Sub
                                            Else
                                                If Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) <> "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                    oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                                Else
                                                    'SBO_Application.SetStatusBarMessage("Material Quantity Must Fill AND Run Time Must Blank!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oApp.SetStatusBarMessage("Material Quantity Must Fill AND Run Time Must Blank!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)

                                Case "U_NBS_RunTime"
                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim oColumns As SAPbouiCOM.Columns = Nothing

                                    oMatrix = objFormProduction.Items.Item("37").Specific
                                    oColumns = oMatrix.Columns

                                    If oColumns.Item("U_MIS_OptNum").Cells.Item(pVal.Row).Specific.value = "" Then
                                        If oColumns.Item("4").Cells.Item(pVal.Row).Specific.value = "" Then
                                            'SBO_Application.SetStatusBarMessage("Please Fill Item Code!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oApp.SetStatusBarMessage("Please Fill Item Code!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub

                                        Else
                                            If oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                'SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oApp.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            ElseIf oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value = "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value = "0.0" Then
                                                'SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oApp.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                'SBO_Application.SetStatusBarMessage("Material Quantity Must Blank AND Run Time Must Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oApp.SetStatusBarMessage("Material Quantity Must Blank AND Run Time Must Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                                Exit Sub
                                            Else
                                                If Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) <> "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                    oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                                Else
                                                    'SBO_Application.SetStatusBarMessage("Material Quantity Must Fill AND Run Time Must Blank!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oApp.SetStatusBarMessage("Material Quantity Must Fill AND Run Time Must Blank!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)



                            End Select

                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Dim oItem As SAPbouiCOM.Item

                            oItem = objFormProduction.Items.Add("BtnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                            oItem.Left = 145
                            oItem.Top = 395
                            oItem.Width = 150
                            oItem.Height = 19
                            oItem.Specific.caption = "Copy From Optimize"

                            oItem = objFormProduction.Items.Add("BtnAddMch", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                            oItem.Left = 300
                            oItem.Top = 395
                            oItem.Width = 150
                            oItem.Height = 19
                            oItem.Specific.caption = "Add Mesin+Runtime"


                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "BtnAddMch" Then
                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing
                                Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                                Dim PdOStatus As String
                                Dim PlannedQty As Double
                                Dim StrSql As String

                                PdOStatus = objFormProduction.Items.Item("10").Specific.value
                                PlannedQty = objFormProduction.Items.Item("12").Specific.value

                                If PdOStatus <> "P" Then
                                    objFormProduction.Items.Item("BtnAddMch").Enabled = False
                                End If

                                '???
                                'objFormProduction.Items.Item("BtnAddMch").Enabled = True

                                oMatrix = objFormProduction.Items.Item("37").Specific
                                oColumns = oMatrix.Columns


                                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'StrSql = "Select T0.ItemCode ItemCode " & _
                                '        "FROM OITM T0 " & _
                                '        "WHERE T0.ItemCode like '%" & objFormProductionUDF.Items.Item("U_MIS_MachineCode01").Specific.value() & "%' "

                                'objRecSet.DoQuery(StrSql)

                                'If objRecSet.RecordCount > 0 Then
                                '    objRecSet.MoveFirst()
                                '    For Row = 1 To objRecSet.RecordCount
                                '        If objRecSet.Fields.Item("ItemCode").Value = "" Then
                                '            'SBO_Application.SetStatusBarMessage("ItemCode is mandatory, fill first 3digit Item Code! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '            oApp.SetStatusBarMessage("ItemCode is mandatory, fill first 3digit Item Code! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '            objFormProductionIssue.Close()
                                '            Exit Sub
                                '        Else

                                '            'oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineCode").Specific.value() + "-DL"
                                '            oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                                '            oColumns.Item("U_NBS_RunTime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachRuntime01").Specific.value()

                                '            'objRecSet.Fields.Item("ItemCode").Value
                                '        End If
                                '        objRecSet.MoveNext()
                                '    Next
                                'Else
                                '    'SBO_Application.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '    oApp.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '    Exit Sub
                                'End If


                                ''Machine-02 + Runtime
                                'StrSql = "Select T0.ItemCode ItemCode " & _
                                '        "FROM OITM T0 " & _
                                '        "WHERE T0.ItemCode like '%" & objFormProductionUDF.Items.Item("U_MIS_MachineCode02").Specific.value() & "%' "

                                'objRecSet.DoQuery(StrSql)

                                'If objRecSet.RecordCount > 0 Then
                                '    objRecSet.MoveFirst()
                                '    For Row = 1 To objRecSet.RecordCount
                                '        If objRecSet.Fields.Item("ItemCode").Value = "" Then
                                '            'SBO_Application.SetStatusBarMessage("ItemCode is mandatory, fill first 3digit Item Code! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '            oApp.SetStatusBarMessage("ItemCode is mandatory, fill first 3digit Item Code! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '            objFormProductionIssue.Close()
                                '            Exit Sub
                                '        Else

                                '            'oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineCode").Specific.value() + "-DL"
                                '            oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                                '            oColumns.Item("U_NBS_RunTime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachRuntime02").Specific.value()

                                '            'objRecSet.Fields.Item("ItemCode").Value
                                '        End If
                                '        objRecSet.MoveNext()
                                '    Next
                                'Else
                                '    'SBO_Application.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '    oApp.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '    Exit Sub
                                'End If


                                'Dim idx As Integer
                                '???? otomatic fill Machine + Runtime total 5 set!!!
                                For idx = 1 To 5
                                    If objFormProductionUDF.Items.Item("U_MIS_MachineCode" + idx.ToString("00")).Specific.value() <> "" Then
                                        StrSql = "Select T0.ItemCode ItemCode " & _
                                            "FROM OITM T0 " & _
                                            "WHERE T0.ItemCode like '%" & _
                                                objFormProductionUDF.Items.Item("U_MIS_MachineCode" _
                                                                    + idx.ToString("00")).Specific.value() & "%' "

                                        objRecSet.DoQuery(StrSql)

                                        If objRecSet.RecordCount > 0 Then
                                            objRecSet.MoveFirst()
                                            For Row = 1 To objRecSet.RecordCount
                                                If objRecSet.Fields.Item("ItemCode").Value = "" Then
                                                    'SBO_Application.SetStatusBarMessage("ItemCode is mandatory, fill first 3digit Item Code! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oApp.SetStatusBarMessage("ItemCode is mandatory, fill first 3digit Item Code! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    objFormProductionIssue.Close()
                                                    Exit Sub
                                                Else

                                                    'oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineCode").Specific.value() + "-DL"
                                                    oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                                                    'oColumns.Item("U_NBS_RunTime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachRuntime01").Specific.value()
                                                    oColumns.Item("U_NBS_RunTime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachRuntime" + idx.ToString("00")).Specific.value()

                                                    'objRecSet.Fields.Item("ItemCode").Value
                                                End If
                                                objRecSet.MoveNext()
                                            Next
                                        Else
                                            'SBO_Application.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oApp.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        End If

                                    End If

                                Next

                                ''BubbleEvent = False
                                ''oMatrix.AddRow(1, -1)
                                'oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineCode").Specific.value() + "-DL"
                                ''oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineRuntime").Specific.value()
                                ''oColumns.Item("U_NBS_Runtime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = 8 'objFormProductionUDF.Items.Item("U_MIS_MachineRuntime").Specific.value()

                                ''oMatrix.AddRow(1, -1)
                                'oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineCode").Specific.value() + "-FOH"
                                ''oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineRuntime").Specific.value()
                                ''oColumns.Item("U_NBS_Runtime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = 8 'objFormProductionUDF.Items.Item("U_MIS_MachineRuntime").Specific.value()

                                ''oMatrix.AddRow(1, -1)
                                'oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineCode").Specific.value() + "-IL"
                                ''oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = 5
                                'oColumns.Item("U_NBS_RunTime").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objFormProductionUDF.Items.Item("U_MIS_MachineRuntime").Specific.value()

                                ''oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                                ''oColumns.Item("14").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("PlannedQty").Value
                                ''oColumns.Item("U_MIS_PdOGenFlag").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Flag").Value
                                ''oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = Math.Round(objRecSet.Fields.Item("PlannedQty").Value / PlannedQty, 4)
                                ''oColumns.Item("U_NBS_Runt").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = Math.Round(objRecSet.Fields.Item("PlannedQty").Value / PlannedQty, 4)

                                ' ''Reassignment Again! Dari Optimize data harus masuk dulu ke Planned Qty baru nanti kalkulasi dptkan nilai material qty. Jadi Optimize Planned qty = PdO Planned Qty 
                                ''oColumns.Item("14").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("PlannedQty").Value
                                ''oColumns.Item("U_MIS_OptNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("OptimNumber").Value

                                '????

                                'System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProduction)

                                'System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProductionUDF)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)

                            End If

                            If pVal.ItemUID = "BtnCopy" Then
                                Dim StrSql As String
                                'Dim DocNum As String
                                Dim PlannedQty As Double
                                Dim PdoNumber As String


                                PdoNumber = objFormProduction.Items.Item("18").Specific.value
                                'DocNum = objFormProductionUDF.Items.Item("U_MIS_OptNum").Specific.string
                                PlannedQty = objFormProduction.Items.Item("12").Specific.value

                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing
                                Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                                oMatrix = objFormProduction.Items.Item("37").Specific
                                oColumns = oMatrix.Columns
                                'karno not yet

                                'oGrid = oForm.Items.Item("myGrid2").Specific
                                '201100146 1011001880

                                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'StrSql = "Select T0.U_MIS_ItemCode ItemCode, " & _
                                '        "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyinLembar	PlannedQty, " & _
                                '        "'Y' Flag " & _
                                '        "FROM [@MIS_OPTIM] T0 " & _
                                '        "JOIN [@MIS_OPTIML] T1 " & _
                                '        "ON T0.docentry = T1.docentry " & _
                                '        "WHERE T0.DocNum = '" & DocNum & "' " & _
                                '        "AND T1.U_MIS_PdONum = " & PdoNumber & " "

                                StrSql = "Select T0.U_MIS_ItemCode ItemCode, " & _
                                        "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyinLembar	PlannedQty, " & _
                                        "'Y' Flag, T0.DocNum OptimNumber " & _
                                        "FROM [@MIS_OPTIM] T0 " & _
                                        "JOIN [@MIS_OPTIML] T1 " & _
                                        "ON T0.docentry = T1.docentry " & _
                                        "WHERE T1.U_MIS_PdONum = " & PdoNumber & " "

                                objRecSet.DoQuery(StrSql)

                                If objRecSet.RecordCount > 0 Then
                                    objRecSet.MoveFirst()
                                    For Row = 1 To objRecSet.RecordCount
                                        If objRecSet.Fields.Item("PlannedQty").Value = 0.0 Then
                                            'SBO_Application.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oApp.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            objFormProductionIssue.Close()
                                            Exit Sub
                                        Else
                                            'oColumns.Item("U_MIS_OptNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("OptimNumber").Value
                                            'oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = Math.Round(objRecSet.Fields.Item("PlannedQty").Value / PlannedQty, 4)
                                            oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                                            oColumns.Item("14").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("PlannedQty").Value
                                            oColumns.Item("U_MIS_PdOGenFlag").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Flag").Value
                                            oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = Math.Round(objRecSet.Fields.Item("PlannedQty").Value / PlannedQty, 4)

                                            'Reassignment Again! Dari Optimize data harus masuk dulu ke Planned Qty baru nanti kalkulasi dptkan nilai material qty. Jadi Optimize Planned qty = PdO Planned Qty 
                                            oColumns.Item("14").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("PlannedQty").Value
                                            oColumns.Item("U_MIS_OptNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("OptimNumber").Value
                                        End If
                                        objRecSet.MoveNext()
                                    Next
                                Else
                                    'SBO_Application.SetStatusBarMessage("Please Check Optimazation Number", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    oApp.SetStatusBarMessage("Please Check Optimazation Number", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Exit Sub
                                End If

                                'End If
                                'Next

                                'objFormProductionIssue.Close()

                                objFormProduction.Items.Item("BtnCopy").Enabled = False

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProduction)
                                'System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProductionUDF)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)

                            End If
                    End Select
                End If
                'karno Copy Optim
                If pVal.FormTypeEx = ProductionIssueUDF_FormId Then
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        'objFormProductionIssueUDF = SBO_Application.Forms.Item(pVal.FormUID)
                        objFormProductionIssueUDF = oApp.Forms.Item(pVal.FormUID)
                    End If
                End If

                If pVal.FormTypeEx = ProductionIssue_FormId Then
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        'objFormProductionIssue = SBO_Application.Forms.Item(pVal.FormUID)
                        objFormProductionIssue = oApp.Forms.Item(pVal.FormUID)
                    End If

                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Dim oItem As SAPbouiCOM.Item
                            'Dim oListOptimizeGrid As SAPbouiCOM.Grid = Nothing

                            oItem = objFormProductionIssue.Items.Add("BtnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                            oItem.Left = 145
                            oItem.Top = 318
                            oItem.Width = 150
                            oItem.Height = 19
                            oItem.Specific.caption = "Copy From Optimize"

                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oListOptimizeGrid)

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "BtnCopy" Then
                                Dim StrSql As String
                                Dim DocNum As String = ""
                                Dim OptimLembar As Integer
                                Dim Row As Integer

                                If objFormProductionIssueUDF.Items.Item("U_MIS_OptNum").Specific.string = "" Then
                                    'SBO_Application.SetStatusBarMessage("Optimize Number Must Fill", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    oApp.SetStatusBarMessage("Optimize Number Must Fill", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Exit Sub
                                Else
                                    If objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.value = "" Then
                                        objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.value = 1
                                        OptimLembar = objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.value
                                    Else
                                        OptimLembar = objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.string
                                        DocNum = objFormProductionIssueUDF.Items.Item("U_MIS_OptNum").Specific.string
                                    End If
                                End If

                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing
                                Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                                oMatrix = objFormProductionIssue.Items.Item("13").Specific
                                oColumns = oMatrix.Columns

                                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                StrSql = "select T2.DocNum OrderNo, T3.linenum + 1 RowNo, T3.ItemCode, " & _
                                "CASE T0.U_MIS_ItemCode WHEN T3.ItemCode THEN T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & " " & _
                                "ELSE (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty END Qty,  T0.U_MIS_QtyInLembar Lembar, " & _
                                "T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                                "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar MaximumQtyLembarPdoIssue, " & _
                                "SUM(ISNULL(T4.Quantity,0)) QuantityPdoIssue " & _
                                "From [@MIS_OPTIM] T0 join [@MIS_OPTIML] T1  " & _
                                "ON T0.DocEntry = T1.DocEntry INNER JOIN OWOR T2  " & _
                                "ON T2.DocNum = T1.U_MIS_PdONum INNER JOIN WOR1 T3  " & _
                                "ON T2.Docentry = T3.Docentry  AND T0.DocNum = T3.U_MIS_OptNum " & _
                                "AND ROUND((1 / T0.U_MIS_QtyInLembar) * T3.PlannedQty, 4) = ROUND(T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar / T0.U_MIS_QtyInLembar, 4) " & _
                                "LEFT JOIN IGE1 T4 ON T2.DocNum = T4.BaseRef " & _
                                "AND T3.ItemCode = T4.ItemCode " & _
                                "AND T3.LineNum = T4.BaseLine " & _
                                "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' " & _
                                "AND T0.DocNum = '" & DocNum & "' " & _
                                "GROUP BY T2.DocNum, T3.linenum + 1, T3.ItemCode, " & _
                                "T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ", (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty,  T0.U_MIS_QtyInLembar,  " & _
                                "T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                                "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar " & _
                                "HAVING (T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ") + SUM(ISNULL(T4.Quantity,0)) <= T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar "

                                objRecSet.DoQuery(StrSql)


                                If objRecSet.RecordCount > 0 Then
                                    objRecSet.MoveFirst()
                                    For Row = 1 To objRecSet.RecordCount
                                        If objRecSet.Fields.Item("Qty").Value = 0.0 Then
                                            'SBO_Application.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oApp.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            objFormProductionIssue.Close()
                                            Exit Sub
                                        Else
                                            oColumns.Item("61").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("OrderNo").Value
                                            oColumns.Item("60").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("RowNo").Value
                                            oColumns.Item("9").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Qty").Value
                                        End If
                                        objRecSet.MoveNext()
                                    Next
                                Else
                                    'SBO_Application.SetStatusBarMessage("Please Check ItemCode, Planned Qty Production Order Not Same With Optimization Or Production Order Status Not Release", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    oApp.SetStatusBarMessage("Please Check ItemCode, Planned Qty Production Order Not Same With Optimization Or Production Order Status Not Release", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Exit Sub
                                End If

                                'End If
                                'Next

                                'objFormProductionIssue.Close()
                                objFormProductionIssue.Items.Item("BtnCopy").Enabled = False
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProductionIssue)
                                'System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                            End If

                    End Select
                End If


                Select Case FormUID
                    ' karno Prodution Status
                    Case "PDOStatus"
                        If pVal.ItemUID = "BtnRelease" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oPDOStatusGrid As SAPbouiCOM.Grid

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("PDOStatusLst")

                            oPDOStatusGrid = oForm.Items.Item("myGridPDO").Specific


                            GeneratePdOStatus(oForm)

                        End If

                        If ((pVal.ItemUID = "BtnShow") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)) Then

                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            Dim PDOStatusQuery As String
                            Dim oPDOStatusGrid As SAPbouiCOM.Grid = Nothing
                            Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

                            'Dim oMis_Utils As MIS_Utils

                            'oMis_Utils = New MIS_Utils

                            oPDOStatusGrid = oForm.Items.Item("myGridPDO").Specific



                            'DelOutQuery = "select T0.DocEntry DocEntry, T0.DocNum SoDocNum, T0.DocDate Sodate, '' SalesRep, T0.CardCode SoCustCode, " & _
                            ' "T0.CardName SOCustName, T0.TrnspCode ShippingType, T1.TrnspName ShippingName, T0.DocStatus SoStatus " & _
                            ' "from ordr T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode " '& _
                            ''" Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("TxtDtFrm").Specific.string), "yyyyMMdd") & "' " & _
                            ''" AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("TxtDtTo").Specific.string), "yyyyMMdd") & "' "

                            If oForm.Items.Item("TxtDtFrm").Specific.string = "" Or oForm.Items.Item("TxtDtTo").Specific.string = "" Then
                                PDOStatusQuery = "SELECT 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                                                "T2.PlannedQty PdoQty, T2.CmpltQty PdoReceiptQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                                                "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                                                "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                                                "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                                                "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                                                "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                                                "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O'  " & _
                                                " AND T2.Status = 'P' " & _
                                                "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "
                            Else
                                PDOStatusQuery = "SELECT '1' gbr, 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                                               "T2.PlannedQty PdoQty, T2.CmpltQty PdoReceiptQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                                               "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                                               "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                                               "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                                               "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                                               "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                                               "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O' " & _
                                               " AND T2.Status = 'P' " & _
                                               " AND T2.PostDate >= '" & MIS_Toolkit.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                                               " AND T2.PostDate <= '" & MIS_Toolkit.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' " & _
                                               "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "

                                'PDOStatusQuery = "SELECT '1' gbr, 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                                '               "T2.PlannedQty PdoQty, T2.CmpltQty PdoReceiptQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                                '               "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                                '               "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                                '               "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                                '               "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                                '               "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                                '               "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O' " & _
                                '               " AND T2.Status = 'P' " & _
                                '               " AND T2.PostDate >= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                                '               " AND T2.PostDate <= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' " & _
                                '               "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "

                            End If

                            ' Grid #: 1
                            'oForm.DataSources.DataTables.Add("DelOutLst")
                            oForm.DataSources.DataTables.Item("PDOStatusLst").ExecuteQuery(PDOStatusQuery)
                            oPDOStatusGrid.DataTable = oForm.DataSources.DataTables.Item("PDOStatusLst")

                            oPDOStatusGrid.Columns.Item("gbr").Type = SAPbouiCOM.BoGridColumnType.gct_Picture


                            oPDOStatusGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                            oPDOStatusGrid.Columns.Item("Release PdO").TitleObject.Sortable = True
                            'oPDOStatusGrid.Columns.Item("Release PdO").BackColor = 7

                            oColumn = oPDOStatusGrid.Columns.Item("PdoDate")
                            oPDOStatusGrid.Columns.Item("PdoDate").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("PdoEntry")
                            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_ProductionOrder
                            oPDOStatusGrid.Columns.Item("PdoEntry").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("Pdo#")
                            oPDOStatusGrid.Columns.Item("Pdo#").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("SOEntry")
                            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
                            oPDOStatusGrid.Columns.Item("SOEntry").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("SoNumber")
                            oPDOStatusGrid.Columns.Item("SoNumber").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("SoDate")
                            oPDOStatusGrid.Columns.Item("SoDate").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("FGItemCode")
                            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
                            oPDOStatusGrid.Columns.Item("FGItemCode").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("PdoQty")
                            oPDOStatusGrid.Columns.Item("PdoQty").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("PdoReceiptQty")
                            oPDOStatusGrid.Columns.Item("PdoReceiptQty").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("UM")
                            oPDOStatusGrid.Columns.Item("UM").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("ExpDelDate")
                            oPDOStatusGrid.Columns.Item("ExpDelDate").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("Delayed")
                            'oPDOStatusGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                            oPDOStatusGrid.Columns.Item("Delayed").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("Progress")
                            oPDOStatusGrid.Columns.Item("Progress").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("CustomerCode")
                            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
                            oPDOStatusGrid.Columns.Item("CustomerCode").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("Customer")
                            oPDOStatusGrid.Columns.Item("Customer").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("SoLine")
                            oPDOStatusGrid.Columns.Item("SoLine").TitleObject.Sortable = True
                            oColumn.Editable = False

                            oColumn = oPDOStatusGrid.Columns.Item("SalesRep")
                            oPDOStatusGrid.Columns.Item("SalesRep").TitleObject.Sortable = True
                            oColumn.Editable = False


                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPDOStatusGrid)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

                        End If
                        ' Karno Out Del
                    Case "OutDel"
                        If ((pVal.ItemUID = "BtnShow") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)) Then

                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            Dim DelOutQuery As String
                            Dim oDelOutGrid As SAPbouiCOM.Grid = Nothing
                            Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

                            'Dim oMis_Utils As MIS_Utils

                            'oMis_Utils = New MIS_Utils

                            oDelOutGrid = oForm.Items.Item("myGrid1").Specific


                            If oForm.Items.Item("TxtDtFrm").Specific.string = "" Or oForm.Items.Item("TxtDtTo").Specific.string = "" Then

                                DelOutQuery = "select DISTINCT T0.DocEntry DocEntry, T0.DocNum So_DocNum, T0.DocDate So_Date, " & _
                                "T4.SlpName Sales_Rep, T0.CardCode Customer_Code, T0.CardName Customer_Name, " & _
                                "T1.TrnspName Shipping_Type, CASE WHEN(select COUNT(P1.itemcode) from ORDR P0 " & _
                                "INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry WHERE P0.DocEntry = T0.DocEntry " & _
                                ") > (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry " & _
                                "LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode WHERE P0.DocEntry = T0.Docentry " & _
                                "AND P2.OnHand > 0) THEN 'Partialy Ready' " & _
                                "WHEN (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.Docentry " & _
                                "WHERE P0.DocEntry = T0.Docentry) = (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1  " & _
                                "ON P0.DocEntry = P1.DocEntry LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode " & _
                                "WHERE P0.DocEntry = T0.Docentry AND P2.OnHand > 0) THEN 'Completely Ready' END Status " & _
                                "From ORDR T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode  LEFT JOIN OSLP T4 ON T0.SlpCode = T4.SlpCode INNER JOIN RDR1 T2 " & _
                                "ON T0.DocEntry = T2.DocEntry LEFT JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " & _
                                "AND T2.WhsCode = T3.WhsCode WHERE T3.OnHand >= T2.Quantity AND T0.DocStatus = 'O' ORDER BY T0.DocDate "
                            Else
                                DelOutQuery = "select DISTINCT T0.DocEntry DocEntry, T0.DocNum So_DocNum, T0.DocDate So_Date, " & _
                                "T4.SlpName Sales_Rep, T0.CardCode Customer_Code, T0.CardName Customer_Name, " & _
                                "T1.TrnspName Shipping_Type, CASE WHEN(select COUNT(P1.itemcode) from ORDR P0 " & _
                                "INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry WHERE P0.DocEntry = T0.DocEntry " & _
                                ") > (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry " & _
                                "LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode WHERE P0.DocEntry = T0.Docentry " & _
                                "AND P2.OnHand > 0) THEN 'Partialy Ready' " & _
                                "WHEN (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.Docentry " & _
                                "WHERE P0.DocEntry = T0.Docentry) = (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1  " & _
                                "ON P0.DocEntry = P1.DocEntry LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode " & _
                                "WHERE P0.DocEntry = T0.Docentry AND P2.OnHand > 0) THEN 'Completely Ready' END Status " & _
                                "From ORDR T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode  LEFT JOIN OSLP T4 ON T0.SlpCode = T4.SlpCode INNER JOIN RDR1 T2 " & _
                                "ON T0.DocEntry = T2.DocEntry LEFT JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " & _
                                "AND T2.WhsCode = T3.WhsCode WHERE T3.OnHand >= T2.Quantity AND T0.DocStatus = 'O' " & _
                                " AND T0.DocDate >= '" & MIS_Toolkit.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                                " AND T0.DocDate <= '" & MIS_Toolkit.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' ORDER BY T0.DocDate "

                                'DelOutQuery = "select DISTINCT T0.DocEntry DocEntry, T0.DocNum So_DocNum, T0.DocDate So_Date, " & _
                                '"T4.SlpName Sales_Rep, T0.CardCode Customer_Code, T0.CardName Customer_Name, " & _
                                '"T1.TrnspName Shipping_Type, CASE WHEN(select COUNT(P1.itemcode) from ORDR P0 " & _
                                '"INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry WHERE P0.DocEntry = T0.DocEntry " & _
                                '") > (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry " & _
                                '"LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode WHERE P0.DocEntry = T0.Docentry " & _
                                '"AND P2.OnHand > 0) THEN 'Partialy Ready' " & _
                                '"WHEN (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.Docentry " & _
                                '"WHERE P0.DocEntry = T0.Docentry) = (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1  " & _
                                '"ON P0.DocEntry = P1.DocEntry LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode " & _
                                '"WHERE P0.DocEntry = T0.Docentry AND P2.OnHand > 0) THEN 'Completely Ready' END Status " & _
                                '"From ORDR T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode  LEFT JOIN OSLP T4 ON T0.SlpCode = T4.SlpCode INNER JOIN RDR1 T2 " & _
                                '"ON T0.DocEntry = T2.DocEntry LEFT JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " & _
                                '"AND T2.WhsCode = T3.WhsCode WHERE T3.OnHand >= T2.Quantity AND T0.DocStatus = 'O' " & _
                                '" AND T0.DocDate >= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                                '" AND T0.DocDate <= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' ORDER BY T0.DocDate "

                            End If


                            ' Grid #: 1
                            'oForm.DataSources.DataTables.Add("DelOutLst")
                            oForm.DataSources.DataTables.Item("DelOutLst").ExecuteQuery(DelOutQuery)
                            oDelOutGrid.DataTable = oForm.DataSources.DataTables.Item("DelOutLst")

                            oColumn = oDelOutGrid.Columns.Item("DocEntry")
                            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("So_DocNum")
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("So_Date")
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("Sales_Rep")
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("Customer_Code")
                            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("Customer_Name")
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("Shipping_Type")
                            oColumn.Editable = False

                            oColumn = oDelOutGrid.Columns.Item("Status")
                            oColumn.Editable = False

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDelOutGrid)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

                        End If


                    Case "mds_p1"
                        ''If pVal.Before_Action = False Then
                        'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And oForm IsNot (Nothing) Then
                        '    'oItemMat = oForm.Items.Item("matrixName")
                        '    'oItemMat.Width = oForm.Width - 200
                        '    RearrangeGrid(oForm)

                        'End If
                        ''End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID

                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)



                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects

                                Dim xval As String


                                xval = oDataTable.GetValue(0, 0)

                                If pVal.ItemUID = "BPCardCode" Or pVal.ItemUID = "BPButton" Then

                                    oForm.DataSources.UserDataSources.Item("BPDS").ValueEx = xval
                                End If

                                oCFL = Nothing
                                oDataTable = Nothing
                            End If

                            'oForm = Nothing
                            'oCFLEvento = Nothing
                            'GC.Collect()

                        End If


                        ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                        ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadSO") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'If Not SOToMFGFormValid(oForm) Then
                            '    SBO_Application.SetStatusBarMessage("Form invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '    BubbleEvent = False

                            'Else
                            '    LoadSO(oForm)
                            'End If

                            LoadSO(oForm)


                        End If

                        If (pVal.ItemUID = "SODateFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'Dim vdate As MISToolbox
                            'vdate = New MISToolbox
                            'Dim validDate As Boolean


                            If Len(oForm.Items.Item("SODateFrom").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'validDate = vdate.SBODateisValid("2010918")

                            'validDate = vdate.SBODateisValid(oForm.Items.Item("SODateFrom").Specific.string)
                            'If validDate = False Then
                            '    SBO_Application.SetStatusBarMessage("SO Date From is invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If

                            If Len(oForm.Items.Item("SODateFrom").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            If Len(oForm.Items.Item("SODateFrom").Specific.string) = 8 Then
                                oForm.Items.Item("SODateFrom").Specific.string = _
                                    CDate(Left(oForm.Items.Item("SODateFrom").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("SODateFrom").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("SODateFrom").Specific.string, 2))
                            End If

                            If oForm.Items.Item("SODateFrom").Specific.string = "" Then
                                oForm.Items.Item("SODateFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                            End If

                            If oForm.Items.Item("SODateTo").Specific.string = "" Then
                                oForm.Items.Item("SODateTo").Specific.string = oForm.Items.Item("SODateFrom").Specific.string
                            End If

                            'vdate = Nothing

                            'oForm.Items.Item("SODateFrom").Click()
                            '                        BubbleEvent = False
                        End If

                        If pVal.ItemUID = "SODateTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("SODateTo").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                'Exit Sub
                            End If

                            If Len(oForm.Items.Item("SODateTo").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                'BubbleEvent = False
                            End If

                            If Len(oForm.Items.Item("SODateTo").Specific.string) = 8 Then
                                oForm.Items.Item("SODateTo").Specific.string = _
                                    CDate(Left(oForm.Items.Item("SODateTo").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("SODateTo").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("SODateTo").Specific.string, 2))
                            End If
                            'BubbleEvent = True
                            'oForm.Items.Item("SODateTo").Click('')

                            'oForm = Nothing
                            'GC.Collect()

                        End If

                        If pVal.ItemUID = "cmdGenPdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oSOToMFGGrid As SAPbouiCOM.Grid

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("SOToMFGLst")

                            oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

                            'get total row count selected
                            'oSOToMFGGrid.Rows.SelectedRows.Count.ToString()


                            'selection rows -> e.g: user select row# by order respectively: 1, 3, 2, 5

                            'get row index of selected grid, has two method:
                            'method# 1: ot_RowOrder (value=1)
                            'result row selected: 1, 2, 3, 5

                            'method# 2: ot_SelectionOrder (value=0)
                            'result row selected: 1, 3, 2, 5

                            'For idx = 0 To oSOToMFGGrid.Rows.SelectedRows.Count - 1
                            '    MsgBox("selected row#:" & idx.ToString & _
                            '           "; selectedrow->row#: " & oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
                            '           & "docnum: " & oSOToMFGGrid.DataTable.GetValue(0, oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

                            'Next

                            'Dim oPdO As SAPbobsCOM.ProductionOrders
                            'oPdO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                            ''Fill PdO properties...
                            'oPdO.ItemNo = "LM4029"
                            ''oPdO.DueDate = oSOToMFGGrid.DataTable.GetValue(13, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                            'oPdO.DueDate = DateTime.Today.ToString("yyyyMMdd")
                            'oPdO.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                            'oPdO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                            'oPdO.PlannedQuantity = 188
                            'oPdO.PostingDate = DateTime.Today 'DateTime.Today.ToString("yyyyMMdd")
                            'oPdO.Add()

                            '???
                            GeneratePdOFromSO(oForm)

                            LoadSO(oForm)


                        End If

                        'toggle select/unselect all
                        If pVal.ColUID = "Release PdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oSOToMFGGrid As SAPbouiCOM.Grid

                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("SOToMFGLst")

                            oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

                            'get total row count selected
                            'oSOToMFGGrid.Rows.SelectedRows.Count.ToString()


                            oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

                            If oSOToMFGGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oSOToMFGGrid.Rows.Count - 1
                                    dt.SetValue("Release PdO", idx, "Y")
                                Next
                                oSOToMFGGrid.Columns.Item(1).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oSOToMFGGrid.Rows.Count - 1
                                    dt.SetValue("Release PdO", idx, "N")
                                Next
                                oSOToMFGGrid.Columns.Item(1).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                            'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        End If


                    Case "mds_p3"
                        'If pVal.Before_Action = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                            oForm = oApp.Forms.Item(FormUID)

                            'oItemMat = oForm.Items.Item("matrixName")
                            'oItemMat.Width = oForm.Width - 200
                            RearrangeFormOptimEntry(oForm)
                            DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))
                        End If
                        'End If

                        If pVal.ItemUID = "SeriesOptm" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then ' 3 = fm_ADD_MODE 
                            Dim lNextSeriesNumOptimization As Long
                            Dim Series As String
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            Dim cmbSeries As SAPbouiCOM.ComboBox
                            cmbSeries = oForm.Items.Item("SeriesOptm").Specific
                            Series = cmbSeries.Selected.Value
                            lNextSeriesNumOptimization = oForm.BusinessObject.GetNextSerialNumber(Series)

                            Dim oItem As SAPbouiCOM.EditText
                            oItem = oForm.Items.Item("DocNum").Specific
                            oItem.Value = lNextSeriesNumOptimization

                            oItem = oForm.Items.Item("ByUser").Specific
                            oItem.Value = oCompany.UserName.ToString

                            oForm.Items.Item("KcSisaPctg").Specific.value = 0
                            oForm.Items.Item("TotWastPct").Specific.value = 0

                            Dim oMatrix As SAPbouiCOM.Matrix
                            oMatrix = oForm.Items.Item("OptimMtx").Specific
                            oMatrix.AddRow()
                            oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount
                            oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount

                            ''????
                            'Dim txtColor As SAPbouiCOM.EditText
                            'txtColor = oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific
                            'txtColor.BackColor = 3

                            oForm.Items.Item("GTabc").Specific.value = 0
                            oForm.Items.Item("GTaloc").Specific.value = 0
                            oForm.Items.Item("GTplanPdO").Specific.value = 0

                            oForm.Items.Item("QtyLembar").Specific.value = 1
                            oForm.Items.Item("OptimDate").Specific.value = DateTime.Today.ToString("yyyyMMdd")

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(cmbSeries)


                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID

                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable = Nothing
                                oDataTable = oCFLEvento.SelectedObjects

                                If Not oDataTable Is Nothing Then
                                    Dim xVal As String
                                    xVal = oDataTable.GetValue(0, 0)

                                    Dim oDBDataSource As SAPbouiCOM.DBDataSource
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                                    If pVal.ItemUID = "ItemCode" Or pVal.ItemUID = "ItemButton" Then
                                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_ItemCode", oDBDataSource.Offset, xVal)
                                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_ItemDesc", oDBDataSource.Offset, oDataTable.GetValue(1, 0))
                                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_Pcm", oDBDataSource.Offset, oDataTable.GetValue("SHeight1", 0))
                                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_Lcm", oDBDataSource.Offset, oDataTable.GetValue("SWidth1", 0))

                                        Dim oRecLengthWidth As SAPbobsCOM.Recordset = Nothing
                                        Dim StrQuery As String

                                        oRecLengthWidth = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        StrQuery = "SELECT UnitCode, SizeInMM FROM OITM T0 JOIN OLGT T1 ON T0.SHght1Unit = T1.UnitCode WHERE ItemCode = '" & xVal & "' "

                                        oRecLengthWidth.DoQuery(StrQuery)

                                        Dim inMM As Double
                                        Dim LuasInM2 As Double

                                        inMM = oRecLengthWidth.Fields.Item("SizeInMM").Value
                                        LuasInM2 = (oDataTable.GetValue("SHeight1", 0) * inMM) * (oDataTable.GetValue("SWidth1", 0) * inMM) / 1000000
                                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_LuasM2", oDBDataSource.Offset, LuasInM2)

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecLengthWidth)
                                        oRecLengthWidth = Nothing

                                    End If
                                    If pVal.ItemUID = "ItemKcSisa" Or pVal.ItemUID = "ItmSisaBtn" Then
                                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_ItemCdKacaSisa", oDBDataSource.Offset, xVal)
                                    End If

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable)
                                    GC.Collect()

                                End If

                            End If

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLEvento)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)

                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then

                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then
                                oMatrix = oForm.Items.Item("OptimMtx").Specific

                                oForm.Freeze(True)


                                Dim idx As Long
                                Dim gtabc As Double
                                Dim gtaloc As Double
                                Dim gtplanpdo As Double

                                gtabc = 0
                                gtaloc = 0
                                gtplanpdo = 0
                                For idx = 1 To oMatrix.RowCount
                                    gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                                    'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                                    gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))
                                    'oForm.Items.Item("#").Specific.value = idx
                                    'oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.value = oMatrix.VisualRowCount
                                    oMatrix.Columns.Item("#").Cells.Item(CInt(idx)).Specific.value = idx


                                Next

                                Dim oDBDataSource As SAPbouiCOM.DBDataSource = Nothing
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                                Dim i As Integer
                                gtabc = 0
                                Dim oDBDataSource_OptimL As SAPbouiCOM.DBDataSource = Nothing
                                oDBDataSource_OptimL = oForm.DataSources.DBDataSources.Item("@MIS_OPTIML")
                                For i = 0 To oDBDataSource_OptimL.Size - 1
                                    gtabc += oDBDataSource_OptimL.GetValue("U_MIS_TotalABC", i)
                                Next

                                Dim docnum As Integer
                                Dim LuasKaca As Double
                                Dim SisaKacaUtuh As Double
                                Dim TotalWaste As Double

                                docnum = oDBDataSource.GetValue("docnum", 0)
                                LuasKaca = oDBDataSource.GetValue("U_MIS_LuasM2", 0)
                                SisaKacaUtuh = oDBDataSource.GetValue("U_MIS_KcSisaUtuh", 0)
                                TotalWaste = oDBDataSource.GetValue("U_MIS_TotalWaste", 0)

                                ' by Toin 2011-02-10
                                oForm.Items.Item("TotWastPct").Specific.value = IIf(LuasKaca = 0, 0, Math.Round(TotalWaste / LuasKaca * 100, 2))
                                ' by Toin 2011-02-10
                                ' by Toin 2011-03-01
                                oForm.Items.Item("KcSisaPctg").Specific.value = IIf(LuasKaca = 0, 0, Math.Round(SisaKacaUtuh / LuasKaca * 100, 2))
                                ' by Toin 2011-03-01

                                gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                                oForm.Items.Item("GTabc").Specific.value = gtabc
                                'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                                oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo



                                ' by Toin 2011-02-28 = Grand Total Waste = Total Waste + Sisa Kaca Utuh
                                'LuasKaca(-SisaKacaUtuh - gtabc)
                                oForm.Items.Item("GTaloc").Specific.value = _
                                TotalWaste + SisaKacaUtuh
                                'TotalWaste

                                BubbleEvent = False

                                oForm.Freeze(False)

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDBDataSource)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDBDataSource_OptimL)

                                oMatrix = Nothing
                                oDBDataSource = Nothing
                                oDBDataSource_OptimL = Nothing
                                oForm = Nothing

                                GC.Collect()


                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD _
                                 Then

                            End If

                        End If


                        If (pVal.ItemUID = "LebarKaca" Or pVal.ItemUID = "PnjangKaca") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            Dim oItem As SAPbouiCOM.Item
                            Dim oMatrix As SAPbouiCOM.Matrix
                            Dim oEditText As SAPbouiCOM.EditText

                            'Dim sb As String

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            oItem = oForm.Items.Item("OptimMtx")
                            oMatrix = oItem.Specific
                            'oEditText = oMatrix.GetCellSpecific(5, 1)
                            'sb = oEditText.Value

                            Dim oRecLengthWidth As SAPbobsCOM.Recordset = Nothing
                            Dim StrQuery As String

                            oRecLengthWidth = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            StrQuery = "SELECT UnitCode, SizeInMM FROM OITM T0 JOIN OLGT T1 ON T0.SHght1Unit = T1.UnitCode WHERE ItemCode = '" & oForm.Items.Item("ItemCode").Specific.value & "' "

                            oRecLengthWidth.DoQuery(StrQuery)

                            Dim inMM As Double
                            'Dim LuasInM2 As Double

                            inMM = oRecLengthWidth.Fields.Item("SizeInMM").Value
                            'LuasInM2 = oDataTable.GetValue("SHeight1", 0) * inMM * oDataTable.GetValue("SWidth1", 0) / 1000
                            'oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_LuasM2", oDBDataSource.Offset, LuasInM2)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecLengthWidth)
                            oRecLengthWidth = Nothing
                            GC.Collect()

                            oEditText = oForm.Items.Item("LuasKaca").Specific
                            oEditText.Value = _
                                IIf(oForm.Items.Item("PnjangKaca").Specific.value = "", 0, oForm.Items.Item("PnjangKaca").Specific.value) _
                                * inMM * _
                                IIf(oForm.Items.Item("LebarKaca").Specific.value = "", 0, oForm.Items.Item("LebarKaca").Specific.value) _
                                * inMM _
                                / 1000000

                            'oEditText = oForm.Items.Item("KacaPakai").Specific
                            'oEditText.Value = oForm.Items.Item("LuasKaca").Specific.value - oForm.Items.Item("SisaKcUtuh").Specific.value

                            'oEditText = oForm.Items.Item("GTaloc").Specific
                            'oEditText.Value = _
                            '    IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) _
                            '    - IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) _
                            '    - IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)

                            '2011-02-28
                            oEditText = oForm.Items.Item("GTaloc").Specific
                            oEditText.Value = _
                                CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                                + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))
                            'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)

                            'oEditText = oForm.Items.Item("TotalABC").Specific

                            oForm.Items.Item("KacaPakai").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                            - CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) _
                            - CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            GC.Collect()

                        End If

                        If pVal.ItemUID = "KcSisaPctg" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            Dim oEditText As SAPbouiCOM.EditText

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            oEditText = oForm.Items.Item("SisaKcUtuh").Specific

                            If oForm.Items.Item("SisaKcUtuh").Specific.value = "" Then
                                oEditText.Value = (IIf(oForm.Items.Item("KcSisaPctg").Specific.value = "", 0, oForm.Items.Item("KcSisaPctg").Specific.value) * IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) / 100
                            End If

                            oForm.Items.Item("KacaPakai").Specific.value = _
                            IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) - _
                            IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) - _
                            IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
                            GC.Collect()

                        End If

                        If pVal.ItemUID = "SisaKcUtuh" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            oForm.Items.Item("KacaPakai").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                            - CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) _
                            - CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            'GC.Collect()

                            Dim TotalWaste As Double
                            Dim Kolom As Double

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                            oForm.Items.Item("TotalWaste").Specific.Value = (CDbl(IIf(oForm.Items.Item("TotWastPct").Specific.value = "", 0, oForm.Items.Item("TotWastPct").Specific.value)) * CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value))) / 100

                            ' 2011-02-28
                            TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                                + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                            Dim idx As Long
                            Dim gtabc As Double
                            Dim gtplanpdo As Double

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oMatrix = oForm.Items.Item("OptimMtx").Specific

                            oForm.Freeze(True)

                            gtabc = 0

                            'gtplanpdo = 0
                            'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                            For idx = 1 To oMatrix.RowCount
                                gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                                'gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                            Next
                            'End If

                            'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And oMatrix.RowCount = 0 Then
                            '    oForm.Items.Item("GTabc").Specific.value = 0
                            'Else
                            oForm.Items.Item("GTabc").Specific.value = gtabc
                            'End If

                            'oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                            ' 2011-02-28
                            oForm.Items.Item("GTaloc").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                            'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) _
                            '+ IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, CDbl(oForm.Items.Item("SisaKcUtuh").Specific.value))

                            'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value))

                            Kolom = oForm.Items.Item("GTabc").Specific.value

                            'oMatrix = oForm.Items.Item("OptimMtx").Specific

                            'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then

                            For Row = 1 To oMatrix.RowCount
                                'Allocated Waste
                                oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = _
                                Math.Round( _
                                    (CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, _
                                    oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                    / Kolom) * TotalWaste, _
                                4)


                                'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                                oMatrix.Columns.Item("PlanPdIsue").Cells.Item(Row).Specific.value = _
                                Math.Round( _
                                    CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) + _
                                    CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value)) _
                                , 4)


                            Next
                            'End If

                            gtplanpdo = 0
                            'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                            For idx = 1 To oMatrix.RowCount
                                'gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                                gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                            Next
                            'End If

                            oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                            ' 2011-02-28
                            oForm.Items.Item("GTaloc").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                            ' by Toin 2011-03-01
                            If CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) = 0 Or _
                                CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) = 0 Then
                                oForm.Items.Item("KcSisaPctg").Specific.value = 0
                            Else
                                oForm.Items.Item("KcSisaPctg").Specific.value = _
                                    (IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, _
                                        Math.Round(CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, _
                                                       oForm.Items.Item("SisaKcUtuh").Specific.value)) / _
                                                   CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, _
                                                       oForm.Items.Item("LuasKaca").Specific.value)) * 100, 2)))
                            End If
                            ' by Toin 2011-03-01

                            oForm.Items.Item("KacaPakai").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) - _
                            CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) - _
                            CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))


                            oForm.Freeze(False)
                            'oForm.Refresh()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            GC.Collect()

                        End If


                        If (pVal.ItemUID = "TotalWaste" Or pVal.ItemUID = "TotWastPct") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            Dim TotalWaste As Double
                            Dim Kolom As Double
                            'Dim oForm As SAPbouiCOM.Form = Nothing

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            oForm.Items.Item("TotalWaste").Specific.value = _
                                Math.Round( _
                                    CDbl(IIf(oForm.Items.Item("TotWastPct").Specific.value = "", 0.0, oForm.Items.Item("TotWastPct").Specific.value)) / 100 * _
                                    CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                                , 4)


                            Dim totwastpct As Double
                            Dim luaskaca As Double
                            totwastpct = CDbl(IIf(oForm.Items.Item("TotWastPct").Specific.value = "", 0.0, oForm.Items.Item("TotWastPct").Specific.value))
                            luaskaca = CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value))

                            ' 2011-02-28
                            TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                                + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                            'oForm.Items.Item("GTaloc").Specific.value = TotalWaste

                            Dim idx As Long
                            Dim gtabc As Double
                            Dim gtplanpdo As Double

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oMatrix = oForm.Items.Item("OptimMtx").Specific

                            oForm.Freeze(True)

                            gtabc = 0

                            'gtplanpdo = 0
                            'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                            For idx = 1 To oMatrix.RowCount
                                gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                                'gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                            Next
                            'End If
                            'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And oMatrix.RowCount = 0 Then
                            '    oForm.Items.Item("GTabc").Specific.value = 0
                            'Else
                            oForm.Items.Item("GTabc").Specific.value = gtabc
                            'End If
                            'oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                            ' 2011-02-28
                            oForm.Items.Item("GTaloc").Specific.value = _
                                CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                                + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))
                            'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value))

                            Kolom = oForm.Items.Item("GTabc").Specific.value

                            'oMatrix = oForm.Items.Item("OptimMtx").Specific

                            'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then

                            For Row = 1 To oMatrix.RowCount
                                'Allocated Waste
                                oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = _
                                Math.Round( _
                                    (CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, _
                                    oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                    / Kolom) * TotalWaste, _
                                4)


                                'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                                oMatrix.Columns.Item("PlanPdIsue").Cells.Item(Row).Specific.value = _
                                Math.Round( _
                                    CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) + _
                                    CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value)) _
                                , 4)


                            Next
                            'End If

                            gtplanpdo = 0
                            'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                            For idx = 1 To oMatrix.RowCount
                                'gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                                gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                            Next
                            'End If

                            oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                            ' 2011-02-28
                            oForm.Items.Item("GTaloc").Specific.value = _
                                CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                                + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))


                            oForm.Items.Item("KacaPakai").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                            - CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) _
                            - CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                            oForm.Freeze(False)
                            'oForm.Refresh()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            GC.Collect()

                        End If

                        '???
                        If pVal.ItemUID = "OptimMtx" And _
                            pVal.ColUID = "PdO#" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then

                            'Dim oForm As SAPbouiCOM.Form = Nothing

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            oMatrix = oForm.Items.Item("OptimMtx").Specific

                            oForm.Freeze(True)

                            'Total AxBxC = Jumlah Potong x P x L
                            'oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = _

                            If oMatrix.Columns.Item("PdO#").Cells.Item(pVal.Row).Specific.value <> "" Then
                                Dim oRecPdo As SAPbobsCOM.Recordset = Nothing
                                Dim StrQuery As String

                                oRecPdo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                StrQuery = "SELECT T0.DocNum, OriginNum, T2.LineNum, T0.CardCode, T1.CardName FROM OWOR T0 " & _
                                    " LEFT JOIN ORDR T1 ON T1.DocNum = OriginNum " & _
                                    " LEFT JOIN RDR1 T2 ON T2.DocEntry = T1.DocEntry " & _
                                    " WHERE T0.DocNum = " & oMatrix.Columns.Item("PdO#").Cells.Item(pVal.Row).Specific.value
                                '" & oForm.Items.Item("ItemCode").Specific.value & "' "

                                oRecPdo.DoQuery(StrQuery)


                                oMatrix.Columns.Item("SO#").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("OriginNum").Value
                                oMatrix.Columns.Item("SOLine").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("LineNum").Value
                                oMatrix.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("CardCode").Value
                                oMatrix.Columns.Item("CardName").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("CardName").Value


                                'LuasInM2 = oDataTable.GetValue("SHeight1", 0) * inMM * oDataTable.GetValue("SWidth1", 0) / 1000
                                'oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_LuasM2", oDBDataSource.Offset, LuasInM2)

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecPdo)
                                oRecPdo = Nothing

                                'GC.WaitForPendingFinalizers()
                                GC.Collect()


                            End If

                            'oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)

                            'BubbleEvent = False
                            oForm.Freeze(False)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            GC.Collect()

                        End If


                        If pVal.ItemUID = "OptimMtx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            oForm.EnableMenu("1292", True) 'Add Row
                            oForm.EnableMenu("1293", True) 'Delete Row

                            oForm.EnableMenu("1287", True) 'Duplicate

                        End If

                        'If pVal.ItemUID = "OptimMtx" And _
                        '    ((pVal.ColUID = "P" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                        '     (pVal.ColUID = "L" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)) Then

                        If pVal.ItemUID = "OptimMtx" And _
                            ((pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                            (pVal.ColUID = "P" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                             (pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                             (pVal.ColUID = "L" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)) Then

                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            Dim oEditText As SAPbouiCOM.EditText = Nothing

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                            Dim idx As Long
                            Dim gtabc As Double
                            Dim gtaloc As Double
                            Dim gtplanpdo As Double

                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            oMatrix = oForm.Items.Item("OptimMtx").Specific

                            oForm.Freeze(True)

                            'Total AxBxC = Jumlah Potong x P x L
                            oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = _
                            Math.Round( _
                                CDbl(IIf(oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Specific.value)) _
                                * CDbl(IIf(oMatrix.Columns.Item("P").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("P").Cells.Item(pVal.Row).Specific.value)) _
                                * CDbl(IIf(oMatrix.Columns.Item("L").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("L").Cells.Item(pVal.Row).Specific.value)) _
                                / 1000000 _
                            , 4)

                            gtabc = 0
                            gtaloc = 0
                            gtplanpdo = 0

                            If CDbl(IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, oForm.Items.Item("GTaloc").Specific.value)) <> 0 Then
                                For idx = 1 To oMatrix.RowCount
                                    gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                                    'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                                    gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                                Next
                            End If

                            'gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                            oForm.Items.Item("GTabc").Specific.value = gtabc
                            'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                            'oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                            'oEditText.Value = _
                            '    IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) _
                            '    - IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) _
                            '    - IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)

                            ' 2011-02-28
                            oEditText = oForm.Items.Item("GTaloc").Specific
                            oEditText.Value = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                            'Allocated Waste
                            oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = _
                            Math.Round( _
                                CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value)) _
                                / CDbl(IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)) _
                                * CDbl(IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, oForm.Items.Item("GTaloc").Specific.value)) _
                            , 4)


                            'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                            oMatrix.Columns.Item("PlanPdIsue").Cells.Item(pVal.Row).Specific.value = _
                            Math.Round( _
                                CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value)) + _
                                CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value)) _
                            , 4)

                            'Dim totalABC As Double
                            'Dim alocatedWaste As Double

                            'totalABC = IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value))
                            'alocatedWaste = IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value))

                            Dim TotalWaste As Double
                            Dim Kolom As Double

                            'oForm = SBO_Application.Forms.Item(FormUID)

                            TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                            ' 2011-02-28
                            TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                                + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                            Kolom = oForm.Items.Item("GTabc").Specific.value

                            'oMatrix = oForm.Items.Item("OptimMtx").Specific

                            If CDbl(IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, oForm.Items.Item("GTaloc").Specific.value)) <> 0 Then

                                For Row = 1 To oMatrix.RowCount
                                    'Allocated Waste
                                    oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = _
                                    Math.Round( _
                                        (CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, _
                                        oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                        / Kolom) * TotalWaste _
                                    , 4)


                                    'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                                    oMatrix.Columns.Item("PlanPdIsue").Cells.Item(Row).Specific.value = _
                                    Math.Round( _
                                        CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                        + CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value)) _
                                    , 4)

                                    'totalABC = IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value))
                                    'alocatedWaste = IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value))

                                Next

                            End If

                            gtplanpdo = 0
                            'If IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, CDbl(oForm.Items.Item("GTaloc").Specific.value)) <> 0 Then
                            For idx = 1 To oMatrix.RowCount
                                'gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                                'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                                gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                            Next
                            'End If

                            'gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                            'oForm.Items.Item("GTabc").Specific.value = gtabc
                            'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                            oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo


                            'BubbleEvent = False

                            'oForm.Refresh()
                            oForm.Freeze(False)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
                            GC.Collect()

                        End If

                        '-------------- Yadi FC ----------------------------
                    Case "MDS_P6"
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnShow") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            LoadProductionClosed(oForm)
                        End If


                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnCancel") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            oForm.Close()
                        End If

                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnUpdate") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            UpdateProductionClosed(oForm)
                            LoadProductionClosed(oForm)
                        End If

                        '-------------- Yadi FC ----------------------------
                        'toggle select/unselect all
                        'If pVal.ColUID = "Release PdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                        If pVal.ColUID = "Check" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oProdClosedGrid As SAPbouiCOM.Grid

                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable

                            'Dim str As String

                            dt = oForm.DataSources.DataTables.Item("PCLst")

                            oProdClosedGrid = oForm.Items.Item("grdPC").Specific

                            'get total row count selected
                            'oProdClosedGrid.Rows.SelectedRows.Count.ToString()


                            oProdClosedGrid = oForm.Items.Item("grdPC").Specific

                            If oProdClosedGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oProdClosedGrid.Rows.Count - 1
                                    If oProdClosedGrid.DataTable.GetValue(6, oProdClosedGrid.GetDataTableRowIndex(idx)) = _
                                        oProdClosedGrid.DataTable.GetValue(7, oProdClosedGrid.GetDataTableRowIndex(idx)) Then
                                        dt.SetValue("Check", idx, "Y")
                                    End If
                                    'str = oProdClosedGrid.Columns.Item(idx).Description
                                    'str = oProdClosedGrid.DataTable.GetValue(1, oProdClosedGrid.GetDataTableRowIndex(idx))
                                Next
                                oProdClosedGrid.Columns.Item(1).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oProdClosedGrid.Rows.Count - 1
                                    dt.SetValue("Check", idx, "N")
                                Next
                                oProdClosedGrid.Columns.Item(1).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                            'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        End If

                End Select
            End If

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    'Sub OnBPFormLoadAfter(ByRef oForm As SAPbouiCOM.Form)
    '    Try
    '        '1. Disable the CardName
    '        Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("7")
    '        oItem.Enabled = False

    '        '2. Hide the VAT number
    '        oItem = oForm.Items.Item("41")
    '        oItem.Visible = False

    '        '3 Add the "second" button next cancel button
    '        oItem = oForm.Items.Item("2")
    '        Dim layout As Layout = New Layout(oItem, New Layout(0, oItem.Width + 10, 0, 0))
    '        oItem = CreateItem(oForm, "bSecond", SAPbouiCOM.BoFormItemTypes.it_BUTTON, layout)
    '        Dim oButton As SAPbouiCOM.Button = oItem.Specific
    '        oButton.Caption = "bSecond"

    '        '4. Update form Title
    '        oForm.Title = "My BP"

    '    Catch ex As Exception
    '        MsgBoxWrapper(ex.Message)
    '    End Try
    'End Sub

    'Sub OnBPFormValidate(ByRef oForm As SAPbouiCOM.Form, _
    '                     ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    '    If pVal.ItemUID = "128" Then
    '        If pVal.BeforeAction Then
    '            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

    '                'Get foreign name
    '                Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("128")
    '                Dim oEdit As SAPbouiCOM.EditText = oItem.Specific

    '                If String.IsNullOrEmpty(oEdit.String.Trim()) Then
    '                    MsgBoxWrapper("Enter Foreign Name!", MsgBoxType.B1StatusBarMsg, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    BubbleEvent = False
    '                End If
    '            End If
    '        End If
    '    End If

    'End Sub

End Module
