Module MenuEventHandler
    Public WithEvents oApp4MenuEvent As SAPbouiCOM.Application = Nothing

    Sub MenuEventHandler(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) _
    Handles oApp4MenuEvent.MenuEvent
        Try

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "PROD01_01"
                        SOToMFGEntry()
                    Case "PROD01_03"
                        OptimizationEntry()
                    Case "PROD01_04"
                        OutDelEntry()
                    Case "PROD01_05"
                        ProductionStatus()
                        '-------------- Yadi FC ----------------------------
                    Case "PROD01_06"
                        ProductionClosed()
                        '-------------- Yadi FC ----------------------------
                End Select
            End If

            If pVal.BeforeAction = True Then
                Dim oForm As SAPbouiCOM.Form

                'oForm = SBO_Application.Forms.ActiveForm
                oForm = oApp.Forms.ActiveForm
                'MsgBox(oForm.Type)
                'MsgBox(oForm.TypeEx)
                'MsgBox(oForm.UniqueID)
                Select Case pVal.MenuUID
                    Case "1290" ' 1st Record
                    Case "1289" ' Prev Record
                    Case "1288" ' Next Record

                        If oForm.UniqueID = "mds_p3" Then
                            'Dim oForm As SAPbouiCOM.Form = Nothing
                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                            Dim oDBDataSource As SAPbouiCOM.DBDataSource

                            'oForm = SBO_Application.Forms.Item(FormUID)

                            oMatrix = oForm.Items.Item("OptimMtx").Specific

                            oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                            oForm.Freeze(True)

                            Dim docnum As Integer

                            docnum = oDBDataSource.GetValue("docnum", 0)

                            Dim idx As Long
                            Dim gtabc As Double
                            Dim gtaloc As Double
                            Dim gtplanpdo As Double

                            gtabc = 0
                            gtaloc = 0
                            gtplanpdo = 0
                            For idx = 1 To oMatrix.RowCount
                                gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                                'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                                gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))
                                'oForm.Items.Item("#").Specific.value = idx
                                'oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.value = oMatrix.VisualRowCount
                                oMatrix.Columns.Item("#").Cells.Item(CInt(idx)).Specific.value = idx

                            Next

                            gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                            oForm.Items.Item("GTabc").Specific.value = gtabc
                            'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                            oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo


                            'IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) _
                            '- IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) _
                            '- IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)

                            ' 2011-02-28
                            oForm.Items.Item("GTaloc").Specific.value = _
                            IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value) _
                            + IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDBDataSource)
                            oMatrix = Nothing
                            oDBDataSource = Nothing

                            GC.Collect()


                            'BubbleEvent = False

                            oForm.Freeze(False)


                        End If

                    Case "1291" ' Last Record

                    Case "1292" ' Add a row
                        'form "mds_p3" = Optimization Entry
                        If oForm.UniqueID = "mds_p3" Then

                            ''MsgBox("Add a row optimization entry")

                            Dim oMatrix As SAPbouiCOM.Matrix
                            oMatrix = oForm.Items.Item("OptimMtx").Specific
                            oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").Clear()
                            oMatrix.AddRow()
                            oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount
                            oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            oMatrix = Nothing
                            GC.Collect()

                        End If
                        'MsgBox("Add a row")

                    Case "1293" ' Delete a row
                        'Dim oMatrix As SAPbouiCOM.Matrix
                        'oMatrix = oForm.Items.Item("OptimMtx").Specific
                        'oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").Clear()
                        'oMatrix.DeleteRow(oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string)
                        '' = oMatrix.VisualRowCount
                        ''oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount

                        'If oForm.UniqueID = "mds_p3" Then

                        '    Dim oMatrix As SAPbouiCOM.Matrix
                        '    oMatrix = oForm.Items.Item("OptimMtx").Specific
                        '    'Dim oDBDataSource As SAPbouiCOM.DBDataSource


                        '    'oMatrix = oForm.Items.Item("OptimMtx").Specific

                        '    oForm.Freeze(True)
                        '    oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").RemoveRecord(oMatrix.VisualRowCount)
                        '    oMatrix.FlushToDataSource()
                        '    'oMatrix.LoadFromDataSource()

                        '    'oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                        '    oForm.Freeze(False)


                        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        '    BubbleEvent = False

                        'End If
                        'MsgBox("Delete a row")
                    Case "1282"

                        'MsgBox("add new doc!")
                        If oForm.UniqueID = "mds_p3" Then
                            'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)

                            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                            oForm.Items.Item("GTabc").Specific.value = 0
                            oForm.Items.Item("GTaloc").Specific.value = 0
                            oForm.Items.Item("GTplanPdO").Specific.value = 0
                            oForm.Items.Item("QtyLembar").Specific.value = 1
                            oForm.Items.Item("TotalWaste").Specific.value = 0
                            oForm.Items.Item("TotWastPct").Specific.value = 0
                            oForm.Items.Item("SisaKcUtuh").Specific.value = 0
                            oForm.Items.Item("KcSisaPctg").Specific.value = 0

                            'End If
                        End If

                End Select

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                GC.Collect()

            End If

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

    End Sub

    'Sub MenuEventHandler(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) _
    'Handles oApp4MenuEvent.MenuEvent
    '    Try
    '        If pVal.MenuUID = "S_MIS_OPCR" Then
    '            If pVal.BeforeAction = True Then
    '                'CreatePurchaseRequestFormViaXml()
    '                BubbleEvent = False
    '            End If
    '        ElseIf pVal.MenuUID = "1281" Or pVal.MenuUID = "1282" Then
    '            'Switching form mode add/find, then call the manageseries
    '            Dim oForm As SAPbouiCOM.Form = oApp.Forms.ActiveForm

    '            If oForm.TypeEx = "MIS_OPCR" Then
    '                'ManageSeries(oForm, "10", "11", "Primary", "@MIS_OPCR")
    '            End If

    '        End If
    '    Catch ex As Exception
    '        MsgBoxWrapper(ex.Message)
    '    End Try
    'End Sub

End Module
