Module MIS_Toolkit
    'Public WithEvents oComp As SAPbobsCOM.Company = Nothing

    'Public Function fctFormatDate(ByVal pdate As Date, ByVal oCompany As SAPbobsCOM.Company, Optional ByVal sngFormat As Integer = 5) As String
    Public Function fctFormatDate(ByVal pdate As Date, ByVal oCompany As SAPbobsCOM.Company, Optional ByVal sngFormat As Integer = 5) As String
        Dim strSeparator As String
        Dim oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        fctFormatDate = ""

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate
        strSeparator = oAdminInfo.DateSeparator

        Select Case sngFormat
            Case 0
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "yy")
            Case 1
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + "20" + Format(pdate, "yy")
            Case 2
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + Format(pdate, "yy")
            Case 3
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + "20" + Format(pdate, "yy")
            Case 4
                fctFormatDate = "20" + Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
            Case 5
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MMMM") + strSeparator + Format(pdate, "yyyy")
            Case 6
                fctFormatDate = Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
        End Select

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
    End Function

    Public Sub AddCFL1(ByVal oForm As SAPbouiCOM.Form, ByVal oLinkedObject As SAPbouiCOM.BoLinkedObject, _
            ByVal CFLtxt As String, ByVal CFLbtn As String, _
            ByVal CFLCondField As String, _
            ByVal CFLCondOperator As SAPbouiCOM.BoConditionOperation, _
            ByVal CFLCondFieldValue As String)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFLConds As SAPbouiCOM.Conditions
            Dim oCFLCond As SAPbouiCOM.Condition


            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            'oCFLCreationParams = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Add 2 CFL
            'one for button (windows popup) & one for edit textbox
            oCFLCreationParams.MultiSelection = False
            '            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            Dim oLinkedObjectType As SAPbouiCOM.BoLinkedObject
            oLinkedObjectType = oLinkedObject
            oCFLCreationParams.ObjectType = oLinkedObject ' "2"-> BP Master
            oCFLCreationParams.UniqueID = CFLtxt ' "CFL1" -> txtbox cfl Field

            oCFL = oCFLs.Add(oCFLCreationParams)

            'Add conditions to CFL1
            oCFLConds = oCFL.GetConditions()

            oCFLCond = oCFLConds.Add()
            oCFLCond.Alias = CFLCondField ' "CardType" -> BP Master where CardType = ??
            'oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCFLCond.Operation = CFLCondOperator
            oCFLCond.CondVal = CFLCondFieldValue ' "C" -> CardType value = C -> BP Customer data 
            oCFL.SetConditions(oCFLConds)

            oCFLCreationParams.UniqueID = CFLbtn ' "CFL2" -> button CFL field
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

End Module
