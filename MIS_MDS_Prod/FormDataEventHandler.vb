Module FormDataEventHandler
    Public WithEvents oApp4FormData As SAPbouiCOM.Application = Nothing

    Sub FormDataEventHandler(ByRef BoInfo As SAPbouiCOM.BusinessObjectInfo, _
                             ByRef BubbleEvent As Boolean) Handles oApp4FormData.FormDataEvent
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If BoInfo.FormTypeEx = "134" Then
                If BoInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or _
                BoInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                    If BoInfo.BeforeAction = False Then
                        '1. Get the object from UI
                        oForm = oApp.Forms.Item(BoInfo.FormUID)
                        Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("5")
                        Dim oEdit As SAPbouiCOM.EditText = oItem.Specific
                        oApp.MessageBox(String.Format("BP Code: {0}", oEdit.String))

                        '2. Get the objectkey xml
                        Dim xmlData As String = BoInfo.ObjectKey
                        '"<?xml version="1.0" encoding="UTF-16" ?><BusinessPartnerParams><CardCode>c1</CardCode></BusinessPartnerParams>"

                    End If
                End If
            End If
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub
End Module
