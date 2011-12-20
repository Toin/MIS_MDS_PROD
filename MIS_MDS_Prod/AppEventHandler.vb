Module AppEventHandler
    Public WithEvents oApp4AppEventHandler As SAPbouiCOM.Application = Nothing

    Sub AppEventHandler(ByVal eventType As SAPbouiCOM.BoAppEventTypes) _
    Handles oApp4AppEventHandler.AppEvent
        Try
            Select Case eventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    'Release the resource and exit
                    OnShutDown()
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    OnCompanyChanged()
                Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                    OnLanguageChanged()
                Case SAPbouiCOM.BoAppEventTypes.aet_FontChanged
                    OnFontChanged()
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Sub OnShutDown()
        Try
            'todo:
            '1. Close you form
            ''2. Remove your menus
            'RemoveMenus()
            '3. Exit
            MsgBoxWrapper("ShutDown", MsgBoxType.B1MsgBox)
            Application.Exit()

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub OnCompanyChanged()
        Try
            MsgBoxWrapper("Company Changed", MsgBoxType.B1MsgBox)

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub OnLanguageChanged()
        Try
            MsgBoxWrapper("Language Changed", MsgBoxType.B1MsgBox)

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub OnFontChanged()
        Try
            MsgBoxWrapper("Font Changed", MsgBoxType.B1MsgBox)

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

End Module
