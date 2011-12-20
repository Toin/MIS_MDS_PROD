Module MainAddon

    Public oCompany As SAPbobsCOM.Company = Nothing
    Public lRetCode As Integer = 0
    Public lErrCode As Integer = 0
    Public sErrMsg As String = String.Empty

    Public CompDB As String

    Public WithEvents oApp As SAPbouiCOM.Application = Nothing
    Public WithEvents oEventFilters As SAPbouiCOM.EventFilters = Nothing

    Enum MsgBoxType
        WindowMsgBox = 0
        B1MsgBox = 1
        B1StatusBarMsg = 2
    End Enum


    Sub Main()
        '1. Connect via UI/DI/SSO/Multiple
        Connect()
        ''2. CreateUDTs, if not exists
        'CreateUDTs()
        ''3. RegisterUDOs. if not exists
        'RegisterUDOs()
        '4. CreateMenus.
        'CreateMenus()
        CreateAddOnMenus()

        '5. UpdateMenus.
        'UpdateMenus()

        ''6. SetFilter
        'SetFilters()

        'Dim oMainForm As Form = New MainForm
        'oMainForm.ShowDialog()

        System.Windows.Forms.Application.Run()

    End Sub

    Sub Connect()
        'If My.Settings.ConnectionType.Equals("DI") Then
        '    ConnectViaDISample()
        'ElseIf My.Settings.ConnectionType.Equals("UI") Then
        '    ConnectViaUI()
        'ElseIf My.Settings.ConnectionType.Equals("MultiAddOn") Then
        '    ConnectViaMultipleAddon()
        'Else
        '    ConnectViaSSO()
        'End If

        'ConnectViaDISample()

        'ConnectViaUI()

        'ConnectViaMultipleAddon()

        ConnectViaSSO()

    End Sub

    Sub ConnectViaDISample()
        'ConnectViaDI("toin-pc", "toin-pc:30000", SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008, _
        '             "maruni", "sa", "P@ssw0rd", "manager", "gk88", SAPbobsCOM.BoSuppLangs.ln_English)
        ConnectViaDI("toin-pc", "toin-pc:30000", SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008, _
                     "SBODemoUS", "sa", "P@ssw0rd", "manager", "1234", SAPbobsCOM.BoSuppLangs.ln_English)
    End Sub

    Sub ConnectViaDI(ByVal server As String, ByVal licSrv As String, _
                     ByVal dbType As SAPbobsCOM.BoDataServerTypes, ByVal companyDB As String, _
                     ByVal dbUser As String, ByVal dbPassword As String, _
                     ByVal userName As String, ByVal password As String, _
                     ByVal language As SAPbobsCOM.BoSuppLangs, _
                     Optional ByVal addonID As String = "")
        Try
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = server
            oCompany.LicenseServer = licSrv
            oCompany.DbServerType = dbType
            oCompany.DbUserName = dbUser
            oCompany.DbPassword = dbPassword
            oCompany.CompanyDB = companyDB
            oCompany.UserName = userName
            oCompany.Password = password
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            If String.IsNullOrEmpty(addonID) = False Then
                oCompany.AddonIdentifier = addonID
            End If

            oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

            lRetCode = oCompany.Connect

            DIErrHandler("Connectiong Company")

            CompDB = oCompany.CompanyDB

            'If lRetCode <> 0 Then
            '    oCompany.GetLastError()
            'End If

        Catch ex As Exception

        End Try
    End Sub

    Sub DIErrHandler(ByVal action As String)
        Try
            Dim msg As String

            If lRetCode = 0 Then
                msg = String.Format("{0} Succeeded", action)
            Else
                oCompany.GetLastError(lErrCode, sErrMsg)
                msg = String.Format("{0} failed. ErrCode: {1}. ErrMsg: {2}", _
                                    action, lErrCode, sErrMsg)
            End If
            MsgBoxWrapper(msg)
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub MsgBoxWrapper(ByVal msg As String, _
                      Optional ByVal msgboxType As MsgBoxType = MsgBoxType.B1StatusBarMsg, _
                      Optional ByVal msgType As SAPbouiCOM.BoStatusBarMessageType = _
                      SAPbouiCOM.BoStatusBarMessageType.smt_None)
        If Not (oApp Is Nothing) Then
            If msgboxType = MainAddon.MsgBoxType.B1MsgBox Then
                oApp.MessageBox(msg)
            ElseIf msgboxType = MainAddon.MsgBoxType.B1StatusBarMsg Then
                Dim isErr As Boolean = (msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oApp.SetStatusBarMessage(msg, SAPbouiCOM.BoMessageTime.bmt_Medium, isErr)
            Else
                MsgBox(msg)
            End If
        End If
        'MsgBox(msg)
    End Sub

    Sub ConnectViaUI()
        Try
            Dim uiAPI As SAPbouiCOM.SboGuiApi = New SAPbouiCOM.SboGuiApi
            Dim sConnStr As String = Environment.GetCommandLineArgs.GetValue(1)

            uiAPI.Connect(sConnStr)

            oApp = uiAPI.GetApplication()

            'delegate the event handler
            oApp4AppEventHandler = oApp
            oApp4ItemEvent = oApp
            oApp4FormData = oApp
            oApp4MenuEvent = oApp


            oEventFilters = New SAPbouiCOM.EventFilters

            MsgBoxWrapper("UI API Connected.", MsgBoxType.B1StatusBarMsg, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'uiAPI = Nothing

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub ConnectViaSSO()
        Try
            '1. Connect to UI
            ConnectViaUI()

            oCompany = New SAPbobsCOM.Company
            Dim sCookie As String = oCompany.GetContextCookie

            Dim connInfo As String = oApp.Company.GetConnectionContext(sCookie)

            'It will set Server, db, username, password to the DI Company
            oCompany.SetSboLoginContext(connInfo)

            lRetCode = oCompany.Connect
            MsgBoxWrapper("Connected via SSO")

            CompDB = oCompany.CompanyDB

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub ConnectViaMultipleAddon()
        Try
            ConnectViaUI()
            oCompany = oApp.Company.GetDICompany
            MsgBoxWrapper("Connected via Multiple Addon")

            CompDB = oCompany.CompanyDB

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

End Module
