Module MenuCreation

    Sub CreateAddOnMenus()

        Try
            LoadFromXML_Menu("MDSProdMenus.xml")

        Catch ex As Exception
            'SBO_Application.MessageBox(ex.Message)
            MsgBoxWrapper(ex.Message)
        End Try

        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Function LoadFromXML(ByVal FileName As String) As String

        Dim oXmlDoc As Xml.XmlDocument
        Dim sPath As String

        oXmlDoc = New Xml.XmlDocument

        '// load the content of the XML File

        sPath = System.Windows.Forms.Application.StartupPath
        ''remove dir BIN
        'sPath = sPath.Remove(sPath.Length - 3, 3)

        'sPath = "E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK\"

        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString
        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString & "\"

        oXmlDoc.Load(sPath & "\" & FileName)

        '// load the form to the SBO application in one batch
        Return (oXmlDoc.InnerXml)

        'oXmlDoc = Nothing
        'sPath = Nothing
        'GC.Collect()
    End Function

    Sub LoadFromXML_Menu(ByVal FileName As String)
        'method Trial for adding menu using xml

        Dim oXmlDoc As Xml.XmlDocument

        oXmlDoc = New Xml.XmlDocument

        '// load the content of the XML File
        Dim sPath As String

        '        sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        ' Check build output path; remove the bin

        sPath = System.Windows.Forms.Application.StartupPath
        ' Check build output path; remove directory the "bin" to get app root path 
        '   e.g: E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK\bin
        'sPath = sPath.Remove(sPath.Length - 3, 3)


        '' Or
        '' Get Startup app path directory e.g: E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK
        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString

        '        sPath = "E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK\"
        oXmlDoc.Load(sPath & "\" & FileName)

        '' e.g Adding Menu
        ''// load the form to the SBO application in one batch
        'SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
        'sPath = SBO_Application.GetLastBatchResults()
        oApp.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = oApp.GetLastBatchResults()

        'MsgBox(GC.GetTotalMemory(True))
        'oXmlDoc = Nothing
        'sPath = Nothing

        ''not compatible to release oxmldoc using releaseComObject
        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc)
        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(sPath)

        'GC.Collect()

    End Sub



    Function CreateMenuItem(ByVal parentUID As String, ByVal type As SAPbouiCOM.BoMenuType, _
                            ByVal menuUID As String, _
                            ByVal menuString As String, _
                            Optional ByVal position As Integer = 0, _
                            Optional ByVal image As String = "") As SAPbouiCOM.MenuItem
        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
        Try
            '1. Get the Parent menu
            Dim oMenus As SAPbouiCOM.Menus = oApp.Menus
            If oMenus.Exists(parentUID) = False Then
                MsgBoxWrapper("Parent Menu not Found")
                Return Nothing
                Exit Function
            End If

            '2. Check the menu already exists or not
            If oMenus.Exists(menuUID) Then
                oMenus.RemoveEx(menuUID)
            End If

            '3 Add the menu
            Dim oParentMenu As SAPbouiCOM.MenuItem = oMenus.Item(parentUID)
            oMenus = oParentMenu.SubMenus

            Dim oMCP As SAPbouiCOM.MenuCreationParams = _
            oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            oMCP.Type = type
            oMCP.UniqueID = menuUID
            oMCP.String = menuString
            oMCP.Position = position
            oMCP.Image = image

            Try
                oMenuItem = oMenus.AddEx(oMCP)
            Catch ex As Exception
                MsgBoxWrapper(ex.Message)
            End Try

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

        Return oMenuItem
    End Function

    Sub CreateMenus()
        CreateMenuItem("43520", SAPbouiCOM.BoMenuType.mt_POPUP, "P_MIS_OPCR", "Purchase Request", 13)
        CreateMenuItem("P_MIS_OPCR", SAPbouiCOM.BoMenuType.mt_STRING, "S_MIS_OPCR", "Purchase Request", 0)

    End Sub

    Sub RemoveMenus()
        RemoveMenuItem("P_MIS_OPCR")
        RemoveMenuItem("S_MIS_OPCR")

    End Sub

    Sub RemoveMenuItem(ByVal menuUID As String)
        Try
            Dim oMenus As SAPbouiCOM.Menus = oApp.Menus
            If oMenus.Exists(menuUID) Then
                oMenus.RemoveEx(menuUID)
            End If
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

    End Sub

    Sub ExportMenuAsXmlFile(ByVal menuUID As String)
        Try
            Dim oMenus As SAPbouiCOM.Menus = oApp.Menus
            If oMenus.Exists(menuUID) = False Then
                MsgBoxWrapper(String.Format("Menu {0} not found", menuUID))
            End If

            Dim oMenuItem As SAPbouiCOM.MenuItem = oMenus.Item(menuUID)
            Dim xmlMenu As String = oMenuItem.SubMenus.GetAsXML()
            Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument
            xmlDoc.LoadXml(xmlMenu)

            Dim filePath As String = String.Format("{0}\{1}.xml", _
                                                   Application.StartupPath, menuUID)
            xmlDoc.Save(filePath)
            DIErrHandler(String.Format("Menu {0} saved as xml file - {1}", oMenuItem.String, filePath))
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub UpdateMenus()
        Try
            Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument

            'xmlDoc.LoadXml("43520.xml")
            'xmlDoc.LoadXml("2048.xml")


            'oApp.LoadBatchActions(xmlDoc.InnerXml)
            oApp.LoadBatchActions(IO.File.ReadAllText("43520.xml"))
            Dim result As String = oApp.GetLastBatchResults

            'hide the banking menu
            oApp.Menus.RemoveEx("43537")

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub
End Module
