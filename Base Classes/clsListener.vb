Public Class clsListener
    Inherits Object
    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter
#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property
#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)
            objFilter.AddEx("0")

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)
            objFilter.AddEx(frm_sales)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            objFilter.AddEx(frm_sales)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            objFilter.AddEx(frm_itemmaster)
            objFilter.AddEx(frm_BPMaster)
            objFilter.AddEx(frm_SalesOrder)
            objFilter.AddEx(frm_PurchaseOrder)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            objFilter.AddEx(frm_itemmaster)
            objFilter.AddEx(frm_BPMaster)
            objFilter.AddEx(frm_SalesOrder)
            objFilter.AddEx(frm_PurchaseOrder)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            Case frm_DriverList
                Dim objInvoice As clsDriverMaster
                objInvoice = New clsDriverMaster
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_RouteMaster
                Dim objInvoice As clsRouteMaster
                objInvoice = New clsRouteMaster
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_CustomerRoute
                Dim objInvoice As clsCustomerRoute
                objInvoice = New clsCustomerRoute
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_BPMaster
                Dim objInvoice As clsBP
                objInvoice = New clsBP
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_Type
                Dim objInvoice As clsType
                objInvoice = New clsType
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_sales
                Dim objInvoice As clsSalesOrder
                objInvoice = New clsSalesOrder
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_Delivery
                Dim objInvoice As clsDelivery
                objInvoice = New clsDelivery
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_Delivery2
                Dim objInvoice As clasDelivered
                objInvoice = New clasDelivered
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_SalesOrder
                Dim objInvoice As clsSalesOrderSystem
                objInvoice = New clsSalesOrderSystem
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_ReserveInvoice
                Dim objInvoice As clsReserveInvoice
                objInvoice = New clsReserveInvoice
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        End Select
        '  End If
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID

                    Case mnu_DriverList
                        oMenuObject = New clsDriverMaster
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_RouteMaster
                        oMenuObject = New clsRouteMaster
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_CustomerRoute
                        oMenuObject = New clsCustomerRoute
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Type
                        oMenuObject = New clsType
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_sales
                        oMenuObject = New clsSalesOrder
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Delivery
                        oMenuObject = New clsDelivery
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Delivery2
                        oMenuObject = New clasDelivered
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If

                End Select

            Else
                Select Case pVal.MenuUID
                    Case mnu_CLOSE
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                End Select

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub
#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID


            If pVal.Before_Action = True Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_ALL_EVENTS Then
                    If pVal.FormTypeEx = "-139" Then
                        If pVal.ItemUID = "U_Dfresh" Or pVal.ItemUID = "U_Dfrozen" _
                            Or pVal.ItemUID = "U_freshroute" Or pVal.ItemUID = "U_frozenroute" _
                            Or pVal.ItemUID = "U_salesperson" Or pVal.ItemUID = "U_datetiming" _
                            Or pVal.ItemUID = "U_typepayment" Or pVal.ItemUID = "U_deliverydate" _
                            Or pVal.ItemUID = "U_TypeRoute" Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If pVal.Before_Action = True Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.FormTypeEx = "-60091" Then
                        If pVal.ItemUID = "U_Dfresh" Or pVal.ItemUID = "U_Dfrozen" _
                            Or pVal.ItemUID = "U_freshroute" Or pVal.ItemUID = "U_frozenroute" _
                            Or pVal.ItemUID = "U_salesperson" Or pVal.ItemUID = "U_datetiming" _
                            Or pVal.ItemUID = "U_typepayment" Or pVal.ItemUID = "U_deliverydate" _
                            Or pVal.ItemUID = "U_TypeRoute" Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If pVal.Before_Action = True Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.FormTypeEx = "-140" Then
                        If pVal.ItemUID = "U_Dfresh" Or pVal.ItemUID = "U_Dfrozen" _
                           Or pVal.ItemUID = "U_freshroute" Or pVal.ItemUID = "U_frozenroute" _
                           Or pVal.ItemUID = "U_salesperson" Or pVal.ItemUID = "U_datetiming" _
                           Or pVal.ItemUID = "U_typepayment" Or pVal.ItemUID = "U_deliverydate" _
                           Or pVal.ItemUID = "U_TypeRoute" Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If


            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                Select Case pVal.FormTypeEx
                    Case "0"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New SystemMessage
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_DriverList
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDriverMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_RouteMaster
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsRouteMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_CustomerRoute
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCustomerRoute
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Type
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsType
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "65213", frm_Issue_Inventory, frm_InventoryTransfer
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProduction
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "42"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProduction
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "41"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New batchessetup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "720"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProduction
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_itemmaster

                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsType
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "65214"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New batchessetup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_sales
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSalesOrder
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SalesOrder
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSalesOrderSystem
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_BPMaster
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBP
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Delivery
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDelivery
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ReserveInvoice
                        oItemObject = New clsReserveInvoice
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Delivery2
                        oItemObject = New clasDelivered
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)


                    Case frm_DeliverySystem
                        oItemObject = New clsDeliverySystem
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                End Select


            End If


            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx
                    'Case frm_FuturaSetup
                    '    If Not _Collection.ContainsKey(FormUID) Then
                    '        oItemObject = New clsAcctMapping
                    '        oItemObject.FrmUID = FormUID
                    '        _Collection.Add(FormUID, oItemObject)
                    '    End If
                    'Case frm_Import
                    '    If Not _Collection.ContainsKey(FormUID) Then
                    '        oItemObject = New clsImport
                    '        oItemObject.FrmUID = FormUID
                    '        _Collection.Add(FormUID, oItemObject)
                    '    End If
                    'Case frm_Export
                    '    If Not _Collection.ContainsKey(FormUID) Then
                    '        oItemObject = New clsExport
                    '        oItemObject.FrmUID = FormUID
                    '        _Collection.Add(FormUID, oItemObject)
                    '    End If
                End Select
            End If

            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

End Class
