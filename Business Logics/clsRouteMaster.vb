Imports SAPbobsCOM

Public Class clsRouteMaster
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oTemp As SAPbobsCOM.Recordset
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_RouteMaster) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_RouteMaster, frm_RouteMaster)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)

        AddChooseFromList(oForm)
        oForm.DataBrowser.BrowseBy = "12"
        oEditText = oForm.Items.Item("7").Specific
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "U_DriverCode"
        If oForm.TypeEx = frm_RouteMaster Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            AddMode(oForm)
        End If
        oForm.Freeze(False)
    End Sub

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#Region "AddMode"
    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            'strCode = oApplication.Utilities.getMaxCode("@Z_ORUT", "DocEntry")
            'aForm.Items.Item("7").Enabled = True
            'oApplication.Utilities.setEdittextvalue(aForm, "5", strCode)
            'aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'aForm.Items.Item("5").Enabled = False
            oForm.Items.Item("12").Enabled = True
            aForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Items.Item("11").Enabled = False
        End If
    End Sub
#End Region

#Region "Validate details"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strECode, strEname, strQuery As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim strTypeRoute As String
        Dim strDriverCode As String

        oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strECode = oApplication.Utilities.getEdittextvalue(aForm, "12")
        strEname = oApplication.Utilities.getEdittextvalue(aForm, "13")
        strTypeRoute = oApplication.Utilities.getcomboboxvalue(aForm, "Item_1")
        strDriverCode = oApplication.Utilities.getEdittextvalue(aForm, "7")



        If strTypeRoute.Contains("None") Then
            oApplication.Utilities.Message("Route Type can not be None ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Return False
        End If
        If strECode = "" Then
            oApplication.Utilities.Message("Route Code can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If
        If strEname = "" Then
            oApplication.Utilities.Message("Route Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If
        If oApplication.Utilities.getEdittextvalue(aForm, "7") = "" Then
            oApplication.Utilities.Message("Driver Detail is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strQuery = "Select * from [@Z_ORUT] where U_RouteCode = '" & strECode & "' and U_RouteName = '" & strEname & "'"
            oRecSet.DoQuery(strQuery)
            If oRecSet.RecordCount > 0 Then
                oApplication.Utilities.Message("Duplicate couple(Route Name/ Route Code) ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            End If
        End If

        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strQuery = "Select * from  [@Z_ORUT] where U_DriverCode = '" & strDriverCode & "' and  U_TypeRoute = 'Frozen'"
            oRecSet.DoQuery(strQuery)
            Dim ScoreFrozen As Integer = oRecSet.RecordCount
            strQuery = "Select * from  [@Z_ORUT] where U_DriverCode = '" & strDriverCode & "' and  U_TypeRoute = 'Fresh'"
            oRecSet.DoQuery(strQuery)
            Dim ScoreFresh As Integer = oRecSet.RecordCount

            If ScoreFrozen = 1 And ScoreFresh = 1 Then
                oApplication.Utilities.Message("Each Driver can only be linked to two types of Routes", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False

            ElseIf (ScoreFrozen = 1 And ScoreFresh = 0) Or (ScoreFrozen = 0 And ScoreFresh = 1) Then
                If strTypeRoute.Trim = "Frozen" And ScoreFrozen = 1 Then
                    oApplication.Utilities.Message("Each Driver can only be linked to one route type frozen", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                ElseIf strTypeRoute.Trim = "Fresh" And ScoreFresh = 1 Then
                    oApplication.Utilities.Message("Each Driver can only be linked to one route type fresh", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
            End If



        ElseIf aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

            Return True
        End If
        Return True
    End Function

#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_RouteMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub

                                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        Dim ocheck As SAPbouiCOM.CheckBox = oForm.Items.Item("9").Specific
                                        oEditText = oForm.Items.Item("12").Specific
                                        Dim U_RouteCode As String = oEditText.Value

                                        If ocheck.Checked = True Then
                                            Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                            Dim query As String = "update [@Z_OCURT] set U_Active = 'Y'  where U_RouteCode = '" & U_RouteCode & "'"
                                            oRS.DoQuery(query)
                                            'update z_ocurt - U_Active


                                        Else
                                            Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                            Dim query As String = "update [@Z_OCURT] set U_Active = 'N'  where U_RouteCode = '" & U_RouteCode & "'"
                                            oRS.DoQuery(query)
                                        End If
                                    End If
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '  ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "7" Then
                                            oApplication.Utilities.setEdittextvalue(oForm, "11", oDataTable.GetValue("U_DriverName", 0))
                                            oApplication.Utilities.setEdittextvalue(oForm, "7", oDataTable.GetValue("U_DriverCode", 0))

                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If


                                        End If
                                        oForm.Freeze(False)

                                        oEditText = oForm.Items.Item("13").Specific
                                        oEditText.Item.Click()
                                        oEditText = oForm.Items.Item("12").Specific
                                        oEditText.Item.Enabled = False
                                    End If
                                Catch ex As Exception
                                    If pVal.ItemUID = "7" Then
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_RouteMaster
                    LoadForm()
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddMode(oForm)
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("12").Enabled = True
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_RouteMaster Then
                    oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("12").Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_DriverList

                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
