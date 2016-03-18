Public Class clsType
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
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_DriverList) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Type, frm_Type)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "7"
        If oForm.TypeEx = frm_Type Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            AddMode(oForm)
        End If
        oForm.Freeze(False)
    End Sub

    Private Sub addChooseFromList(aform As SAPbouiCOM.Form)
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition

        oCFLs = aform.ChooseFromLists

        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFLCreationParams.MultiSelection = False
        oCFLCreationParams.ObjectType = "Z_ItemType"
        oCFLCreationParams.UniqueID = "ItemType"

        oCFL = oCFLs.Add(oCFLCreationParams)
    End Sub
#Region "AddMode"
    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
          
            oForm.Items.Item("7").Enabled = True
            aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
    End Sub
#End Region

#Region "Validate details"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strECode, strEname, strQuery As String
        Dim nbrdays As Integer = 0
        Dim deliverydays As Double = 0
        Dim oRecSet As SAPbobsCOM.Recordset
        oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strECode = oApplication.Utilities.getEdittextvalue(aForm, "7")


        strEname = oApplication.Utilities.getcomboboxvalue(aForm, "11")
        Dim tnbrdays As String = oApplication.Utilities.getEdittextvalue(aForm, "Item_1")
        Dim tndeliverydays As String = oApplication.Utilities.getEdittextvalue(aForm, "Item_3")

        If tnbrdays = "" Then
            oApplication.Utilities.Message("Number of days can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("Item_1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If
        If tndeliverydays = "" Then
            oApplication.Utilities.Message("Delivery days can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("Item_3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If


        nbrdays = oApplication.Utilities.getEdittextvalue(aForm, "Item_1")
        deliverydays = oApplication.Utilities.getEdittextvalue(aForm, "Item_3")

        If strECode = "" Then
            oApplication.Utilities.Message("Type Code can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If
        If strEname.trim = "None" Then
            oApplication.Utilities.Message("Prefix Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        End If




        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strQuery = "Select * from ""@Z_ItemType"" where U_TypeCode='" & strECode & "' and U_BatchNumberPrefix = '" & strEname & "'"
            oRecSet.DoQuery(strQuery)

            If oRecSet.RecordCount > 0 Then
                oApplication.Utilities.Message("This Entry already exists(TypeCode)", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            End If





        End If
        Return True
    End Function

#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Type Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                        
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

                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                    oApplication.Utilities.setEdittextvalue(oForm, "7", "")

                                    oApplication.Utilities.setEdittextvalue(oForm, "Item_1", "")
                                    oApplication.Utilities.setEdittextvalue(oForm, "Item_3", "")
                                    oEditText = oForm.Items.Item("7").Specific
                                    oEditText.Item.Click()
                                End If

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
                                        oForm.Update()
                                        oForm.Freeze(False)
                                        oEditText = oForm.Items.Item("7").Specific
                                        oEditText.Item.Enabled = False
                                    End If
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try
                        End Select
                End Select
            ElseIf pVal.FormTypeEx = frm_itemmaster Then
                Select pVal.BeforeAction

                    Case False
                        Select Case pVal.EventType


                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim objEdit As SAPbouiCOM.EditTextColumn
                                Dim oGr As SAPbouiCOM.Grid
                                Dim oItm As SAPbobsCOM.BusinessPartners
                                Dim sCHFL_ID, val, strBPCode As String
                                Try

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If (oCFLEvento.BeforeAction = False) Then


                                        If pVal.ItemUID = "oType" Then
                                            Form.Items.Item("5").Click()

                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, oDataTable.GetValue("U_TypeCode", 0))
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE


                                        End If

                                    End If
        Catch ex As Exception
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                If pVal.ItemUID = "oType" Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If
        End Try

                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'add the label in the sales order form.
                                Dim oStatic As SAPbouiCOM.StaticText = oForm.Items.Item("25").Specific
                                Dim oEditText As SAPbouiCOM.ComboBox = oForm.Items.Item("24").Specific

                                Dim oNewItem As SAPbouiCOM.Item = oForm.Items.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                                Dim oNewItem1 As SAPbouiCOM.Item = oForm.Items.Add("oType", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                ' Dim oNewItem1 As SAPbouiCOM.Item = oForm.Items.Add("comboType", 
                                ')

                                oNewItem.Left = oStatic.Item.Left
                                oNewItem.Width = oStatic.Item.Width
                                oNewItem.Height = oStatic.Item.Height
                                oNewItem.Top = oStatic.Item.Top + oStatic.Item.Height

                                oNewItem1.Left = oEditText.Item.Left
                                oNewItem1.Width = oEditText.Item.Width
                                oNewItem1.Height = oEditText.Item.Height
                                oNewItem1.Top = oEditText.Item.Top + oEditText.Item.Height

                                Dim oStatictmp As SAPbouiCOM.StaticText
                                oStatictmp = oNewItem.Specific
                                oStatictmp.Caption = "Type"

                                Dim oEdittext1 As SAPbouiCOM.EditText
                                oEdittext1 = oNewItem1.Specific

                                oEdittext1.DataBind.SetBound(True, "OITM", "U_Z_ItemType")
                                addChooseFromList(oForm)
                                oEdittext1.ChooseFromListUID = "ItemType"
                                oEdittext1.ChooseFromListAlias = "U_TypeCode"
                                '   oApplication.Utilities.AddChooseFromList(frm_itemmaster, "CFL_0", "oType", Nothing, , , )


                        End Select
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim tmp As String = oApplication.Utilities.getEdittextvalue(oForm, "oType")
                                    If tmp = "" Then
                                        oApplication.Utilities.Message("You should specify the Item Type before adding it!", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oEditText = oForm.Items.Item("oType").Specific
                                        oEditText.Item.Click()
                                        BubbleEvent = False
                                    End If
                                End If

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
                Case mnu_Type
                    LoadForm()
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddMode(oForm)
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("7").Enabled = True
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
                If oForm.TypeEx = frm_Type Then
                    ' oForm.Items.Item("11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("7").Enabled = False
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
                    Case mnu_Type


                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
