Imports SAPbobsCOM

Public Class clsCustomerRoute
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oColumn As SAPbouiCOM.Column
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private RowtoDelete As Integer
    Private oMenuobject As Object
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Private toDetele As ArrayList
    Private toAdd As ArrayList
    Dim MatrixId As Integer
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource
    Dim strQuery As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()

        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_CustomerRoute) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_CustomerRoute, frm_CustomerRoute)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        AddChooseFromList(oForm)
        oEditText = oForm.Items.Item("4").Specific
        oEditText.ChooseFromListUID = "CFL_3"
        oEditText.ChooseFromListAlias = "U_RouteCode"

        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu("1283", False)
        oForm.DataBrowser.BrowseBy = "4"
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oMatrix = oForm.Items.Item("8").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        oMatrix = oForm.Items.Item("8").Specific
        oMatrix.Columns.Item("sort").Visible = False


        toDetele = New ArrayList
        toAdd = New ArrayList

    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CustomerRoute Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oEditText = oForm.Items.Item("Item_1").Specific
                                Dim oType As String = oEditText.Value
                                oMatrix = oForm.Items.Item("8").Specific
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                                    Dim ocheck As SAPbouiCOM.CheckBox = oForm.Items.Item("7").Specific
                                    oEditText = oForm.Items.Item("4").Specific
                                    Dim U_RouteCode As String = oEditText.Value

                                    'If Validation(oForm) = False Or validatematrix(oMatrix, oType) = False Or Matrix_Validation(oForm) = False Then
                                    '    BubbleEvent = False
                                    '    Exit Sub
                                    'End If

                                    If Validation(oForm) = False Or Matrix_Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    ''------------------------------------------------------------------------------------------------------------------

                                    'Dim Tasks As New TasksClass(oType, oMatrix, oForm)
                                    'Dim Thread1 As New System.Threading.Thread( _
                                    '    AddressOf Tasks.SomeTask)

                                    ''Tasks.StrArg = "Some Arg" ' Set a field that is used as an argument
                                    'Thread1.Start() ' Start the new thread.
                                    'Thread1.Join() ' Wait for thread 1 to finish.


                                    ''--------------------------------------------------------------------------------=
                                    ''oType,oForm

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If (pVal.ItemUID = "8") And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "8"
                                    frmSourceMatrix = oMatrix
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "8"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oMatrix = oForm.Items.Item("8").Specific
                                            If oMatrix.RowCount = 0 Then
                                                oMatrix.AddRow()

                                                BubbleEvent = False

                                            End If


                                        End If

                                    Case "Item_2"

                                        oMatrix = oForm.Items.Item("8").Specific
                                        Dim i As Integer
                                        For i = 1 To oMatrix.RowCount
                                            If oMatrix.IsRowSelected(i) = True And CType(oMatrix.Columns.Item("sort").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value <> 1 Then
                                                Dim U_CardCodeselected As String = CType(oMatrix.Columns.Item("V_1").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value
                                                Dim U_CardCode As String = CType(oMatrix.Columns.Item("V_1").Cells().Item(i - 1).Specific, SAPbouiCOM.EditText).Value


                                                CType(oMatrix.Columns.Item("V_1").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value = U_CardCode
                                                CType(oMatrix.Columns.Item("V_1").Cells().Item(i - 1).Specific, SAPbouiCOM.EditText).Value = U_CardCodeselected

                                                Dim U_Activecurrent As String = CType(oMatrix.Columns.Item("V_3").Cells().Item(i).Specific, SAPbouiCOM.CheckBox).Checked
                                                Dim U_Active As String = CType(oMatrix.Columns.Item("V_3").Cells().Item(i - 1).Specific, SAPbouiCOM.CheckBox).Checked


                                                CType(oMatrix.Columns.Item("V_3").Cells().Item(i).Specific, SAPbouiCOM.CheckBox).Checked = U_Active
                                                CType(oMatrix.Columns.Item("V_3").Cells().Item(i - 1).Specific, SAPbouiCOM.CheckBox).Checked = U_Activecurrent
                                                oMatrix.SelectRow(i - 1, True, False)
                                                Exit For
                                            End If

                                        Next


                                    Case "Item_3"

                                        oMatrix = oForm.Items.Item("8").Specific
                                        Dim i As Integer
                                        Dim flag As Boolean = True

                                        For i = 1 To oMatrix.RowCount
                                            If CType(oMatrix.Columns.Item("V_1").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value = "" Then
                                                flag = False
                                            End If
                                        Next
                                        If flag = True Then
                                            For i = 1 To oMatrix.RowCount
                                                If i >= oMatrix.VisualRowCount Then

                                                Else
                                                    If oMatrix.IsRowSelected(i) = True And CType(oMatrix.Columns.Item("sort").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value <> oMatrix.VisualRowCount Then
                                                        Dim U_CardCodeselected As String = CType(oMatrix.Columns.Item("V_1").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value
                                                        Dim U_CardCode As String = CType(oMatrix.Columns.Item("V_1").Cells().Item(i + 1).Specific, SAPbouiCOM.EditText).Value


                                                        CType(oMatrix.Columns.Item("V_1").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value = U_CardCode
                                                        CType(oMatrix.Columns.Item("V_1").Cells().Item(i + 1).Specific, SAPbouiCOM.EditText).Value = U_CardCodeselected

                                                        Dim U_Activecurrent As String = CType(oMatrix.Columns.Item("V_3").Cells().Item(i).Specific, SAPbouiCOM.CheckBox).Checked
                                                        Dim U_Active As String = CType(oMatrix.Columns.Item("V_3").Cells().Item(i + 1).Specific, SAPbouiCOM.CheckBox).Checked


                                                        CType(oMatrix.Columns.Item("V_3").Cells().Item(i).Specific, SAPbouiCOM.CheckBox).Checked = U_Active
                                                        CType(oMatrix.Columns.Item("V_3").Cells().Item(i + 1).Specific, SAPbouiCOM.CheckBox).Checked = U_Activecurrent
                                                        oMatrix.SelectRow(i + 1, True, False)
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                        ElseIf flag = False Then

                                            oApplication.SBO_Application.SetStatusBarMessage("You cannot move any raw if one is empty")
                                        End If
                                        '    oForm.PaneLevel = 1
                                        'Case "11"
                                        '    oForm.PaneLevel = 2
                                        'Case "12"
                                        '    oForm.PaneLevel = 3
                                End Select
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
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    Dim index As Integer
                                    Dim count1 As Integer = 0
                                    count1 = pVal.Row

                                    If (oCFLEvento.BeforeAction = False) Then

                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "4" Then

                                            oApplication.Utilities.setEdittextvalue(oForm, "6", oDataTable.GetValue("U_RouteName", 0))

                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "4", oDataTable.GetValue("U_RouteCode", 0))

                                            Catch ex As Exception


                                                oApplication.Utilities.setEdittextvalue(oForm, "4", oDataTable.GetValue("U_RouteCode", 0))
                                                oApplication.Utilities.setEdittextvalue(oForm, "Item_1", oDataTable.GetValue("U_TypeRoute", 0))
                                                oForm.Items.Item("Item_1").Enabled = False

                                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                                                End If
                                            End Try
                                        End If
                                        If pVal.ItemUID = "8" And pVal.ColUID = "V_1" Then

                                            For index = 0 To oDataTable.Rows.Count - 1

                                                oMatrix = oForm.Items.Item("8").Specific
                                                If count1 > oMatrix.VisualRowCount Then
                                                    oMatrix.AddRow()
                                                End If
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", count1, oDataTable.GetValue("CardName", index))

                                                Try
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", count1, oDataTable.GetValue("CardCode", index))
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", count1, count1)
                                                    CType(oMatrix.Columns.Item("V_3").Cells().Item(introw).Specific, SAPbouiCOM.CheckBox).Checked = True

                                                    ' oApplication.Utilities.SetMatrixValues(oMatrix, "sort", count, count)
                                                    'oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", count, "Y")
                                                Catch ex As Exception
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", count1, oDataTable.GetValue("CardCode", index))
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", count1, count1)
                                                    'oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", count1, "Y")
                                                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                                                    End If
                                                End Try
                                                count1 += 1
                                            Next


                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try

                        End Select
                End Select

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try

            Select Case pVal.MenuUID
                Case mnu_CustomerRoute
                    LoadForm()
                Case mnu_ADD_ROW

                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Try
                        If pVal.BeforeAction = False Then
                            Exit Sub
                        End If

                        AddRow(oForm)
                    Catch ex As Exception

                    End Try


                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Try

                        If pVal.BeforeAction = False Then
                            AddMode(oForm)
                        End If

                    Catch ex As Exception

                    End Try
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Try


                        If pVal.BeforeAction = False Then

                            oForm.Items.Item("4").Enabled = True
                            oForm.Items.Item("6").Enabled = False
                            oForm.Items.Item("8").Enabled = True
                        End If
                    Catch ex As Exception

                    End Try

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Try
                        oEditText = oForm.Items.Item("oFresh").Specific
                        If pVal.BeforeAction = False Then
                            oEditText = oForm.Items.Item("5").Specific
                            Dim cardcode As String = oEditText.Value
                            Dim query As String = "select * from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t0.DocEntry = t1.DocEntry  where U_CardCode = '" & cardcode & "' and U_TypeRoute = 'Fresh' and t0.U_Active = 'Y' "
                            Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery(query)
                            oEditText = oForm.Items.Item("oFresh").Specific
                            oEditText.Value = oRs.Fields.Item("U_RouteCode").Value & "-" & oRs.Fields.Item("U_RouteName").Value
                            oForm.Items.Item("7").Click()
                            oEditText.Item.Enabled = False

                            oEditText = oForm.Items.Item("5").Specific
                            cardcode = oEditText.Value
                            query = "select * from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t0.DocEntry = t1.DocEntry  where U_CardCode = '" & cardcode & "' and U_TypeRoute = 'Frozen' and t0.U_Active = 'Y' "
                            oRs.DoQuery(query)
                            oEditText = oForm.Items.Item("oFrozen").Specific
                            oEditText.Value = oRs.Fields.Item("U_RouteCode").Value & "-" & oRs.Fields.Item("U_RouteName").Value
                            oForm.Items.Item("7").Click()
                            oEditText.Item.Enabled = False

                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                oForm.Items.Item("1").Click()
                                oEditText = oForm.Items.Item("oFrozen").Specific
                                oEditText.Item.Enabled = False
                                oEditText = oForm.Items.Item("oFresh").Specific
                                oEditText.Item.Enabled = False
                            End If
                        ElseIf pVal.BeforeAction = True Then


                        End If

                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("8").Enabled = True
                    Catch ex As Exception
                    End Try

            End Select

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

            Select Case BusinessObjectInfo.BeforeAction
                Case True
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                            oForm.Items.Item("4").Enabled = False
                            oForm.Items.Item("6").Enabled = False
                            oForm.Items.Item("8").Enabled = True
                    End Select
                Case False
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                            Dim oRecordSet As SAPbobsCOM.Recordset
                            oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                            Dim DocEntry As String = oXmlDoc.SelectSingleNode("/Customer_Route_MappingParams/DocEntry").InnerText
                            oRecordSet = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                            strQuery = "Select * From [@Z_OCURT] Where DocEntry = '" + DocEntry + "'"
                            oRecordSet.DoQuery(strQuery)
                            If Not oRecordSet.EoF Then
                                Dim strType As String = oRecordSet.Fields.Item("U_TypeRoute").Value.ToString

                                If strType = "Fresh" Then 'Update Fresh


                                    strQuery = " Update T0 SET T0.U_freshroute = '' From OCRD T0  JOIN [@Z_OCURT] T2 On T0.U_freshroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "'"
                                    oRecordSet.DoQuery(strQuery)



                                    strQuery = " Update T0 SET T0.U_freshroute = '', T0.U_Dfresh = '' From ORDR T0  JOIN [@Z_OCURT] T2 On T0.U_freshroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "' and T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)


                                    strQuery = " Update T0 SET T0.U_freshroute = '', T0.U_Dfresh = '' From OINV T0  JOIN [@Z_OCURT] T2 On T0.U_freshroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "' and T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                                    strQuery = " Update T0 SET T0.U_freshroute = '', T0.U_Dfresh = '' From ODLN T0  JOIN [@Z_OCURT] T2 On T0.U_freshroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "' and T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                     
                                    'Update All Customers
                                    strQuery = " Update T0 SET T0.U_freshroute = T2.U_RouteCode "
                                    strQuery += " From OCRD T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    oRecordSet.DoQuery(strQuery)

                                    'Update All Open Orders
                                    strQuery = " Update T0 SET T0.U_freshroute = T2.U_RouteCode,T0.U_Dfresh = T3.U_DriverCode "
                                    strQuery += " From ORDR T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " JOIN [@Z_ORUT] T3 On T2.U_RouteCode = T3.U_RouteCode  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    strQuery += " And T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)






                                    'Update All Open Delivery
                                    strQuery = " Update T0 SET T0.U_freshroute = T2.U_RouteCode,T0.U_Dfresh = T3.U_DriverCode "
                                    strQuery += " From ODLN T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " JOIN [@Z_ORUT] T3 On T2.U_RouteCode = T3.U_RouteCode  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    strQuery += " And T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                                    'Update All Open Invoice
                                    strQuery = " Update T0 SET T0.U_freshroute = T2.U_RouteCode,T0.U_Dfresh = T3.U_DriverCode "
                                    strQuery += " From OINV T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " JOIN [@Z_ORUT] T3 On T2.U_RouteCode = T3.U_RouteCode  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    strQuery += " And T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                                ElseIf strType = "Frozen" Then 'Update Frozen


                                    strQuery = " Update T0 SET T0.U_frozenroute = '' From OCRD T0  JOIN [@Z_OCURT] T2 On T0.U_frozenroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "'"
                                    oRecordSet.DoQuery(strQuery)

                                    strQuery = " Update T0 SET T0.U_frozenroute = '', T0.U_Dfrozen = '' From ORDR T0  JOIN [@Z_OCURT] T2 On T0.U_frozenroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "' and T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)



                                    strQuery = " Update T0 SET T0.U_frozenroute = '', T0.U_Dfrozen = '' From OINV T0  JOIN [@Z_OCURT] T2 On T0.U_frozenroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "' and T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)


                                    strQuery = " Update T0 SET T0.U_frozenroute = '', T0.U_Dfrozen = '' From ODLN T0  JOIN [@Z_OCURT] T2 On T0.U_frozenroute = T2.U_RouteCode Where T2.DocEntry = '" + DocEntry + "' and T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)



                                    'Update All Customers
                                    strQuery = " Update T0 SET T0.U_frozenroute = T2.U_RouteCode "
                                    strQuery += " From OCRD T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    oRecordSet.DoQuery(strQuery)

                                    'Update All Open Orders
                                    strQuery = " Update T0 SET T0.U_frozenroute = T2.U_RouteCode,T0.U_Dfrozen = T3.U_DriverCode "
                                    strQuery += " From ORDR T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " JOIN [@Z_ORUT] T3 On T2.U_RouteCode = T3.U_RouteCode  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    strQuery += " And T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                                    'Update All Open Delivery
                                    strQuery = " Update T0 SET T0.U_frozenroute = T2.U_RouteCode,T0.U_Dfrozen = T3.U_DriverCode "
                                    strQuery += " From ODLN T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " JOIN [@Z_ORUT] T3 On T2.U_RouteCode = T3.U_RouteCode  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    strQuery += " And T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                                    'Update All Open Invoice
                                    strQuery = " Update T0 SET T0.U_frozenroute = T2.U_RouteCode,T0.U_Dfrozen = T3.U_DriverCode "
                                    strQuery += " From OINV T0 JOIN [@Z_CURT1] T1 ON T0.CardCode = T1.U_CardCode  "
                                    strQuery += " JOIN [@Z_OCURT] T2 On T1.DocEntry = T2.DocEntry  "
                                    strQuery += " JOIN [@Z_ORUT] T3 On T2.U_RouteCode = T3.U_RouteCode  "
                                    strQuery += " Where T2.DocEntry = '" + DocEntry + "'"
                                    strQuery += " And T0.DocStatus = 'O'"
                                    oRecordSet.DoQuery(strQuery)

                                End If

                            End If
                    End Select
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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


            oCFL = oCFLs.Item("CFL_3") 'it is bugging here ! 

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "AddMode"
    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String
        Try
            aForm.Freeze(True)

            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                Try
                    oForm.Items.Item("4").Enabled = True
                    oForm.Items.Item("6").Enabled = False
                    oForm.Items.Item("8").Enabled = True
                    oForm.Items.Item("7").Enabled = True
                Catch ex As Exception

                End Try
                oMatrix = aForm.Items.Item("8").Specific
                oMatrix.FlushToDataSource()
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
                For count = 1 To oDataSrc_Line.Size - 1
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()
            End If
            aForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Validations"
    ' check if the  customer exists...

    Private Function validatematrix(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal oType As String) As Boolean
        Dim strCode, strCode1, strName, strEname As String

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            oEditText = oForm.Items.Item("4").Specific
            Dim routecode As String = oEditText.Value

            For intRow As Integer = 1 To aMatrix.RowCount
                strCode = CType(aMatrix.Columns.Item("V_1").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                'Dim querytoremove As String = "delete from [@Z_CURT1] where DocEntry = (select DocEntry from [@Z_OCURT] where U_RouteCode = '" & routecode & "')"
                ' oRS.DoQuery(querytoremove)

                ' if fresh and froxen
                Dim query As String = "select * from  [@Z_CURT1] t0 inner join [@Z_OCURT] t1 on t1.DocEntry = t0.DocEntry where t0.U_CardCode = '" & strCode & "' and t1.U_TypeRoute = 'Frozen' and  t1.U_Active='Y' and t1.U_RouteCode <> '" & routecode & "'"

                oRS.DoQuery(query)

                query = "select * from  [@Z_CURT1] t0 inner join [@Z_OCURT] t1 on t1.DocEntry = t0.DocEntry where t0.U_CardCode = '" & strCode & "' and t1.U_TypeRoute = 'Fresh' and  t1.U_Active='Y' and t1.U_RouteCode <> '" & routecode & "'"
                Dim oRS1 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRS1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRS1.DoQuery(query)
                If (oRS.RecordCount >= 1 And oRS1.RecordCount >= 1) Then
                    oApplication.Utilities.Message("Each customer can be linked to two routes(Fresh/Frozen): " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False

                ElseIf (oRS.RecordCount >= 1 And oRS1.RecordCount = 0) Or (oRS.RecordCount = 0 And oRS1.RecordCount >= 1) Then
                    If oType.Trim = "Frozen" And oRS.RecordCount = 1 Then
                        oApplication.Utilities.Message("Each Client can only be linked to one route type frozen", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf oType.Trim = "Fresh" And oRS1.RecordCount = 1 Then
                        oApplication.Utilities.Message("Each Driver can only be linked to one route type fresh", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If

                'For intLoop As Integer = intRow + 1 To aMatrix.RowCount
                '    strCode1 = CType(aMatrix.Columns.Item("V_1").Cells().Item(intLoop).Specific, SAPbouiCOM.EditText).Value
                '    If strCode1 <> "" Then
                '        strEname = CType(aMatrix.Columns.Item("V_1").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                '        If strCode.ToUpper() = strCode1.ToUpper() Then
                '            oApplication.Utilities.Message("This Customer already exists : " & strCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '            aMatrix.Columns.Item("V_1").Cells().Item(intRow).Click()
                '            Return False
                '        End If

                '    End If
                'Next

            Next
        End If
        Return True
    End Function

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strsubfee, strMAfee As Integer
        aForm.Freeze(True)
        If oApplication.Utilities.getEdittextvalue(oForm, "4") = "" Then
            oApplication.Utilities.Message("Route Code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End If
        Dim strCode As String = oApplication.Utilities.getEdittextvalue(aForm, "4")
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            Dim strterms, strLeavecode As String
            AddMode(aForm)
            strterms = oApplication.Utilities.getEdittextvalue(oForm, "4")
            otemp.DoQuery("Select * from ""@Z_OCURT"" where ""U_RouteCode""='" & strterms & "'")
            If otemp.RecordCount > 0 Then
                oApplication.Utilities.Message("This Entry already exists... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
        End If
        oMatrix = aForm.Items.Item("8").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Customer details are missing... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End If
        oMatrix.FlushToDataSource()
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        For count = 1 To oDataSrc_Line.Size
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        aForm.Freeze(False)
        Return True
    End Function

    Private Function Matrix_Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strType, strValue, strCode As String
        oMatrix = aForm.Items.Item("8").Specific

        For intRow As Integer = 1 To oMatrix.RowCount
            strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
            If strCode = "" Then
                oApplication.Utilities.Message("Customer is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub RefereshRowLineValues(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("8").Specific
            oMatrix.FlushToDataSource()
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
            For count = 1 To oDataSrc_Line.Size - 1
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
        Catch ex As Exception

        End Try
    End Sub

    Private Function CheckDuplicate(ByVal aCode As String, ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from ""@Z_OCURT"" where ""U_RouteCode""='" & aCode & "'")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.Message("This entry already exists .....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return True
        End If
        Return False
    End Function

#End Region

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)

        oMatrix = aForm.Items.Item("8").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        Try
            aForm.Freeze(True)
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If
            Dim intRowCount As Integer = 1
            If intSelectedMatrixrow > 0 Then
                oEditText = oMatrix.Columns.Item("V_1").Cells.Item(intSelectedMatrixrow).Specific
                If oEditText.String <> "" Then
                    oMatrix.AddRow(1, intSelectedMatrixrow)
                    oMatrix.ClearRowData(intSelectedMatrixrow + 1)
                    ' oMatrix.Columns.Item("V_1").Cells.Item(intSelectedMatrixrow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    intRowCount = intSelectedMatrixrow + 1
                End If
            Else
                oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                If oEditText.String <> "" Then
                    oMatrix.AddRow()
                    oMatrix.ClearRowData(oMatrix.RowCount)
                    ' oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    intRowCount = oMatrix.RowCount
                End If
            End If

            Try

            Catch ex As Exception
                aForm.Freeze(False)
                oMatrix.AddRow()
            End Try
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.Columns.Item("V_1").Cells.Item(intRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
            aForm.Freeze(False)
        Catch ex As Exception
            ' oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#End Region

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)

        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        frmSourceMatrix = aForm.Items.Item("8").Specific
        If intSelectedMatrixrow <= 0 Then
            Exit Sub
        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
        aForm.Freeze(False)

    End Sub

    Private Sub DeleteRow(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)

        oMatrix = aform.Items.Item("8").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
            End If
        Next
        aform.Freeze(False)
    End Sub

    Class TasksClass

        Private oForm As SAPbouiCOM.Form
        Private oEditText As SAPbouiCOM.EditText
        Private oMatrix As SAPbouiCOM.Matrix
        Private oType As String

        Sub New(oType As String, oMatrix As SAPbouiCOM.Matrix, oForm As SAPbouiCOM.Form)
            Me.oForm = oForm
            Me.oMatrix = oMatrix
            Me.oType = oType
        End Sub

        Sub SomeTask()

            Try



                Dim ocheck As SAPbouiCOM.CheckBox = oForm.Items.Item("7").Specific
                oEditText = oForm.Items.Item("4").Specific
                Dim U_RouteCode As String = oEditText.Value

                If ocheck.Checked = True Then
                    Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    Dim query As String = "update [@Z_ORUT] set U_Active = 'Y'  where U_RouteCode = '" & U_RouteCode & "'"
                    oRS.DoQuery(query)
                    'update z_ocurt - U_Active


                Else
                    Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    Dim query As String = "update [@Z_ORUT] set U_Active = 'N'  where U_RouteCode = '" & U_RouteCode & "'"
                    oRS.DoQuery(query)
                End If


                ' -------------------------------------------------------------------------------

                ' Dim Thread1 As New System.Threading.Thread(

                Dim oRS2 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                Dim oRS8 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                oEditText = oForm.Items.Item("4").Specific
                Dim routecode As String = oEditText.Value

                Dim query3 As String = ""
                Dim query4 As String = ""
                Dim query5 As String = ""
                Dim query6 As String = ""


                query3 = "select * from [@Z_ORUT] where U_RouteCode='" & routecode & "'"
                oRS8.DoQuery(query3)

                If oType = "Frozen" Then

                    Dim querytotakeover As String = "update OCRD set U_frozenroute='' where U_frozenroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)
                    querytotakeover = "update ORDR set U_frozenroute='' , U_Dfrozen='' where docstatus = 'O' and  U_frozenroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)
                    querytotakeover = "update OINV set U_frozenroute='' , U_Dfrozen=''  where docstatus = 'O' and  U_frozenroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)
                    querytotakeover = "update ODLN set  U_frozenroute='' , U_Dfrozen='' where docstatus = 'O' and  U_frozenroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)



                    query3 = "update OCRD set U_frozenroute = '" & routecode & "' "
                    query4 = "update ORDR set U_frozenroute = '" & routecode & "' , U_Dfrozen= '" & oRS8.Fields.Item("U_DriverCode").Value & "' "
                    query5 = "update OINV set U_frozenroute = '" & routecode & "' , U_Dfrozen= '" & oRS8.Fields.Item("U_DriverCode").Value & "' "
                    query6 = "update ODLN set U_frozenroute = '" & routecode & "' , U_Dfrozen= '" & oRS8.Fields.Item("U_DriverCode").Value & "' "

                ElseIf oType = "Fresh" Then


                    Dim querytotakeover As String = "update OCRD set  U_freshroute='' where U_freshroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)
                    querytotakeover = "update ORDR set U_freshroute='',U_Dfresh='' where docstatus = 'O' and  U_freshroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)
                    querytotakeover = "update OINV set U_freshroute='',U_Dfresh='' where docstatus = 'O' and  U_freshroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)
                    querytotakeover = "update ODLN set U_freshroute='',U_Dfresh='' where docstatus = 'O' and  U_freshroute='" & routecode & "'"
                    oRS2.DoQuery(querytotakeover)

                    query3 = "update OCRD set U_freshroute = '" & routecode & "'"
                    query4 = "update ORDR set U_freshroute = '" & routecode & "' , U_Dfresh= '" & oRS8.Fields.Item("U_DriverCode").Value & "'"
                    query5 = "update OINV set U_frozenroute = '" & routecode & "' , U_Dfrozen= '" & oRS8.Fields.Item("U_DriverCode").Value & "'"
                    query6 = "update ODLN set U_frozenroute = '" & routecode & "' , U_Dfrozen= '" & oRS8.Fields.Item("U_DriverCode").Value & "'"

                End If

                Dim z As Integer
                For z = 1 To oMatrix.VisualRowCount
                    oRS2.DoQuery(query3 & " where CardCode ='" & CType(oMatrix.Columns.Item("V_1").Cells().Item(z).Specific, SAPbouiCOM.EditText).Value & "'")
                    oRS2.DoQuery(query4 & " where CardCode ='" & CType(oMatrix.Columns.Item("V_1").Cells().Item(z).Specific, SAPbouiCOM.EditText).Value & "'")
                    oRS2.DoQuery(query5 & " where CardCode ='" & CType(oMatrix.Columns.Item("V_1").Cells().Item(z).Specific, SAPbouiCOM.EditText).Value & "'")
                    oRS2.DoQuery(query6 & " where CardCode ='" & CType(oMatrix.Columns.Item("V_1").Cells().Item(z).Specific, SAPbouiCOM.EditText).Value & "'")
                Next
                '-----------


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


        End Sub


    End Class

End Class
