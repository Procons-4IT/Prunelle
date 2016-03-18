Public Class clsSalesOrderSystem

    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private CardCode As String = ""
    Public Shared flagitem As Boolean = False
    Public Shared flagmatrix As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub createEditText(oForm As SAPbouiCOM.Form)



        Dim oStatic As SAPbouiCOM.StaticText = oForm.Items.Item("135").Specific


        oEditText = oForm.Items.Item("134").Specific

        Dim oNewItem1 As SAPbouiCOM.Item = oForm.Items.Add("oFresh", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        Dim oNewItem3 As SAPbouiCOM.Item = oForm.Items.Add("oFrozen", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        Dim oNewItem4 As SAPbouiCOM.Item = oForm.Items.Add("oSalesp", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        Dim oNewItem5 As SAPbouiCOM.Item = oForm.Items.Add("otime", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        Dim oNewItem6 As SAPbouiCOM.Item = oForm.Items.Add("oDFresh", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        Dim oNewItem7 As SAPbouiCOM.Item = oForm.Items.Add("oDFrozen", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        Dim oNewItem8 As SAPbouiCOM.Item = oForm.Items.Add("otype", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)

        oNewItem1.FromPane = 7
        oNewItem1.ToPane = 7
        oNewItem3.FromPane = 7
        oNewItem3.ToPane = 7
        oNewItem4.FromPane = 7
        oNewItem4.ToPane = 7
        oNewItem5.FromPane = 7
        oNewItem5.ToPane = 7
        oNewItem6.FromPane = 7
        oNewItem6.ToPane = 7
        oNewItem7.FromPane = 7
        oNewItem7.ToPane = 7
        oNewItem8.FromPane = 7
        oNewItem8.ToPane = 7

        Dim oEditText1 As SAPbouiCOM.EditText
        Dim oEditText2 As SAPbouiCOM.EditText
        Dim oEditText3 As SAPbouiCOM.EditText
        Dim oEditText4 As SAPbouiCOM.EditText
        Dim oEditText5 As SAPbouiCOM.EditText
        Dim oEditText6 As SAPbouiCOM.EditText
        Dim oEditText7 As SAPbouiCOM.EditText

        oNewItem1.Left = oEditText.Item.Left
        oNewItem1.Width = oEditText.Item.Width
        oNewItem1.Height = oEditText.Item.Height
        oNewItem1.Top = oEditText.Item.Top + oEditText.Item.Height
        oEditText1 = oNewItem1.Specific
        oEditText1.DataBind.SetBound(True, "ORDR", "U_freshroute")

        oNewItem3.Left = oEditText1.Item.Left
        oNewItem3.Width = oEditText1.Item.Width
        oNewItem3.Height = oEditText1.Item.Height
        oNewItem3.Top = oEditText1.Item.Top + oEditText1.Item.Height
        oEditText2 = oNewItem3.Specific
        oEditText2.DataBind.SetBound(True, "ORDR", "U_frozenroute")


        oNewItem4.Left = oEditText2.Item.Left
        oNewItem4.Width = oEditText2.Item.Width
        oNewItem4.Height = oEditText2.Item.Height
        oNewItem4.Top = oEditText2.Item.Top + oEditText2.Item.Height
        oEditText3 = oNewItem4.Specific
        oEditText3.DataBind.SetBound(True, "ORDR", "U_salesperson")

        oNewItem5.Left = oEditText3.Item.Left
        oNewItem5.Width = oEditText3.Item.Width
        oNewItem5.Height = oEditText3.Item.Height
        oNewItem5.Top = oEditText3.Item.Top + oEditText3.Item.Height
        oEditText4 = oNewItem5.Specific
        oEditText4.DataBind.SetBound(True, "ORDR", "U_datetiming")


        oNewItem6.Left = oEditText4.Item.Left
        oNewItem6.Width = oEditText4.Item.Width
        oNewItem6.Height = oEditText4.Item.Height
        oNewItem6.Top = oEditText4.Item.Top + oEditText4.Item.Height
        oEditText5 = oNewItem6.Specific
        oEditText5.DataBind.SetBound(True, "ORDR", "U_Dfresh")

        oNewItem7.Left = oEditText5.Item.Left
        oNewItem7.Width = oEditText5.Item.Width
        oNewItem7.Height = oEditText5.Item.Height
        oNewItem7.Top = oEditText5.Item.Top + oEditText5.Item.Height
        oEditText6 = oNewItem7.Specific
        oEditText6.DataBind.SetBound(True, "ORDR", "U_Dfrozen")


        oNewItem8.Left = oEditText6.Item.Left
        oNewItem8.Width = oEditText6.Item.Width
        oNewItem8.Height = oEditText6.Item.Height
        oNewItem8.Top = oEditText6.Item.Top + oEditText6.Item.Height
        oEditText7 = oNewItem8.Specific
        oEditText7.DataBind.SetBound(True, "ORDR", "U_typepayment")



    End Sub

    Public Sub filltxtUDF(oForm As SAPbouiCOM.Form)


        oEditText = oForm.Items.Item("4").Specific
        Dim cardcode As String = oEditText.Value
        Dim query As String = "select * From OCRD where CardCode = '" & cardcode & "'"
        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery(query)

        oEditText = oForm.Items.Item("otype").Specific
        oEditText.Value = oRS.Fields.Item("U_typepayment").Value

        query = "select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode where t1.U_CardCode = '" & cardcode & "'"
        oRS.DoQuery(query)


        Dim i As Integer
        For i = 0 To oRS.RecordCount - 1
            If oRS.Fields.Item("U_TypeRoute").Value = "Frozen" Then
                oEditText = oForm.Items.Item("oFrozen").Specific
                oEditText.Value = oRS.Fields.Item("U_RouteCode").Value.ToString

                oEditText = oForm.Items.Item("oDFrozen").Specific
                oEditText.Value = oRS.Fields.Item("U_DriverCode").Value.ToString


            ElseIf oRS.Fields.Item("U_TypeRoute").Value = "Fresh" Then

                oEditText = oForm.Items.Item("oFresh").Specific
                oEditText.Value = oRS.Fields.Item("U_RouteCode").Value.ToString

                oEditText = oForm.Items.Item("oDFresh").Specific
                oEditText.Value = oRS.Fields.Item("U_DriverCode").Value.ToString


            End If


            oRS.MoveNext()
        Next





        oEditText = oForm.Items.Item("otime").Specific
        Dim thisDay As DateTime = DateTime.Today
        oEditText.Value = DateTime.Now


    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then



                                    oMatrix = oForm.Items.Item("38").Specific



                                    oEditText = oForm.Items.Item("12").Specific
                                    If oEditText.Value = "" Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Due Date Missing", SAPbouiCOM.BoMessageTime.bmt_Long)
                                        Exit Sub
                                    End If

                                
                                    Dim oForm2 As SAPbouiCOM.Form
                                    Try

                                        oForm2 = oApplication.SBO_Application.Forms.GetForm("-139", 0)
                                    Catch ex As Exception
                                        oApplication.SBO_Application.Menus.Item("6913").Activate()
                                        oForm2 = oApplication.SBO_Application.Forms.GetForm("-139", 0)

                                    End Try

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)



                                    '  oApplication.SBO_Application.Menus.Item(mnu_sales).Activate()

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ALL_EVENTS
                                BubbleEvent = False
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.Row > 0 Then
                                    Dim strValue As String = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)
                                    Dim strValue1 As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    If strValue.Length > 0 Then
                                        Dim oRecordSet As SAPbobsCOM.Recordset
                                        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim strQuery As String = "Select ItemCode From OITM Where ItemCode = '" + strValue1 + "'"
                                        oRecordSet.DoQuery(strQuery)
                                        If oRecordSet.EoF Then
                                            If oApplication.Utilities.getMatrixValues(oMatrix, "25", pVal.Row) <> "" Then
                                                If strValue1 <> "" Then
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", pVal.Row, "")
                                                End If
                                                'oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                        End If
                                    End If
                                End If
                        End Select


                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific

                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.Row > 0 Then
                                    Dim strValue As String = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)
                                    If strValue <> "" Then
                                        fillDeliveryDate(oForm)
                                        'oMatrix.Columns.Item("11").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        '  oMatrix.SetCellFocus(pVal.Row, "11")
                                    End If
                                ElseIf pVal.ItemUID = "12" Then
                                    fillDeliveryDateByDD(oForm)

                                End If
                                'If pVal.Row > 0 And pVal.ItemUID = "38" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                                ' End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                Try
                                    'here is when create a sales order customized ! 

                                    If flagitem = True Then


                                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                        createEditText(oForm)



                                        oEditText = oForm.Items.Item("4").Specific
                                        oEditText.Value = clsSalesOrder.cardcode
                                        clsSalesOrder.cardcode = ""
                                        oEditText = oForm.Items.Item("12").Specific
                                        Dim thisDay As DateTime = DateTime.Today
                                        Dim x As String = Format(thisDay, "yyyyMMdd")
                                        oEditText.Value = Format(thisDay, "yyyyMMdd")

                                        'insert the item
                                        oMatrix = oForm.Items.Item("38").Specific


                                        Dim i As Integer
                                        If clsSalesOrderSystem.flagitem = False Then

                                            clsSalesOrder.anArray = New ArrayList
                                        End If

                                        'MsgBox(Numberofdays)
                                        For i = 0 To clsSalesOrder.anArray.Count - 1
                                            Dim query As String = "select t0.itemname,t0.itemcode,t1.U_TypeCode,t1.U_Numberofdays as Numberofdays,t1.U_DeliveryDaysSales as DeliveryDaysSales from OITM t0 inner join [@Z_ItemType] t1 on t1.U_TypeCode = t0.U_Z_ItemType where t0.itemcode= '" & clsSalesOrder.anArray(i).split(",")(0) & "'"
                                            Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRS.DoQuery(query)
                                            Dim Numberofdays As Integer = oRS.Fields.Item("DeliveryDaysSales").Value
                                            Dim tmp As String = clsSalesOrder.anArray(i).ToString.Split(",")(0)

                                            ' oMatrix.SetCellWithoutValidation(i + 1, "1", "A00001")
                                            oMatrix.Columns.Item("1").Cells.Item(i + 1).Specific.value = tmp
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "11", i + 1, clsSalesOrder.anArray(i).split(",")(1))
                                            Dim tmpdata As DateTime = thisDay.AddDays(Numberofdays)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "25", i + 1, Format(tmpdata, "yyyyMMdd"))
                                        Next
                                        'oMatrix.Columns.Item("11").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)



                                        '    Dim oNewItem1 As SAPbouiCOM.Item = oForm.Items.Add("oFresh", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                        ' Dim oNewItem3 As SAPbouiCOM.Item = oForm.Items.Add("oFrozen", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                        ' Dim oNewItem4 As SAPbouiCOM.Item = oForm.Items.Add("oSalesp", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                        ' Dim oNewItem5 As SAPbouiCOM.Item = oForm.Items.Add("otime", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                        ' Dim oNewItem6 As SAPbouiCOM.Item = oForm.Items.Add("oDFresh", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                        ' Dim oNewItem7 As SAPbouiCOM.Item = oForm.Items.Add("oDFrozen", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)


                                        'fillUDF(oForm, oForm2)
                                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                        filltxtUDF(oForm)



                                        clsSalesOrder.anArray = New ArrayList
                                        flagitem = False


                                    End If

                                    clsSalesOrder.anArray = New ArrayList

                                Catch ex As Exception

                                End Try

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oForm2 As SAPbouiCOM.Form
                                Try
                                    oForm2 = oApplication.SBO_Application.Forms.GetForm("-139", 0)
                                Catch ex As Exception

                                End Try

                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                oCFLEvent = pVal
                                Dim oCFL2 As SAPbouiCOM.ChooseFromList
                                oCFL2 = oForm.ChooseFromLists.Item(oCFLEvent.ChooseFromListUID)

                                Dim oDataTable2 As SAPbouiCOM.DataTable

                                oDataTable2 = oCFLEvent.SelectedObjects
                                Try
                                    If oCFL2.ObjectType = "2" Then


                                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                        CardCode = oDataTable2.GetValue("CardCode", 0)
                                        Dim query As String = "select * From OCRD where CardCode = '" & oDataTable2.GetValue("CardCode", 0) & "'"
                                        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRS.DoQuery(query)

                                        If oRS.Fields.Item("U_freshroute").Value = "" And oRS.Fields.Item("U_frozenroute").Value = "" Then

                                            oApplication.Utilities.Message("You can't add a Sales order if the Customer is not linked to any Route (Please check the customer route master from the main menu).", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If




                                        oEditText = oForm2.Items.Item("U_typepayment").Specific
                                        oEditText.Value = oRS.Fields.Item("U_typepayment").Value

                                        query = "select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode where t1.U_CardCode = '" & oDataTable2.GetValue("CardCode", 0) & "'"
                                        oRS.DoQuery(query)


                                        Dim i As Integer
                                        For i = 0 To oRS.RecordCount - 1
                                            If oRS.Fields.Item("U_TypeRoute").Value = "Frozen" Then
                                                oEditText = oForm2.Items.Item("U_frozenroute").Specific
                                                oEditText.Value = oRS.Fields.Item("U_RouteCode").Value.ToString

                                                oEditText = oForm2.Items.Item("U_Dfrozen").Specific
                                                oEditText.Value = oRS.Fields.Item("U_DriverCode").Value.ToString


                                            ElseIf oRS.Fields.Item("U_TypeRoute").Value = "Fresh" Then

                                                oEditText = oForm2.Items.Item("U_freshroute").Specific
                                                oEditText.Value = oRS.Fields.Item("U_RouteCode").Value.ToString

                                                oEditText = oForm2.Items.Item("U_Dfresh").Specific
                                                oEditText.Value = oRS.Fields.Item("U_DriverCode").Value.ToString


                                            End If


                                            oRS.MoveNext()
                                        Next



                                    End If

                                    If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.Row > 0 Then
                                        Dim strItemCode = oDataTable2.GetValue("ItemCode", 0)
                                        If pVal.Action_Success Then
                                            fillDeliveryDateByRow(oForm, pVal.Row, strItemCode)
                                        End If



                                        'oEditText = oForm.Items.Item("12").Specific
                                        'Dim thisDay As DateTime = DateTime.Today
                                        'Dim x As String = Format(thisDay, "yyyyMMdd")
                                        'oEditText.Value = Format(thisDay, "yyyyMMdd")



                                        '  oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                        '   Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                        ' Dim oCFL As SAPbouiCOM.ChooseFromList
                                        ' Dim objEdit As SAPbouiCOM.EditTextColumn
                                        ' Dim oGr As SAPbouiCOM.Grid
                                        '  Dim oItm As SAPbobsCOM.BusinessPartners
                                        '  Dim sCHFL_ID, val, strBPCode As String
                                        '   sCHFL_ID = "CFL_0"
                                        '  oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                        '  oCFLEvento = pVal
                                        '  sCHFL_ID = oCFLEvento.ChooseFromListUID
                                        ' oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                        ' Dim oDataTable As SAPbouiCOM.DataTable
                                        ' oDataTable = oCFLEvento.SelectedObjects


                                        'insert the item
                                        '   oMatrix = oForm.Items.Item("38").Specific
                                        ' Dim i As Integer
                                        ' Dim c As Integer = pVal.Row






                                    End If

                                Catch ex As Exception

                                End Try

                                'Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                                '    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '    oMatrix = oForm.Items.Item("38").Specific
                                '    If oMatrix.VisualRowCount > 0 And flagmatrix = True And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                                '        Dim thisDay As DateTime = DateTime.Today

                                '        oForm.Freeze(True)

                                '        Dim i As Integer
                                '        For i = 1 To oMatrix.VisualRowCount
                                '            If oApplication.Utilities.getMatrixValues(oMatrix, "1", i) <> "" Then

                                '                Dim query As String = "select t0.itemname,t0.itemcode,t1.U_TypeCode,t1.U_Numberofdays as Numberofdays,t1.U_DeliveryDaysSales as DeliveryDaysSales from OITM t0 inner join [@Z_ItemType] t1 on t1.U_TypeCode = t0.U_Z_ItemType where t0.itemcode= '" & oApplication.Utilities.getMatrixValues(oMatrix, "1", i) & "'"
                                '                Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '                oRS.DoQuery(query)
                                '                Dim Numberofdays As Integer = oRS.Fields.Item("DeliveryDaysSales").Value

                                '                Dim tmpdata As DateTime = thisDay.AddDays(Numberofdays)
                                '                ' If c > oMatrix.RowCount - 1 Then
                                '                'oMatrix.AddRow()
                                '                '  End If

                                '                Try
                                '                    ' oApplication.Utilities.SetMatrixValues(oMatrix, "1", c, oDataTable.GetValue("ItemCode", i))
                                '                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", i, Format(tmpdata, "yyyyMMdd"))
                                '                    oMatrix.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '                    'CType(oMatrix.Columns.Item("38").Cells().Item(i).Specific, SAPbouiCOM.EditText).Item.Click()
                                '                Catch ex As Exception
                                '                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", i, Format(tmpdata, "yyyyMMdd"))
                                '                    oMatrix.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '                    ' CType(oMatrix.Columns.Item("38").Cells().Item(i).Specific, SAPbouiCOM.EditText).Item.Click()


                                '                End Try


                                '            End If


                                '        Next
                                '        oForm.Freeze(False)
                                '        flagmatrix = False
                                '    End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Try

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                    createEditText(oForm)

                                Catch ex As Exception

                                End Try

                                '          oForm.Items.Item("oFresh").Enabled = False
                                '        oForm.Items.Item("oFrozen").Enabled = False
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                    If pVal.ActionSuccess = True Then
                                        oApplication.SBO_Application.Menus.Item(mnu_sales).Activate()
                                        oForm.Close()
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub fillDeliveryDate(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("38").Specific
            oForm.Freeze(True)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                If oMatrix.RowCount > 0 Then
                    Dim thisDay As DateTime = DateTime.Today
                    Dim i As Integer
                    Dim oDBSource As SAPbouiCOM.DBDataSource
                    oDBSource = oForm.DataSources.DBDataSources.Item("ORDR")
                    Dim strHDeliveryDate As String = oDBSource.GetValue("DocDueDate", 0) 'oApplication.Utilities.getEdittextvalue(oForm, "12")
                    For i = 1 To oMatrix.RowCount
                        If oApplication.Utilities.getMatrixValues(oMatrix, "1", i) <> "" Then

                            Dim query As String = "select t0.itemname,t0.itemcode,t1.U_TypeCode,t1.U_Numberofdays as Numberofdays, " & _
                                " t1.U_DeliveryDaysSales as DeliveryDaysSales from OITM t0 inner join [@Z_ItemType] t1 on t1.U_TypeCode = t0.U_Z_ItemType " & _
                                 " where t0.itemcode= '" & oApplication.Utilities.getMatrixValues(oMatrix, "1", i) & "'"
                            Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery(query)
                            If Not oRS.EoF Then
                                Dim Numberofdays As Integer = oRS.Fields.Item("DeliveryDaysSales").Value
                                Dim tmpdata As DateTime = thisDay.AddDays(Numberofdays)
                                Dim strDelDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "25", i)
                                If (strDelDate = "" Or strHDeliveryDate = "" Or strHDeliveryDate = strDelDate) Then
                                    Try
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "25", i, Format(tmpdata, "yyyyMMdd"))
                                        oMatrix.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Catch ex As Exception
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "25", i, Format(tmpdata, "yyyyMMdd"))
                                        oMatrix.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End Try
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub fillDeliveryDateByRow(ByVal aForm As SAPbouiCOM.Form, ByVal iRow As Integer, Optional ByVal blnItemCode As String = "")
        Try
            oMatrix = aForm.Items.Item("38").Specific
            oForm.Freeze(True)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                If oMatrix.RowCount > 0 Then
                    Dim thisDay As DateTime = DateTime.Today
                    'For i = 1 To oMatrix.VisualRowCount
                    If oApplication.Utilities.getMatrixValues(oMatrix, "1", iRow) <> "" Or blnItemCode <> "" Then
                        Dim query As String = String.Empty
                        If blnItemCode = "" Then
                            query = "select t0.itemname,t0.itemcode,t1.U_TypeCode,t1.U_Numberofdays as Numberofdays, " & _
                            " t1.U_DeliveryDaysSales as DeliveryDaysSales from OITM t0 inner join [@Z_ItemType] t1 on t1.U_TypeCode = t0.U_Z_ItemType " & _
                             " where t0.itemcode= '" & oApplication.Utilities.getMatrixValues(oMatrix, "1", iRow) & "'"
                        Else
                            query = "select t0.itemname,t0.itemcode,t1.U_TypeCode,t1.U_Numberofdays as Numberofdays, " & _
                           " t1.U_DeliveryDaysSales as DeliveryDaysSales from OITM t0 inner join [@Z_ItemType] t1 on t1.U_TypeCode = t0.U_Z_ItemType " & _
                            " where t0.itemcode= '" & blnItemCode & "'"
                        End If

                        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery(query)
                        If Not oRS.EoF Then
                            Dim Numberofdays As Integer = oRS.Fields.Item("DeliveryDaysSales").Value
                            Dim tmpdata As DateTime = thisDay.AddDays(Numberofdays)
                            Dim strDelDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "25", iRow)
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "25", iRow, Format(tmpdata, "yyyyMMdd"))
                                oMatrix.Columns.Item("11").Cells.Item(iRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Catch ex As Exception
                                oApplication.Utilities.SetMatrixValues(oMatrix, "25", iRow, Format(tmpdata, "yyyyMMdd"))
                                oMatrix.Columns.Item("11").Cells.Item(iRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End Try
                        End If
                    End If
                    'Next
                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub fillDeliveryDateByDD(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("38").Specific
            oForm.Freeze(True)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                If oMatrix.RowCount > 0 Then
                    Dim thisDay As DateTime = DateTime.Today
                    Dim i As Integer
                    For i = 1 To oMatrix.RowCount
                        If oApplication.Utilities.getMatrixValues(oMatrix, "1", i) <> "" Then

                            Dim query As String = "select t0.itemname,t0.itemcode,t1.U_TypeCode,t1.U_Numberofdays as Numberofdays, " & _
                                " t1.U_DeliveryDaysSales as DeliveryDaysSales from OITM t0 inner join [@Z_ItemType] t1 on t1.U_TypeCode = t0.U_Z_ItemType " & _
                                 " where t0.itemcode= '" & oApplication.Utilities.getMatrixValues(oMatrix, "1", i) & "'"
                            Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery(query)
                            If Not oRS.EoF Then
                                Dim Numberofdays As Integer = oRS.Fields.Item("DeliveryDaysSales").Value
                                Dim tmpdata As DateTime = thisDay.AddDays(Numberofdays)
                                Dim strDelDate As String = oApplication.Utilities.getMatrixValues(oMatrix, "25", i)
                                Try
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", i, Format(tmpdata, "yyyyMMdd"))
                                    oMatrix.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Catch ex As Exception
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", i, Format(tmpdata, "yyyyMMdd"))
                                    oMatrix.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End Try
                            End If
                        End If
                    Next
                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub

End Class
