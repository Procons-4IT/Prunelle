Imports SAPbobsCOM

Public Class clsSalesOrder
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
    Private oItem As SAPbouiCOM.Item
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private oNewItem As SAPbouiCOM.Item
    Private baseEntry As String
    Public Shared anArray As ArrayList
    Public Shared cardcode As String
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Private Sub DataBind(aform As SAPbouiCOM.Form)
        Dim strFrmCardCode, strToCardCode As String
        strFrmCardCode = oApplication.Utilities.getEdittextvalue(aform, "5")
        strToCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        oGrid = aform.Items.Item("7").Specific
        Dim query As String
        If strFrmCardCode <> "" Then
            query = "Select CardCode as 'Customer Code',CardName as 'Customer Name' from OCRD where CardCode ='" & strFrmCardCode & "' and CardType= 'C'"
        Else
            query = "Select CardCode as 'Customer Code',CardName as 'Customer Name' from OCRD where CardType= 'C'"
        End If
        oGrid.DataTable.ExecuteQuery(query)


        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oGrid.Columns.Item("Customer Code").Editable = False
        oGrid.Columns.Item("Customer Name").Editable = False

    End Sub
    ' Private Sub setPrice(oForm As SAPbouiCOM.Form, cardcode As String)

    ' oGrid = oForm.Items.Item("8").Specific

    '  Dim i As Integer
    '   For i = 0 To oGrid.DataTable.Rows.Count - 1
    '  Dim itemcode As String = oGrid.DataTable.GetValue("ItemCode", i)
    'Dim query As String = "  select oc.cardcode , it.itemcode , it.PriceList,Price from OCRD oc inner join ITM1 it on it.PriceList = oc.ListNum where it.Itemcode ='" & itemcode & "' and oc.cardcode= '" & cardcode & "'"
    ' Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    ' oRs.DoQuery(query)
    '' Dim price As Double = oRs.Fields.Item("Price").Value
    '  oGrid.DataTable.SetValue("Unit Price", i, price)


    '   Next


    ' End Sub

    Private Sub DataBindOnClick(aform As SAPbouiCOM.Form, ByVal cardcode As String)
        Dim strFrmCardCode, strToCardCode As String
        Dim strCondition As String
        strFrmCardCode = oApplication.Utilities.getEdittextvalue(aform, "5")
        strToCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")

        oGrid = aform.Items.Item("8").Specific
        Dim query As String = "select distinct D1.itemcode as ItemCode,(Select ItemName from OITM where itemcode = D1.itemcode) as 'Description', Convert(Decimal,0) As 'RequesteQty', (select  Sum(Quantity) from INV1 where itemcode = D1.itemcode and basecard = '" & cardcode & "'  and ShipDate = max(D1.ShipDate)) as 'LastDeliveredQuantity' ,max(D1.ShipDate) as Late, sum(D1.Quantity) as 'QtyDeliveredinLast',  (select max(T0.docdate) from ORDR T0 JOIN RDR1 T1 On T0.DocEntry = T1.DocEntry where CardCode =  '" & cardcode & "' And T1.ItemCode = D1.ItemCode ) as 'Last Sales Order Date'  from INV1 D1 where D1.docdate between dateadd(day, -30, getdate()) and GETDATE() and D1.BaseCard = '" & cardcode & "' group by D1.itemcode Order by max(D1.ShipDate) Desc "
        oGrid.DataTable.ExecuteQuery(query)
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oGrid.Columns.Item("ItemCode").Editable = True
        'oGrid.Columns.Item("Dscription").Editable = False
        oGrid.Columns.Item("QtyDeliveredinLast").TitleObject.Caption = "Qty Delivered in Last 30 Days"
        oGrid.Columns.Item("Late").Editable = False
        oGrid.Columns.Item("Late").TitleObject.Caption = "Last Delivery Date"
        oGrid.Columns.Item("LastDeliveredQuantity").Editable = False
        oGrid.Columns.Item("LastDeliveredQuantity").TitleObject.Caption = "Last Delivered Quantity"
        oGrid.Columns.Item("QtyDeliveredinLast").Editable = False
        oGrid.Columns.Item("Last Sales Order Date").Editable = False


    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try


            If pVal.FormTypeEx = frm_sales Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" And pVal.ColUID = "RequesteQty" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("8").Specific
                                    If pVal.CharPressed = 40 Then 'Down
                                        Dim iRow As Integer = pVal.Row
                                        If iRow + 1 < oGrid.Rows.Count Then
                                            oGrid.SetCellFocus(iRow + 1, 2)

                                        End If
                                    ElseIf pVal.CharPressed = 38 Then 'Up
                                        Dim iRow As Integer = pVal.Row
                                        If iRow - 1 > -1 Then
                                            oGrid.SetCellFocus(iRow - 1, 2)
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "9" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)


                                    oGrid = oForm.Items.Item("7").Specific
                                    cardcode = oGrid.DataTable.GetValue("Customer Code", 0)
                                    Dim bool As Boolean = validateQuantity(cardcode, oForm)
                                    If bool = False Then
                                        Dim iReturnValue As Integer

                                        iReturnValue = oApplication.SBO_Application.MessageBox("The minimun order for this customer is under limit, oo you want to continue?", 3, "&Yes", "&No")

                                        Select Case iReturnValue
                                            Case 1

                                                oGrid = oForm.Items.Item("8").Specific
                                                Dim z As Integer
                                                Dim flagproceed As Boolean = False
                                                For z = 0 To oGrid.Rows.Count - 1
                                                    If oGrid.DataTable.GetValue("RequesteQty", z) <> 0 Then
                                                        flagproceed = True
                                                    End If


                                                Next

                                                If flagproceed = False Then
                                                    BubbleEvent = False
                                                    oApplication.SBO_Application.SetStatusBarMessage("One of the items should contain a quantity different than zero", SAPbouiCOM.BoMessageTime.bmt_Long)
                                                    Exit Sub
                                                End If


                                                'make the loop and reproduce it for above

                                                anArray = New ArrayList
                                                anArray = AddOrderToDatabase(oForm)
                                                clsSalesOrderSystem.flagitem = True


                                                oApplication.SBO_Application.Menus.Item(mnu_salesorder).Activate()
                                                oForm.Items.Item("2").Click()
                                                '   oForm = oApplication.SBO_Application.Forms.ActiveForm()





                                            Case 2
                                                'abort
                                                BubbleEvent = False
                                                oApplication.Utilities.Message("Your Sales order wasn't confirmed", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        End Select
                                    ElseIf bool = True Then

                                        Dim z As Integer
                                        oGrid = oForm.Items.Item("8").Specific

                                        ' Dim flagproceed As Boolean = False

                                        For z = 0 To oGrid.Rows.Count - 1

                                            If oGrid.DataTable.GetValue("RequesteQty", z) <> 0 Then
                                                'flagproceed = True
                                                ''checking if this item is linked to a type if not error
                                                'Dim query As String = "select * from OITM where ItemCode = '" & oGrid.DataTable.GetValue("ItemCode", z) & "'"
                                                'Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                'oRs.DoQuery(query)

                                                'If oRs.Fields.Item("U_Z_ItemType").Value = "" Then
                                                '    oApplication.SBO_Application.SetStatusBarMessage("The item " & oGrid.DataTable.GetValue("ItemCode", z) & " is not linked  to any type ! Please go to Item Master Data and modify it ")
                                                '    BubbleEvent = False
                                                '    Exit Sub
                                                'End If
                                            End If

                                        Next

                                        ' If flagproceed = False Then
                                        'BubbleEvent = False
                                        'oApplication.SBO_Application.SetStatusBarMessage("One of the items should contain a quantity different than zero", SAPbouiCOM.BoMessageTime.bmt_Long)
                                        ' Exit Sub
                                        ' End If

                                        'make the loop and reproduce it for above

                                        anArray = New ArrayList
                                        anArray = AddOrderToDatabase(oForm)
                                        clsSalesOrderSystem.flagitem = True


                                        oApplication.SBO_Application.Menus.Item(mnu_salesorder).Activate()
                                        oForm.Items.Item("2").Click()
                                        clsSalesOrderSystem.flagmatrix = True
                                        '   oForm = oApplication.SBO_Application.Forms.ActiveForm()

                                        ' oForm.Items.Item("12").Click()


                                    End If
                                End If



                        End Select

                    Case False
                        Select Case pVal.EventType


                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                '  Dim oForm2 As SAPbouiCOM.Form
                                Try
                                    ' oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    'oEditText = oForm.Items.Item("5").Specific
                                    'oEditText.Item.Click()
                                    ' oForm2 = oApplication.SBO_Application.Forms.GetForm("0", 0)
                                    ' oForm2.Items.Item("U_freshroute").Editable = False



                                Catch ex As Exception

                                End Try




                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                End If
                                If pVal.ItemUID = "11" Then
                                    oEditText = oForm.Items.Item("5").Specific
                                    Dim cardcode As String = oEditText.Value
                                    Dim query As String = "select * From OCRD where CardCode = '" & cardcode & "'"
                                    Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRS.DoQuery(query)

                                    If oRS.Fields.Item("U_freshroute").Value = "" And oRS.Fields.Item("U_frozenroute").Value = "" Then
                                        'here to change it 
                                        'manage error cast

                                        oApplication.Utilities.Message("You can't add a Sales order if the Customer is not linked to any Route (kindly check the customer route master from the main menu).", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    Else
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        DataBind(oForm)
                                        oGrid = oForm.Items.Item("7").Specific
                                        oGrid.Columns.Item("RowsHeader").Click(0)
                                        oGrid = oForm.Items.Item("7").Specific
                                        cardcode = oGrid.DataTable.GetValue("Customer Code", 0)
                                        DataBindOnClick(oForm, cardcode)
                                        'setPrice(oForm, cardcode)
                                    End If

                                End If


                                If pVal.ItemUID = "btnadd" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oGrid = oForm.Items.Item("8").Specific
                                    oGrid.DataTable.Rows.Add()

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                If pVal.ItemUID = "4" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    Dim objEdit As SAPbouiCOM.EditTextColumn
                                    Dim oGr As SAPbouiCOM.Grid
                                    Dim oItm As SAPbobsCOM.BusinessPartners
                                    Dim sCHFL_ID, val, strBPCode As String
                                    sCHFL_ID = "CFL_0"
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects

                                    If (oCFLEvento.BeforeAction = False) Then
                                        Try

                                            oApplication.Utilities.setEdittextvalue(oForm, "4", oDataTable.GetValue("CardName", 0))
                                        Catch ex As Exception
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "5", oDataTable.GetValue("CardCode", 0))
                                            Catch ex1 As Exception
                                            End Try
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "Item_2", oDataTable.GetValue("CardFName", 0))
                                            Catch ex1 As Exception
                                            End Try
                                        End Try
                                    End If
                                ElseIf pVal.ItemUID = "5" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    Dim objEdit As SAPbouiCOM.EditTextColumn
                                    Dim oGr As SAPbouiCOM.Grid
                                    Dim oItm As SAPbobsCOM.BusinessPartners
                                    Dim sCHFL_ID, val, strBPCode As String
                                    sCHFL_ID = "CFL_1"

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)





                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If (oCFLEvento.BeforeAction = False) Then

                                        Try

                                            oApplication.Utilities.setEdittextvalue(oForm, "4", oDataTable.GetValue("CardName", 0))
                                        Catch ex As Exception
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "5", oDataTable.GetValue("CardCode", 0))
                                            Catch ex1 As Exception
                                            End Try
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "Item_2", oDataTable.GetValue("CardFName", 0))
                                            Catch ex1 As Exception
                                            End Try
                                        End Try

                                    End If
                                ElseIf pVal.ItemUID = "Item_2" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    Dim objEdit As SAPbouiCOM.EditTextColumn
                                    Dim oGr As SAPbouiCOM.Grid
                                    Dim oItm As SAPbobsCOM.BusinessPartners
                                    Dim sCHFL_ID, val, strBPCode As String
                                    sCHFL_ID = "CFL_3"

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)





                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects


                                    If (oCFLEvento.BeforeAction = False) Then

                                        Try

                                            oApplication.Utilities.setEdittextvalue(oForm, "4", oDataTable.GetValue("CardName", 0))
                                        Catch ex As Exception
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "5", oDataTable.GetValue("CardCode", 0))
                                            Catch ex1 As Exception
                                            End Try
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "Item_2", oDataTable.GetValue("CardFName", 0))
                                            Catch ex1 As Exception
                                            End Try
                                        End Try

                                    End If
                                ElseIf pVal.ItemUID = "8" Then

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    Dim objEdit As SAPbouiCOM.EditTextColumn
                                    Dim oGr As SAPbouiCOM.Grid
                                    Dim oItm As SAPbobsCOM.BusinessPartners
                                    Dim sCHFL_ID, val, strBPCode As String
                                    sCHFL_ID = "CFL_2"

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)



                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects



                                    Dim index As Integer
                                    Dim count As Integer = 0
                                    count = pVal.Row
                                    Try


                                        For index = 0 To oDataTable.Rows.Count - 1




                                            If (oCFLEvento.BeforeAction = False) Then
                                                Try
                                                    'checking if this item is linked to a type if not error
                                                    Dim query As String = "select * from OITM where ItemCode = '" & oDataTable.GetValue("ItemCode", index) & "'"
                                                    Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRs.DoQuery(query)

                                                    If oRs.Fields.Item("U_Z_ItemType").Value = "" Then
                                                        oApplication.SBO_Application.SetStatusBarMessage("The item " & oDataTable.GetValue("ItemCode", index) & " is not linked  to any type ! ")
                                                        BubbleEvent = False
                                                    End If
                                                    oGrid.DataTable.SetValue("ItemCode", count, oDataTable.GetValue("ItemCode", index))
                                                    oGrid.DataTable.SetValue("Description", count, oDataTable.GetValue("ItemName", index))

                                                Catch ex As Exception
                                                    oGrid.DataTable.SetValue("Description", count, oDataTable.GetValue("ItemName", index))
                                                End Try

                                            End If
                                            oGrid.DataTable.Rows.Add()
                                            count += 1
                                        Next
                                    Catch ex As Exception

                                    End Try
                                End If


                        End Select

                End Select
            ElseIf pVal.FormTypeEx = frm_SalesOrder Then

                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim i As Integer
                                    For i = 1 To oMatrix.VisualRowCount
                                        Dim TaxCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "18", i)
                                        Dim itemcode As String = oApplication.Utilities.getMatrixValues(oMatrix, "38", i)

                                        If TaxCode = "" Then
                                            CType(oMatrix.Columns.Item("160").Cells().Item(i).Specific, SAPbouiCOM.EditText).Item.Click()
                                            oApplication.Utilities.Message("Tax Code should not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                        End If



                                    Next


                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType


                        End Select
                End Select

            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Add Sales Order"

    Private Function validateQuantity(cardcode As String, oform As SAPbouiCOM.Form) As Boolean
        Dim query As String = "select * from OCRD where cardcode ='" & cardcode & "'"
        Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)


        oGrid = oform.Items.Item("8").Specific
        Dim oRs1 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim j As Integer = 0
        Dim aHash As Hashtable = New Hashtable

        For j = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("RequesteQty", j) <> 0 Then


                Dim itemcode As String = oGrid.DataTable.GetValue("ItemCode", j)
                query = "select * from OITM where ItemCode ='" & itemcode & "'"
                oRs1.DoQuery(query)
                Dim type As String = oRs1.Fields.Item("U_Z_ItemType").Value


                If aHash.Contains(type) Then
                    Dim Qty As Integer = oGrid.DataTable.GetValue("RequesteQty", j)
                    Dim number As Integer = aHash.Item(type)
                    number = number + Qty
                    aHash.Remove(type)
                    aHash.Add(type, number)
                Else
                    aHash.Add(type, oGrid.DataTable.GetValue("RequesteQty", j))
                End If
            End If

        Next


        Dim minfreshString As String = oRs.Fields.Item("U_minfresh").Value
        Dim minfrozenString As String = oRs.Fields.Item("U_minfrozen").Value

        If IsNumeric(minfreshString) Then
            Dim minFresh As Double = oRs.Fields.Item("U_minfresh").Value
            If minFresh <> 0 Then
                For Each element As DictionaryEntry In aHash
                    If element.Key.ToString().ToLower = "fresh" Then
                        If element.Value < minFresh Then
                            Return False
                        End If
                    End If
                Next
            End If
        End If


        If IsNumeric(minfrozenString) Then
            Dim minFrozen As Double = oRs.Fields.Item("U_minfrozen").Value
            If minFrozen <> 0 Then
                For Each element As DictionaryEntry In aHash
                    If element.Key.ToString().ToLower = "frozen" Then
                        If element.Value < minFrozen Then
                            Return False
                        End If
                    End If
                Next
            End If
        End If
        Return True



    End Function
    Private Function AddOrderToDatabase(oForm As SAPbouiCOM.Form) As ArrayList
        Dim oOrder As SAPbobsCOM.Documents ' Order object
        ' Init the Order object
        oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        oGrid = oForm.Items.Item("7").Specific

        ' set properties of the Order object
        oOrder.CardCode = oGrid.DataTable.GetValue("Customer Code", 0)
        oOrder.CardName = oGrid.DataTable.GetValue("Customer Name", 0)

        ' Add lines to the Orer object from the table


        Dim an As ArrayList = New ArrayList
        Dim oLinesGrid As SAPbouiCOM.Grid = oForm.Items.Item("8").Specific
        Dim j As Integer = 1
        For j = 0 To oLinesGrid.Rows.Count - 1

            If oLinesGrid.DataTable.GetValue("RequesteQty", j) <> 0.0 Then
                oOrder.Lines.ItemCode = oLinesGrid.DataTable.GetValue("ItemCode", j)
                an.Add(oLinesGrid.DataTable.GetValue("ItemCode", j) & "," & oLinesGrid.DataTable.GetValue("RequesteQty", j))
            End If

        Next
        Return an

        '   oOrder.DocNum = txtNo.Text
        'oOrder.DocDate = DatePosting.Value
        ' oOrder.DocDueDate = DateDelivery.Value
        ' oOrder.DocCurrency = cmbCurrency.Items(cmbCurrency.SelectedIndex)

        ' TableLines.AcceptChanges() ' Update the lines table

        'Dim thisDay As DateTime = DateTime.Today
        ' Display the date in the default (general) format.


    End Function



#End Region


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID

                Case mnu_sales
                    If pVal.BeforeAction = False Then
                        LoadForm()
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub AddchooseFromList(aForm As SAPbouiCOM.Form)

        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim objEdit As SAPbouiCOM.EditTextColumn
        Dim oGr As SAPbouiCOM.Grid
        Dim oItm As SAPbobsCOM.BusinessPartners
        Dim sCHFL_ID, val, strBPCode As String
        sCHFL_ID = "CFL_0"
        oForm = aForm
        'sCHFL_ID = oCFLEvento.ChooseFromListUID
        oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)

        Dim oCon As SAPbouiCOM.Conditions
        oCon = oCFL.GetConditions()
        Dim oCon1 As SAPbouiCOM.Condition

        oCon1 = oCon.Add()
        oCon1.Alias = "CardType"
        oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon1.CondVal = "C"
        oCFL.SetConditions(oCon)

        sCHFL_ID = "CFL_1"
        oForm = aForm
        'sCHFL_ID = oCFLEvento.ChooseFromListUID
        oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)


        oCon = oCFL.GetConditions()
        oCon1 = oCon.Add()
        oCon1.Alias = "CardType"
        oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon1.CondVal = "C"
        oCFL.SetConditions(oCon)
    End Sub
    Public Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_RouteMaster) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        oForm = oApplication.Utilities.LoadForm(xml_sales, frm_sales)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        'select *  from RDR1 where BaseEntry = '2'
        ' oGrid = oForm.Items.Item("3").Specific

        oForm.DataSources.UserDataSources.Add("frmCust", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("ToCust", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)


        ' FormatGrid(oGrid, baseEntry)
        If oForm.TypeEx = frm_sales Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            '      AddMode(oForm)
        End If
        Try
            AddchooseFromList(oForm)
        Catch ex As Exception

        End Try
        oForm.Freeze(False)
    End Sub

#Region "Format the Grid "
    Private Sub FormatGrid(ByVal oGrid As SAPbouiCOM.Grid, ByVal baseEntry As String)


        oGrid.DataTable.ExecuteQuery("select Quantity  from [RDR1] where BaseEntry = '" & baseEntry & "'")
        'check here
        Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
        oGrid.Columns.Item("U_RoutesCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        oEditTextColumn = oGrid.Columns.Item("U_RoutesCode")

        oGrid.Columns.Item("U_RoutesCode").TitleObject.Caption = "Route Code"
        oGrid.Columns.Item("U_RoutesName").TitleObject.Caption = "Route Name"
        oGrid.Columns.Item("U_DriverCode").TitleObject.Caption = "Driver Name"
        oGrid.Columns.Item("U_DriverCode").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox


        Dim oCombo As SAPbouiCOM.ComboBoxColumn = oGrid.Columns.Item("U_DriverCode")
        Dim oComboColumn As SAPbouiCOM.ComboBoxColumn
        oComboColumn = oGrid.Columns.Item("U_DriverCode")


        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim query As String = "Select * from [@Drivers] where U_chkActive = 'Y'"
        oRS.DoQuery(query)
        Dim i As Integer
        For i = 0 To oRS.RecordCount - 1
            Dim DriverCode As String = oRS.Fields.Item("U_DriverCode").Value
            Dim DriverName As String = oRS.Fields.Item("U_DriverName").Value
            oCombo.ValidValues.Add(DriverCode, DriverName)
            oRS.MoveNext()

        Next

        oGrid.Columns.Item("U_DriverCode").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox

        oGrid.Columns.Item("U_TypeRoute").TitleObject.Caption = "Route Type"
        oGrid.Columns.Item("U_TypeRoute").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox


        oComboColumn = oGrid.Columns.Item("U_TypeRoute")
        oComboColumn.ValidValues.Add("Fresh", "Fresh")
        oComboColumn.ValidValues.Add("Frozen", "Frozen")



        oGrid.Columns.Item("U_chkActive").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oGrid.Columns.Item("U_chkActive").TitleObject.Caption = "Active"

        oComboColumn = oGrid.Columns.Item("U_chkActive")
        oComboColumn.ValidValues.Add("N", "N")
        oComboColumn.ValidValues.Add("Y", "Y")

        oGrid.Columns.Item("DocEntry").Visible = False
        oGrid.Columns.Item("DocNum").Visible = False

        oGrid.Columns.Item("Period").Visible = False
        oGrid.Columns.Item("Instance").Visible = False

        oGrid.Columns.Item("Series").Visible = False
        oGrid.Columns.Item("Handwrtten").Visible = False

        oGrid.Columns.Item("LogInst").Visible = False
        oGrid.Columns.Item("UserSign").Visible = False
        'T1.[LogInst], T1.[UserSign], T1.[Transfered], T1.[Status], T1.[CreateDate], T1.[CreateTime], T1.[], T1.[UpdateDate], T1.[DataSource]
        oGrid.Columns.Item("Transfered").Visible = False
        oGrid.Columns.Item("Status").Visible = False

        oGrid.Columns.Item("CreateDate").Visible = False
        oGrid.Columns.Item("UpdateDate").Visible = False
        oGrid.Columns.Item("Canceled").Visible = False
        oGrid.Columns.Item("Object").Visible = False
        oGrid.Columns.Item("CreateTime").Visible = False
        oGrid.Columns.Item("UpdateTime").Visible = False
        oGrid.Columns.Item("DataSource").Visible = False

        oGrid.AutoResizeColumns()
        For intLoop As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intLoop, intLoop + 1)
        Next
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single




        oForm.Update()
        oForm.Refresh()
    End Sub
#End Region

End Class
