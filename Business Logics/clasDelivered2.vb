Imports SAPbobsCOM
Imports System.IO
Imports System.Threading

Public Class MultiTaskD

    Dim strQuery As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal aForm As SAPbouiCOM.Form, deldate As Date, ByVal strThreadID As String, ByVal strF As String)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ThreadProcess(ByVal oList As ArrayList)

        Try
            Dim strFile As String = "\Delivery_Creation" + System.DateTime.Now.ToString("yyyyMMddmmss") + ".txt"
            oApplication.Utilities.Trace_Process("Started Creating Delivery Document : " + System.DateTime.Now, strFile)
            For Each aDocNum As Object In oList
                Dim oInvoice, oORder As SAPbobsCOM.Documents
                Try
                    oInvoice = oApplication.Company.GetBusinessObject(BoObjectTypes.oInvoices)
                    oORder = oApplication.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes)


                    If oInvoice.GetByKey(aDocNum) Then

                        oORder = oApplication.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes)
                        oORder.CardCode = oInvoice.CardCode
                        oORder.DocDate = Now.Date
                        Dim query5 As String = "select U_sequencetype from OCRD where CardCode = '" & oORder.CardCode & "'"
                        Dim oRs5 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs5.DoQuery(query5)
                        Dim tmpseries As String = oRs5.Fields.Item("U_sequencetype").Value & "D" & Today.Year
                        Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
                        oRs5.DoQuery(query2)
                        Dim series As Integer = oRs5.Fields.Item("Series").Value
                        oORder.Series = series
                        oORder.CardCode = oInvoice.CardCode
                        oORder.CardName = oInvoice.CardName
                        oORder.DocDate = Now

                        'Newly Added by Madhu based on Farid Mail.
                        oORder.DocDueDate = oInvoice.UserFields.Fields.Item("U_deliverydate").Value

                        'Dim query3 As String = "select U_DeliveryDaysSales from [@Z_ITEMTYPE]  where U_TypeCode = '" & ReserveInvoice.Type & "'"
                        'Dim oRs3 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        'oRs3.DoQuery(query3)
                        'Dim daystoexpire As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value
                        ' Dim oDelivery As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value
                        ' oORder.DocDueDate = Now
                        ' oORder.DocDueDate = oInvoice.DocDueDate.AddDays(daystoexpire)

                        Dim query58 As String = "select * from OCRD where cardcode = '" & oInvoice.CardCode & "'"
                        Dim oRs58 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs58.DoQuery(query58)
                        For intLoop As Integer = 0 To oInvoice.UserFields.Fields.Count - 1
                            Try
                                oORder.UserFields.Fields.Item(intLoop).Value = oInvoice.UserFields.Fields.Item(intLoop).Value
                            Catch ex As Exception
                            End Try
                        Next
                        oORder.UserFields.Fields.Item("U_datetiming").Value = DateTime.Now.ToString

                        Dim oItem As SAPbobsCOM.Items
                        Dim intLineCount As Integer = 0

                        For intItems As Integer = 0 To oInvoice.Lines.Count - 1
                            oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.oItems)
                            oInvoice.Lines.SetCurrentLine(intItems)
                            If oInvoice.Lines.LineStatus = BoStatus.bost_Open Then
                                If intLineCount > 0 Then
                                    oORder.Lines.Add()
                                    oORder.Lines.SetCurrentLine(intLineCount)
                                End If
                                intLineCount = intLineCount + 1
                                oORder.Lines.ItemCode = oInvoice.Lines.ItemCode
                                oORder.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oInvoices
                                oORder.Lines.BaseEntry = oInvoice.DocEntry
                                oORder.Lines.BaseLine = oInvoice.Lines.LineNum
                                oORder.Lines.Quantity = oInvoice.Lines.RemainingOpenQuantity
                                oORder.Lines.UserFields.Fields.Item("U_type").Value = oInvoice.Lines.UserFields.Fields.Item("U_type").Value
                                'oItem.GetByKey(oInvoice.Lines.ItemCode)
                                'Dim query3 As String = "select U_DeliveryDaysSales from [@Z_ITEMTYPE]  where U_TypeCode = '" & oItem.UserFields.Fields.Item("U_Z_ItemType").Value.ToString.Trim & "'"
                                'Dim oRs3 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                'oRs3.DoQuery(query3)
                                'Dim daystoexpire As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value

                                'oORder.Lines.ShipDate = Now.AddDays(daystoexpire)
                                'oORder.DocDueDate = Now.AddDays(daystoexpire)


                                Dim dblBatchRequiredQty As Double = oInvoice.Lines.RemainingOpenInventoryQuantity
                                Dim OrecSet1 As SAPbobsCOM.Recordset
                                OrecSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                If oItem.GetByKey(oInvoice.Lines.ItemCode) Then
                                    If oItem.ManageBatchNumbers = BoYesNoEnum.tYES Then
                                        Dim inTbatchLine As Integer = 0

                                        Dim batchquantity As Double
                                        Dim dblAssignqty As Double = 0
                                        strQuery = "select itemcode, ExpDate as exp_date, BatchNum,Quantity, WhsCode from oibt where Quantity <> 0 " & _
                                             " and ItemCode = '" & oInvoice.Lines.ItemCode & "' And WhsCode = '" & oInvoice.Lines.WarehouseCode & "' order by exp_date "
                                        OrecSet1.DoQuery(strQuery)
                                        'Dim w As Integer
                                        For intBatch As Integer = 0 To OrecSet1.RecordCount - 1
                                            While (dblBatchRequiredQty > 0 And Not OrecSet1.EoF)
                                                batchquantity = OrecSet1.Fields.Item("Quantity").Value
                                                If batchquantity >= dblBatchRequiredQty Then
                                                    dblAssignqty = dblBatchRequiredQty
                                                Else
                                                    dblAssignqty = batchquantity
                                                End If

                                                If inTbatchLine > 0 Then
                                                    oORder.Lines.BatchNumbers.Add()
                                                End If
                                                oORder.Lines.BatchNumbers.SetCurrentLine(inTbatchLine)
                                                oORder.Lines.BatchNumbers.BatchNumber = OrecSet1.Fields.Item("BatchNum").Value
                                                oORder.Lines.BatchNumbers.Quantity = dblAssignqty
                                                inTbatchLine = inTbatchLine + 1
                                                dblBatchRequiredQty = dblBatchRequiredQty - dblAssignqty
                                                OrecSet1.MoveNext()
                                            End While
                                        Next
                                    End If
                                End If
                            End If
                        Next

                        If oORder.Add <> 0 Then
                            ' oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oApplication.Utilities.Trace_Process("Invoice No : " & oInvoice.DocNum & "-->ERROR ERRORCODE :" & oApplication.Company.GetLastErrorCode().ToString() + " ERRORDESC : " & oApplication.Company.GetLastErrorDescription().ToString(), strFile)
                        Else
                            Dim stNo As String
                            oApplication.Company.GetNewObjectCode(stNo)
                            oORder.GetByKey(CInt(stNo))
                            oApplication.Utilities.Trace_Process("Invoice No : " & oInvoice.DocNum & "-->Converted to Delivery : Document Number :" & oORder.DocNum, strFile)
                        End If

                    End If
                Catch ex As Exception
                    oApplication.Utilities.Trace_Process("Invoice No : " & oInvoice.DocNum & "-->ERROR : " & ex.Message, strFile)
                Finally
                    
                End Try
            Next
            oApplication.Utilities.Trace_Process("Completed Creating Delivery : " + System.DateTime.Now, strFile)
            Dim strPath As String = Path.GetTempPath().ToString() + strFile
            If (File.Exists(strPath)) Then
                System.Diagnostics.Process.Start(strPath)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub



End Class

Public Class clasDelivered
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
    Private todaydate As String
    Private mHash As New Hashtable
    Private batchnumber As Integer
    Private err As String
    Private flaginteger As Boolean = False
    Dim strQuery As String = String.Empty
    Private thread As System.Threading.Thread


    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Delivery2 Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Dim docnum As Integer
                                If pVal.ItemUID = "9" Then
                                    Dim aHash As Hashtable = New Hashtable
                                    oGrid = oForm.Items.Item("7").Specific

                                    Dim aListHash As New ArrayList
                                    Dim acount As Integer

                                    For intCount As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                        If oGrid.DataTable.GetValue("Select", intCount) = "Y" Then
                                            If Not aListHash.Contains(oGrid.DataTable.GetValue("DocEntry", intCount)) Then
                                                aListHash.Add(oGrid.DataTable.GetValue("DocEntry", intCount))
                                            End If
                                        End If
                                    Next

                                    Dim oMultiTask As New MultiTaskD()
                                    thread = New Thread(AddressOf oMultiTask.ThreadProcess)
                                    thread.Priority = ThreadPriority.Highest
                                    thread.IsBackground = True
                                    thread.ApartmentState = Threading.ApartmentState.STA
                                    thread.Start(aListHash)

                                    Application.DoEvents()
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                    oApplication.SBO_Application.SetStatusBarMessage("Generation of Delivery just Started : " & System.DateTime.Now, SAPbouiCOM.BoMessageTime.bmt_Long, False)

                                    Exit Sub

                                    Dim i As Integer = 1
                                    For i = 0 To oGrid.DataTable.Rows.Count - 1
                                        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                        If oGrid.DataTable.GetValue("Select", i) = "Y" Then
                                            docnum = oGrid.DataTable.GetValue("DocNum", i)
                                            Dim query As String = "select rd.ItemCode as ItemCode,rd.Dscription as Dscription,ord.DocDate as DocDate,ord.CardCode as 'Card Code', " & _
                                                " rd.OpenInvQty as Quantity ,ord.CardName as 'Card Name' , o.U_Z_ItemType as Type , rd.TaxCode as TaxCode , " & _
                                                " rd.DocEntry as DocEntry,rd.LineNum as LineNum, rd.ShipDate as ShipDate,rd.WhsCode " & _
                                                " from OINV ord inner join INV1 rd on rd.DocEntry = ord.DocEntry inner join OITM o on o.ItemCode = rd.ItemCode " & _
                                                " where ord.DocNum = " & docnum & " order by o.U_Z_ItemType"
                                            oRS.DoQuery(query)
                                            Dim j As Integer
                                            For j = 0 To oRS.RecordCount - 1
                                                If aHash.Contains(oRS.Fields.Item("Type").Value) Then
                                                    Dim anArray As ArrayList = aHash.Item(oRS.Fields.Item("Type").Value)
                                                    anArray.Add(New ReverseInvoice(oRS.Fields.Item("ItemCode").Value, oRS.Fields.Item("Dscription").Value, oRS.Fields.Item("DocDate").Value, oRS.Fields.Item("Card Code").Value, oRS.Fields.Item("Card Name").Value, oRS.Fields.Item("Quantity").Value, oRS.Fields.Item("Type").Value, oRS.Fields.Item("TaxCode").Value, oRS.Fields.Item("DocEntry").Value, oRS.Fields.Item("LineNum").Value, oRS.Fields.Item("ShipDate").Value, oRS.Fields.Item("WhsCode").Value))
                                                    aHash.Remove(oRS.Fields.Item("Type").Value)
                                                    aHash.Add(oRS.Fields.Item("Type").Value, anArray)
                                                Else
                                                    Dim anArray As ArrayList = New ArrayList
                                                    anArray.Add(New ReverseInvoice(oRS.Fields.Item("ItemCode").Value, oRS.Fields.Item("Dscription").Value, oRS.Fields.Item("DocDate").Value, oRS.Fields.Item("Card Code").Value, oRS.Fields.Item("Card Name").Value, oRS.Fields.Item("Quantity").Value, oRS.Fields.Item("Type").Value, oRS.Fields.Item("TaxCode").Value, oRS.Fields.Item("DocEntry").Value, oRS.Fields.Item("LineNum").Value, oRS.Fields.Item("ShipDate").Value, oRS.Fields.Item("WhsCode").Value))
                                                    aHash.Add(oRS.Fields.Item("Type").Value, anArray)
                                                End If
                                                oRS.MoveNext()
                                            Next

                                            ' Use For Each loop over the Hashtable.

                                            'AddOrderToDatabase(aHash, docnum, strFile)
                                            aHash = New Hashtable

                                        End If

                                    Next
                                    'oApplication.SBO_Application.MessageBox(err)
                                    'err = ""

                                    'Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
                                    'If (File.Exists(strPath)) Then
                                    '    System.Diagnostics.Process.Start(strPath)
                                    'End If
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                End If

                                If pVal.ItemUID = "Item_7" Then
                                    Dim oButton As SAPbouiCOM.Button
                                    oButton = oForm.Items.Item("Item_7").Specific
                                    Dim i As Integer
                                    If oButton.Caption = "Unselect All" Then

                                        oForm.Freeze(True)
                                        For i = 0 To oGrid.DataTable.Rows.Count - 1
                                            oGrid.DataTable.SetValue("Select", i, "N")
                                        Next
                                        oForm.Freeze(False)
                                        oButton.Caption = "Select All"
                                    ElseIf oButton.Caption = "Select All" Then


                                        oForm.Freeze(True)
                                        For i = 0 To oGrid.DataTable.Rows.Count - 1
                                            oGrid.DataTable.SetValue("Select", i, "Y")
                                        Next
                                        oForm.Freeze(False)
                                        oButton.Caption = "Unselect All"
                                    End If


                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                If pVal.ItemUID = "Item_1" Then
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
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Try
                                            oApplication.Utilities.setEdittextvalue(oForm, "Item_1", oDataTable.GetValue("CardCode", 0))
                                        Catch ex As Exception

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
                                            oForm.Items.Item("Item_1").Click()
                                            oApplication.Utilities.setEdittextvalue(oForm, "Item_2", oDataTable.GetValue("CardCode", 0))
                                        Catch ex As Exception
                                            oApplication.Utilities.setEdittextvalue(oForm, "Item_2", oDataTable.GetValue("CardCode", 0))
                                        End Try

                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                End If
                                If pVal.ItemUID = "11" Then
                                    Dim tmp As String = oApplication.Utilities.getEdittextvalue(oForm, "4")

                                    If tmp.Trim <> "" Then
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        DataBind(oForm)
                                        oGrid = oForm.Items.Item("7").Specific
                                        oGrid.Columns.Item("RowsHeader").Click(0)
                                        oGrid = oForm.Items.Item("7").Specific
                                        Dim cardcode As String = oGrid.DataTable.GetValue("DocNum", 0)
                                        oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                                        '  DataBindOnClick(oForm, cardcode)
                                    Else
                                        oApplication.Utilities.Message("From Date to Date are mandatory ! ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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


#Region "Add Sales Order"

    Private Sub CreateDelivery(ByVal aDocNum As String, ByVal strFile As String)
        Dim oInvoice, oORder As SAPbobsCOM.Documents
        Try
            oInvoice = oApplication.Company.GetBusinessObject(BoObjectTypes.oInvoices)
            oORder = oApplication.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes)
            'Dim strFile As String = "\Delivery_Creation" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt"
            If oInvoice.GetByKey(aDocNum) Then

                oORder = oApplication.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes)
                oORder.CardCode = oInvoice.CardCode
                oORder.DocDate = Now.Date
                Dim query5 As String = "select U_sequencetype from OCRD where CardCode = '" & oORder.CardCode & "'"
                Dim oRs5 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRs5.DoQuery(query5)
                Dim tmpseries As String = oRs5.Fields.Item("U_sequencetype").Value & "D" & Today.Year
                Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
                oRs5.DoQuery(query2)
                Dim series As Integer = oRs5.Fields.Item("Series").Value
                oORder.Series = series
                oORder.CardCode = oInvoice.CardCode
                oORder.CardName = oInvoice.CardName
                oORder.DocDate = Now

                'Newly Added by Madhu based on Farid Mail.
                oORder.DocDueDate = oInvoice.UserFields.Fields.Item("U_deliverydate").Value

                'Dim query3 As String = "select U_DeliveryDaysSales from [@Z_ITEMTYPE]  where U_TypeCode = '" & ReserveInvoice.Type & "'"
                'Dim oRs3 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                'oRs3.DoQuery(query3)
                'Dim daystoexpire As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value
                ' Dim oDelivery As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value
                ' oORder.DocDueDate = Now
                ' oORder.DocDueDate = oInvoice.DocDueDate.AddDays(daystoexpire)

                Dim query58 As String = "select * from OCRD where cardcode = '" & oInvoice.CardCode & "'"
                Dim oRs58 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRs58.DoQuery(query58)
                For intLoop As Integer = 0 To oInvoice.UserFields.Fields.Count - 1
                    Try
                        oORder.UserFields.Fields.Item(intLoop).Value = oInvoice.UserFields.Fields.Item(intLoop).Value
                    Catch ex As Exception
                    End Try
                Next
                oORder.UserFields.Fields.Item("U_datetiming").Value = DateTime.Now.ToString

                Dim oItem As SAPbobsCOM.Items
                Dim intLineCount As Integer = 0

                For intItems As Integer = 0 To oInvoice.Lines.Count - 1
                    oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.oItems)
                    oInvoice.Lines.SetCurrentLine(intItems)
                    If oInvoice.Lines.LineStatus = BoStatus.bost_Open Then
                        If intLineCount > 0 Then
                            oORder.Lines.Add()
                            oORder.Lines.SetCurrentLine(intLineCount)
                        End If
                        intLineCount = intLineCount + 1
                        oORder.Lines.ItemCode = oInvoice.Lines.ItemCode
                        oORder.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oInvoices
                        oORder.Lines.BaseEntry = oInvoice.DocEntry
                        oORder.Lines.BaseLine = oInvoice.Lines.LineNum
                        oORder.Lines.Quantity = oInvoice.Lines.RemainingOpenQuantity

                        'oItem.GetByKey(oInvoice.Lines.ItemCode)
                        'Dim query3 As String = "select U_DeliveryDaysSales from [@Z_ITEMTYPE]  where U_TypeCode = '" & oItem.UserFields.Fields.Item("U_Z_ItemType").Value.ToString.Trim & "'"
                        'Dim oRs3 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        'oRs3.DoQuery(query3)
                        'Dim daystoexpire As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value

                        'oORder.Lines.ShipDate = Now.AddDays(daystoexpire)
                        'oORder.DocDueDate = Now.AddDays(daystoexpire)


                        Dim dblBatchRequiredQty As Double = oInvoice.Lines.RemainingOpenInventoryQuantity
                        Dim OrecSet1 As SAPbobsCOM.Recordset
                        OrecSet1 = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        If oItem.GetByKey(oInvoice.Lines.ItemCode) Then
                            If oItem.ManageBatchNumbers = BoYesNoEnum.tYES Then
                                Dim inTbatchLine As Integer = 0

                                Dim batchquantity As Double
                                Dim dblAssignqty As Double = 0
                                strQuery = "select itemcode, ExpDate as exp_date, BatchNum,Quantity, WhsCode from oibt where Quantity <> 0 " & _
                                     " and ItemCode = '" & oInvoice.Lines.ItemCode & "' And WhsCode = '" & oInvoice.Lines.WarehouseCode & "' order by exp_date "
                                OrecSet1.DoQuery(strQuery)
                                'Dim w As Integer
                                For intBatch As Integer = 0 To OrecSet1.RecordCount - 1
                                    While (dblBatchRequiredQty > 0 And Not OrecSet1.EoF)
                                        batchquantity = OrecSet1.Fields.Item("Quantity").Value
                                        If batchquantity >= dblBatchRequiredQty Then
                                            dblAssignqty = dblBatchRequiredQty
                                        Else
                                            dblAssignqty = batchquantity
                                        End If

                                        If inTbatchLine > 0 Then
                                            oORder.Lines.BatchNumbers.Add()
                                        End If
                                        oORder.Lines.BatchNumbers.SetCurrentLine(inTbatchLine)
                                        oORder.Lines.BatchNumbers.BatchNumber = OrecSet1.Fields.Item("BatchNum").Value
                                        oORder.Lines.BatchNumbers.Quantity = dblAssignqty
                                        inTbatchLine = inTbatchLine + 1
                                        dblBatchRequiredQty = dblBatchRequiredQty - dblAssignqty
                                        OrecSet1.MoveNext()
                                    End While
                                Next
                            End If
                        End If
                    End If
                Next

                If oORder.Add <> 0 Then
                    ' oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.Trace_Process("Invoice No : " & oInvoice.DocNum & "-->ERROR ERRORCODE :" & oApplication.Company.GetLastErrorCode().ToString() + " ERRORDESC : " & oApplication.Company.GetLastErrorDescription().ToString(), strFile)
                Else
                    Dim stNo As String
                    oApplication.Company.GetNewObjectCode(stNo)
                    oORder.GetByKey(CInt(stNo))
                    oApplication.Utilities.Trace_Process("Invoice No : " & oInvoice.DocNum & "-->Converted to Delivery : Document Number :" & oORder.DocNum, strFile)
                End If

            End If
        Catch ex As Exception
            oApplication.Utilities.Trace_Process("Invoice No : " & oInvoice.DocNum & "-->ERROR : " & ex.Message, strFile)
        Finally
            If Not IsNothing(oORder) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oORder)
            If Not IsNothing(oInvoice) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice)
        End Try

    End Sub

    Private Sub AddOrderToDatabase(ByVal aHashtable As Hashtable, ByVal docnum As String, ByVal strFile As String)
        Try
            Dim oOrder As SAPbobsCOM.Documents  ' Order object

            Dim counthash As Hashtable = New Hashtable
            Dim thisDay As DateTime = DateTime.Today
            Dim hashtable As New Hashtable
            Dim oRecSet1 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

            ' getting all the batches... item, (batch,quantity)

            For Each element As DictionaryEntry In aHashtable
                oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                Dim anArray As ArrayList = element.Value  ' array of ReserveInvoice
                Dim a As Integer
                For a = 0 To anArray.Count - 1
                    Dim ReserveInvoice As ReverseInvoice = anArray(a)
                    strQuery = "select itemcode, ExpDate as exp_date, BatchNum,Quantity, WhsCode from oibt where Quantity <> 0 " & _
                        " and ItemCode = '" & ReserveInvoice.ItemCode & "' And WhsCode = '" & ReserveInvoice.WhsCode & "' order by exp_date "
                    oRecSet1.DoQuery(strQuery)

                    Dim x As Integer
                    Dim arraylistbatch As ArrayList

                    'If oRecSet1.EoF = True Then
                    '    flaginteger = True
                    'End If

                    For x = 0 To oRecSet1.RecordCount - 1
                        If counthash.Contains(oRecSet1.Fields.Item("itemcode").Value) Then
                            Dim qty As Double = oRecSet1.Fields.Item("Quantity").Value + counthash.Item(oRecSet1.Fields.Item("itemcode").Value)
                            counthash.Remove(oRecSet1.Fields.Item("itemcode").Value)
                            counthash.Add(oRecSet1.Fields.Item("itemcode").Value, qty)
                        Else
                            Dim qty As Double = oRecSet1.Fields.Item("Quantity").Value
                            counthash.Add(oRecSet1.Fields.Item("itemcode").Value, qty)
                        End If
                        If hashtable.Contains(oRecSet1.Fields.Item("itemcode").Value) Then
                            Dim tmphash As Hashtable = hashtable.Item(oRecSet1.Fields.Item("itemcode").Value)
                            If Not tmphash.Contains(oRecSet1.Fields.Item("BatchNum").Value) Then
                                tmphash.Add(oRecSet1.Fields.Item("BatchNum").Value, oRecSet1.Fields.Item("Quantity").Value)
                                hashtable.Remove(oRecSet1.Fields.Item("itemcode").Value)
                                hashtable.Add(oRecSet1.Fields.Item("itemcode").Value, tmphash)
                            End If
                        Else
                            Dim xhash As Hashtable = New Hashtable
                            xhash.Add(oRecSet1.Fields.Item("BatchNum").Value, oRecSet1.Fields.Item("Quantity").Value)
                            hashtable.Add(oRecSet1.Fields.Item("itemcode").Value, xhash)
                        End If
                        oRecSet1.MoveNext()

                    Next
                Next
            Next



            ' Init the Order object
            ' Use For Each loop over the Hashtable.
            For Each element As DictionaryEntry In aHashtable

                'oOrder.delive = SAPbobsCOM.BoYesNoEnum.tYES
                oOrder.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                oOrder.DocDate = Now

                Console.WriteLine(element.Key) 'type
                Console.WriteLine(element.Value) ' array of ReserveInvoice



                Dim anArray As ArrayList = element.Value  ' array of ReserveInvoice


                Dim j As Integer
                For j = 0 To anArray.Count - 1
                    Dim ReserveInvoice As ReverseInvoice = anArray(j)
                    If j = 0 Then
                        Dim query5 As String = "select U_sequencetype from OCRD where CardCode = '" & ReserveInvoice.CardCode & "'"
                        Dim oRs5 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs5.DoQuery(query5)
                        Dim tmpseries As String = oRs5.Fields.Item("U_sequencetype").Value & "D" & Today.Year
                        Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
                        oRs5.DoQuery(query2)
                        Dim series As Integer = oRs5.Fields.Item("Series").Value
                        oOrder.Series = series
                        oOrder.CardCode = ReserveInvoice.CardCode
                        oOrder.CardName = ReserveInvoice.CardName
                        oOrder.DocDueDate = Now
                        Dim query3 As String = "select U_DeliveryDaysSales from [@Z_ITEMTYPE]  where U_TypeCode = '" & ReserveInvoice.Type & "'"
                        Dim oRs3 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs3.DoQuery(query3)
                        Dim daystoexpire As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value
                        oOrder.DocDueDate = Now
                        oOrder.DocDueDate = ReserveInvoice.DocDueDate.AddDays(daystoexpire)
                        Dim query58 As String = "select * from OCRD where cardcode = '" & ReserveInvoice.CardCode & "'"
                        Dim oRs58 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs58.DoQuery(query58)
                        oOrder.UserFields.Fields.Item("U_TypeRoute").Value = ReserveInvoice.Type
                        oOrder.UserFields.Fields.Item("U_typepayment").Value = oRs58.Fields.Item("U_typepayment").Value
                        query58 = "select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode where t1.U_CardCode = '" & ReserveInvoice.CardCode & "'"
                        oRs58.DoQuery(query58)
                        Dim i As Integer
                        For i = 0 To oRs58.RecordCount - 1
                            If oRs58.Fields.Item("U_TypeRoute").Value = "Frozen" Then
                                oOrder.UserFields.Fields.Item("U_frozenroute").Value = oRs58.Fields.Item("U_RouteCode").Value.ToString
                                oOrder.UserFields.Fields.Item("U_Dfrozen").Value = oRs58.Fields.Item("U_DriverCode").Value.ToString
                            ElseIf oRs58.Fields.Item("U_TypeRoute").Value = "Fresh" Then
                                oOrder.UserFields.Fields.Item("U_freshroute").Value = oRs58.Fields.Item("U_RouteCode").Value.ToString
                                oOrder.UserFields.Fields.Item("U_Dfresh").Value = oRs58.Fields.Item("U_DriverCode").Value.ToString
                            End If
                            oRs58.MoveNext()
                        Next
                        thisDay = DateTime.Today
                        oOrder.UserFields.Fields.Item("U_datetiming").Value = DateTime.Now.ToString
                    End If
                    oOrder.Lines.ItemCode = ReserveInvoice.ItemCode
                    oOrder.Lines.ItemDescription = ReserveInvoice.ItemName
                    'oOrder.Lines.Quantity = ReserveInvoice.Quantity
                    'oOrder.Lines.TaxCode = ReserveInvoice.TaxCode
                    oOrder.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oInvoices
                    oOrder.Lines.BaseEntry = ReserveInvoice.DocEntry
                    oOrder.Lines.BaseLine = ReserveInvoice.BaseLine

                    ' ----------------------------------------------------------------------------------------------------------------

                    Dim oRecSet As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                    Dim oitem As String = ReserveInvoice.ItemCode
                    Dim Qtyneeded As Integer = ReserveInvoice.Quantity
                    Dim query As String = "select ManBtchNum from OITM where itemcode = '" & ReserveInvoice.ItemCode & "'"
                    oRecSet.DoQuery(query)
                    If oRecSet.Fields.Item("ManBtchNum").Value = "Y" Then
                        'hashtable (itemcode , (batch , qty)
                        Dim ahashtableforone As Hashtable = hashtable.Item(ReserveInvoice.ItemCode) 'ahashtableforone(batchnu,qty)
                        'Qtyneeded = Qtyneeded - Quantityavailableforbatches
                        Dim batchnumb As SAPbobsCOM.BatchNumbers = oOrder.Lines.BatchNumbers
                        Dim batchnum As String
                        'Qty is Required qty
                        Dim dblQty As Double = Qtyneeded

                        While (dblQty > 0)

                            Dim batchquantity As Double
                            strQuery = "select itemcode, ExpDate as exp_date, BatchNum,Quantity, WhsCode from oibt where Quantity <> 0 " & _
                                 " and ItemCode = '" & ReserveInvoice.ItemCode & "' And WhsCode = '" & ReserveInvoice.WhsCode & "' order by exp_date "
                            oRecSet1.DoQuery(strQuery)
                            Dim w As Integer

                            If oRecSet1.EoF Or dblQty > counthash.Item(ReserveInvoice.ItemCode) Then
                                ' flag to the batch number checking if qty is available.
                                GoTo err
                            End If

                            For w = 0 To oRecSet1.RecordCount - 1
                                If ahashtableforone(oRecSet1.Fields.Item("BatchNum").Value) <> 0 Then
                                    batchnum = oRecSet1.Fields.Item("BatchNum").Value
                                    batchquantity = ahashtableforone(oRecSet1.Fields.Item("BatchNum").Value)
                                    Exit For
                                End If
                                oRecSet1.MoveNext()
                            Next

                            '   If xlements.Value <> 0 Then

                            'batchnum = xlements.Key
                            '   batchquantity = xlements.Value
                            ' Exit For
                            ' End If


                            If dblQty < CDbl(batchquantity) And dblQty > 0 Then
                                batchnumb.BatchNumber = batchnum
                                batchnumb.Quantity = dblQty
                                Dim quantityofbatch As Double = ahashtableforone.Item(batchnum)
                                quantityofbatch = quantityofbatch - dblQty

                                ahashtableforone.Remove(batchnum)
                                ahashtableforone.Add(batchnum, quantityofbatch)
                                dblQty -= batchquantity

                                batchnumb.Add()
                                Exit While
                            Else
                                batchnumb.BatchNumber = batchnum
                                batchnumb.Quantity = batchquantity
                                batchnumb.BatchNumber = batchnum
                                batchnumb.Quantity = batchquantity
                                batchnumb.Add()

                                Dim quantityofbatch As Double = ahashtableforone.Item(batchnum)
                                quantityofbatch = quantityofbatch - batchquantity
                                ' MsgBox(quantityofbatch & " " & batchquantity)
                                ahashtableforone.Remove(batchnum)
                                ahashtableforone.Add(batchnum, quantityofbatch)
                                dblQty -= batchquantity
                            End If
                        End While

                        hashtable.Remove(ReserveInvoice.ItemCode)
                        hashtable.Add(ReserveInvoice.ItemCode, ahashtableforone)
                    End If
                    'res    
                    'B10000 B1 5
                    'B20000 B1 10
                    'B40000 B1 5
                    'B30000 B1 10

                    'B10000	NULL	B1	10	01
                    'B10000	NULL	B2	10	01

                    ' Add lines to the Orer object from the table
                    If j <> anArray.Count Then
                        oOrder.Lines.Add()
                    End If
                Next
                Dim lRetCode As Integer = oOrder.Add ' Try to add the orer to the database ' Try to add the orer to the database
                If lRetCode <> 0 Then
                    'oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription & "-docnum:" & docnum, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                If flaginteger = True Then
                    '                    err &= "Document Number:" & docnum & "- batch quantity in DB inferior to quantity needed for delivery" & vbCrLf
                    '                    flaginteger = False
                    '                End If
                    '                err &= "Document Number:" & docnum & "-" & oApplication.Company.GetLastErrorDescription & vbCrLf
                    oApplication.Utilities.Trace_Process("Invoice No : " + docnum + "-->ERROR ERRORCODE :" + oApplication.Company.GetLastErrorCode().ToString() + " ERRORDESC : " + oApplication.Company.GetLastErrorDescription().ToString(), strFile)
err:
                    oApplication.Utilities.Trace_Process("Invoice No : " + docnum + "-->ERRORDESC : " + " batch quantity in DB inferior to quantity needed for delivery", strFile)

                Else
                    oApplication.Utilities.Trace_Process("Invoice No : " + docnum + " -->Success", strFile)
                    'oApplication.Utilities.Message("Successfully added to the Database", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            Next
        Catch ex As Exception
            oApplication.Utilities.Trace_Process("Invoice No : " + docnum + "-->ERROR ERRORCODE :" + oApplication.Company.GetLastErrorCode().ToString() + " ERRORDESC : " + oApplication.Company.GetLastErrorDescription().ToString(), strFile)
        End Try
    End Sub

#End Region



    Private Sub DataBind(aform As SAPbouiCOM.Form)
        Dim strFrmCardCode, strToCardCode As String
        Dim strCondition As String
        Dim strFrmCustomer As String
        Dim strToCustomer As String
        Dim toDate As String



        'strFrmCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strFrmCustomer = oApplication.Utilities.getEdittextvalue(aform, "Item_1")
        strToCustomer = oApplication.Utilities.getEdittextvalue(aform, "Item_2")

        Dim tmpdate As Date = oApplication.Utilities.getEdittextvalue(aform, "4")
        strFrmCardCode = Format(tmpdate, "MM/dd/yyyy")

        Dim tmpdate1 As Date
        If oApplication.Utilities.getEdittextvalue(aform, "4_") <> "" Then
            tmpdate1 = oApplication.Utilities.getEdittextvalue(aform, "4_")
            toDate = Format(tmpdate1, "MM/dd/yyyy")
        End If
        If strFrmCardCode <> "" Then
            strCondition = " ord.U_deliverydate >= '" & strFrmCardCode & "'"
        Else
            strCondition = " 1=1 "
        End If

        If toDate <> "" Then
            strCondition &= " and ord.U_deliverydate <= '" & toDate & "'"
        End If


        If strFrmCustomer <> "" And strToCustomer <> "" Then
            strCondition &= " and ord.CardCode between '" & strFrmCustomer & "' and '" & strToCustomer & "'"
        ElseIf strFrmCardCode <> "" And strToCardCode = "" Then
            strCondition &= " and ord.Cardcode >='" & strFrmCustomer & "'"
        ElseIf strFrmCardCode = "" And strToCardCode <> "" Then
            strCondition &= " and ord.Cardcode <='" & strToCustomer & "'"
        Else
            strCondition &= " and 1=1"
        End If

        oCombobox = oForm.Items.Item("Item_5").Specific
        Dim type As String = oCombobox.Value
        oGrid = aform.Items.Item("7").Specific

        If type <> "Both" Then
            strCondition &= " and ocrd.U_typepayment = '" & type & "'"
        End If

        strCondition &= " and t1.LineStatus='O'"
        strCondition &= " and ord.InvntSttus='O'"
        strCondition &= " and (ord.DocStatus='O' Or ord.DocStatus='C') "
        strCondition &= " and ord.CANCELED='N'"

        Dim query As String = "select distinct ord.DocEntry, ord.DocNum,ord.DocDate as DocDate,ord.CardCode as 'Card Code',ord.CardName as 'Card Name',Convert(VarChar(1),'N') As 'Select' from OINV ord inner join INV1 t1 on t1.DocEntry = ord.DocEntry inner join OITM o on o.itemcode = t1.itemcode inner join OCRD ocrd on ocrd.cardcode = ord.cardcode  where " & strCondition
        'Dim query As String = "select rd.ItemCode as ItemCode,rd.Dscription as Dscription,ord.DocDueDate as DocDueDate,ord.CardCode as 'Card Code', rd.Quantity as Quantity ,ord.CardName as 'Card Name',Convert(VarChar(1),'N') As 'Select' , o.U_Z_ItemType as Type from ORDR ord inner join RDR1 rd on rd.DocEntry = ord.DocEntry inner join OITM o on o.ItemCode = rd.ItemCode where " & strCondition & " order by o.U_Z_ItemType"
        oGrid.DataTable.ExecuteQuery(query)

        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oGrid.Columns.Item("DocNum").Editable = False
        oGrid.Columns.Item("DocDate").Editable = False
        oGrid.Columns.Item("Card Code").Editable = False
        oGrid.Columns.Item("Card Name").Editable = False
        oGrid.Columns.Item("DocEntry").Editable = False

    End Sub



    '---------------
    '    Dim query As String = "select rd.ItemCode as ItemCode,rd.Dscription as Dscription,ord.DocDueDate as DocDueDate,ord.CardCode as 'Card Code', rd.Quantity as Quantity ,ord.CardName as 'Card Name',Convert(VarChar(1),'N') As 'Select' , o.U_Z_ItemType as Type from OINV ord inner join INV1 rd on rd.DocEntry = ord.DocEntry inner join OITM o on o.ItemCode = rd.ItemCode where " & strCondition & " order by o.U_Z_ItemType"

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Delivery2
                    LoadForm()
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

    Public Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Delivery) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        oForm = oApplication.Utilities.LoadForm(xml_Delivery2, frm_Delivery2)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        'select *  from RDR1 where BaseEntry = '2'
        ' oGrid = oForm.Items.Item("3").Specifi

        oForm.DataSources.UserDataSources.Add("frmCust", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("ToCust", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oEditText = oForm.Items.Item("Item_1").Specific
        oEditText.DataBind.SetBound(True, "", "frmCust")
        oEditText.ChooseFromListUID = "CFL_2"
        oEditText.ChooseFromListAlias = "CardCode"
        oEditText = oForm.Items.Item("Item_2").Specific
        oEditText.DataBind.SetBound(True, "", "ToCust")
        oEditText.ChooseFromListUID = "CFL_3"
        oEditText.ChooseFromListAlias = "CardCode"


        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFLs = oForm.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)


        oCFL = oCFLs.Item("CFL_2")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "C"
        oCFL.SetConditions(oCons)
        oCon = oCons.Add()


        oCFL = oCFLs.Item("CFL_3")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "C"
        oCFL.SetConditions(oCons)
        oCon = oCons.Add()


        oCombobox = oForm.Items.Item("Item_5").Specific
        oCombobox.Select("Both")

        ' FormatGrid(oGrid, baseEntry)
        If oForm.TypeEx = frm_sales Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            '      AddMode(oForm)
        End If
        oForm.Freeze(False)
    End Sub


End Class
