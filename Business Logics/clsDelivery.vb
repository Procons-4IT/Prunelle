Imports SAPbobsCOM
Imports System.Threading
Imports System.IO

'Public Delegate Sub oThreadCallback(lineCount As Integer)

Public Class MultiTask

    'Private callback As oThreadCallback
    ', callbackDelegate As oThreadCallback

    Dim oForm As SAPbouiCOM.Form
    Dim tmpdate As Date
    Dim strThread As String
    Dim strFile As String
    Dim strInsertStr As String = String.Empty

    Public Sub New()

    End Sub

    Public Sub New(ByVal aForm As SAPbouiCOM.Form, deldate As Date, ByVal strThreadID As String, ByVal strF As String)
        Try

            oForm = aForm
            tmpdate = deldate
            strThread = strThreadID
            strFile = strF

            'Dim oRecordHeader As SAPbobsCOM.Recordset
            'Dim oRecordDetails As SAPbobsCOM.Recordset
            'Dim oRecordSet As SAPbobsCOM.Recordset
            'Dim oTCompany As SAPbobsCOM.Company = oApplication.Company

            'Dim oInvoice As SAPbobsCOM.Documents
            'Dim strQuery As String

            'If 1 = 1 And Not IsNothing(oList) Then
            '    oHeaderDt = ds.Tables(0)
            '    If 1 = 1 Then
            '        For Each num As Object In oList
            '            Console.WriteLine("Executing Order No By Process " & strThread & " : " + num.ToString)

            '            CType(oForm.Items.Item("Item_0").Specific, SAPbouiCOM.StaticText).Caption = "Executing Order No By Process " & strThread & " : " + num.ToString
            '            Dim tmpdate As Date = oApplication.Utilities.getEdittextvalue(oForm, "4")
            '            oRecordHeader = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            strQuery = " Select Distinct T0.CardCode,T2.U_Z_ItemType,T3.U_weekly,T1.ShipDate,T4.U_BatchNumberPrefix,T3.U_typepayment " & _
            '                        " From ORDR T0 JOIN RDR1 T1 On T0.DocEntry = T1.DocEntry " & _
            '                        " JOIN OITM T2 On T1.ItemCode = T2.ItemCode " & _
            '                        " JOIN OCRD T3 On T0.CardCode = T3.CardCode " & _
            '                        " JOIN [@Z_ITEMTYPE] T4 On T4.U_TypeCode = T2.U_Z_ItemType " & _
            '                         " Where T0.DocEntry = '" + num.ToString + "'" & _
            '                        " And Convert(VarChar(80),T1.ShipDate,112) = '" + tmpdate.ToString("yyyyMMdd") + "'"
            '            oRecordHeader.DoQuery(strQuery)
            '            If Not oRecordHeader.EoF Then
            '                While Not oRecordHeader.EoF

            '                    Dim query As String = "select U_sequencetype from OCRD where CardCode = '" & oRecordHeader.Fields.Item("cardcode").Value & "'"
            '                    Dim oRs As SAPbobsCOM.Recordset = oTCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            '                    oRs.DoQuery(query)
            '                    Dim tmpseries As String = oRs.Fields.Item("U_sequencetype").Value & Today.Year

            '                    Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
            '                    oRs.DoQuery(query2)
            '                    Dim series As Integer = oRs.Fields.Item("Series").Value

            '                    oInvoice = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            '                    oInvoice.Series = series
            '                    oInvoice.ReserveInvoice = BoYesNoEnum.tYES
            '                    oInvoice.CardCode = oRecordHeader.Fields.Item("Cardcode").Value
            '                    oInvoice.DocDate = System.DateTime.Now
            '                    oInvoice.Comments = strThread + " " + System.DateTime.Now.ToString("H:mm:ss")
            '                    Dim dtDueDate As Date = System.DateTime.Now 'getDueDate(oCompany, oRecordHeader.Fields.Item("CardCode").Value, System.DateTime.Now)
            '                    oInvoice.DocDueDate = dtDueDate

            '                    =========================================================
            '                    Dim weekly As String = oRecordHeader.Fields.Item("U_weekly").Value
            '                    If weekly = "Y" Then
            '                        oInvoice.TaxDate = System.DateTime.Now
            '                        oInvoice.UserFields.Fields.Item("U_deliverydate").Value = oRecordHeader.Fields.Item("ShipDate").Value
            '                    Else
            '                        oInvoice.UserFields.Fields.Item("U_deliverydate").Value = oRecordHeader.Fields.Item("ShipDate").Value
            '                        oInvoice.TaxDate = oRecordHeader.Fields.Item("ShipDate").Value
            '                    End If

            '                    oInvoice.UserFields.Fields.Item("U_TypeRoute").Value = oRecordHeader.Fields.Item("Cardcode").Value
            '                    oInvoice.UserFields.Fields.Item("U_typepayment").Value = oRecordHeader.Fields.Item("U_typepayment").Value

            '                    If oRecordHeader.Fields.Item("U_BatchNumberPrefix").Value.ToString().ToUpper = "Frozen".ToUpper Then
            '                        oRecordSet = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                        strQuery = "Select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1  " & _
            '                            " on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode " & _
            '                           " where t1.U_CardCode = '" & oRecordHeader.Fields.Item("Cardcode").Value & "'" & _
            '                           " And T0.U_TypeRoute = '" + oRecordHeader.Fields.Item("U_Z_ItemType").Value + "'"
            '                        oRecordSet.DoQuery(strQuery)
            '                        If Not oRecordSet.EoF Then
            '                            oInvoice.UserFields.Fields.Item("U_frozenroute").Value = oRecordSet.Fields.Item("U_RouteCode").Value.ToString
            '                            oInvoice.UserFields.Fields.Item("U_Dfrozen").Value = oRecordSet.Fields.Item("U_DriverCode").Value.ToString
            '                        End If
            '                    Else
            '                        oRecordSet = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                        strQuery = "Select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1  " & _
            '                            " on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode " & _
            '                           " where t1.U_CardCode = '" & oRecordHeader.Fields.Item("Cardcode").Value & "'" & _
            '                           " And T0.U_TypeRoute = '" + oRecordHeader.Fields.Item("U_Z_ItemType").Value + "'"
            '                        oRecordSet.DoQuery(strQuery)
            '                        If Not oRecordSet.EoF Then
            '                            oInvoice.UserFields.Fields.Item("U_frozenroute").Value = oRecordSet.Fields.Item("U_RouteCode").Value.ToString
            '                            oInvoice.UserFields.Fields.Item("U_Dfrozen").Value = oRecordSet.Fields.Item("U_DriverCode").Value.ToString
            '                        End If
            '                    End If
            '                    oInvoice.UserFields.Fields.Item("U_datetiming").Value = DateTime.Now.ToString
            '                    =================================================================================================

            '                    oRecordDetails = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                    strQuery = "Select T1.ItemCode, T1.Dscription, T1.Quantity, T1.Price, T1.DiscPrcnt,T1.WhsCode,T1.DocEntry,T1.LineNum  " & _
            '                                " From RDR1 T1 " & _
            '                                " JOIN OITM T2 On T1.ItemCode = T2.ItemCode " & _
            '                                " Where T1.DocEntry = '" & num.ToString() & "' And T1.LineStatus = 'O' " & _
            '                                " And T2.U_Z_ItemType = '" + oRecordHeader.Fields.Item("U_Z_ItemType").Value + "'"

            '                    oRecordDetails.DoQuery(strQuery)
            '                    Dim intRow As Integer = 0

            '                    While Not oRecordDetails.EoF

            '                        If intRow > 0 Then
            '                            oInvoice.Lines.Add()
            '                        End If

            '                        oInvoice.Lines.SetCurrentLine(intRow)
            '                        oInvoice.Lines.ItemCode = oRecordDetails.Fields.Item("ItemCode").Value 'Drow("ItemCode").ToString().Trim()
            '                        oInvoice.Lines.ItemDescription = oRecordDetails.Fields.Item("Dscription").Value 'Drow("Dscription").ToString().Trim()
            '                        oInvoice.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oOrders
            '                        oInvoice.Lines.BaseEntry = oRecordDetails.Fields.Item("DocEntry").Value
            '                        oInvoice.Lines.BaseLine = oRecordDetails.Fields.Item("LineNum").Value
            '                        oInvoice.Lines.Quantity = oRecordDetails.Fields.Item("Quantity").Value ' Drow("Quantity")
            '                        oInvoice.Lines.UnitPrice = oRecordDetails.Fields.Item("Price").Value 'Drow("Price")
            '                        oInvoice.Lines.DiscountPercent = oRecordDetails.Fields.Item("DiscPrcnt").Value 'Drow("DiscPrcnt")
            '                        oInvoice.Lines.WarehouseCode = oRecordDetails.Fields.Item("WhsCode").Value '("WhsCode").ToString().Trim()

            '                        intRow = intRow + 1

            '                        oRecordDetails.MoveNext()

            '                    End While

            '                    Dim intStatus As Integer = oInvoice.Add()
            '                    If intStatus = 0 Then

            '                    Else
            '                        MessageBox.Show(oTCompany.GetLastErrorDescription())
            '                    End If

            '                    oRecordHeader.MoveNext()
            '                End While

            '            End If

            '        Next
            '    End If
            'End If
            'callback = callbackDelegate

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ThreadProcess(ByVal oList As ArrayList)

        'If Not (callback Is Nothing) Then
        '    callback(1)
        'End If

        Try
            Dim oRecordHeader As SAPbobsCOM.Recordset
            Dim oRecordDetails As SAPbobsCOM.Recordset
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oRecordSet_M As SAPbobsCOM.Recordset
            Dim oTCompany As SAPbobsCOM.Company = oApplication.Company

            Dim oInvoice As SAPbobsCOM.Documents
            Dim strQuery As String

            If 1 = 1 And Not IsNothing(oList) Then
                'oHeaderDt = ds.Tables(0)
                If 1 = 1 Then
                    For Each num As Object In oList
                        'Console.WriteLine("Executing Order No By Process " & strThread & " : " + num.ToString)

                        'CType(oForm.Items.Item("Item_0").Specific, SAPbouiCOM.StaticText).Caption = "Executing Order No By Process " & strThread & " : " + num.ToString
                        'Dim tmpdate As Date = oApplication.Utilities.getEdittextvalue(oForm, "4")

                        oRecordHeader = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet_M = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        strQuery = " Select Distinct T0.CardCode,T3.U_weekly,T1.ShipDate,T4.U_BatchNumberPrefix,T3.U_typepayment,T4.U_BatchNumberPrefix " & _
                                    " ,T3.U_sequencetype,T0.DocNum,T0.U_ordertype " & _
                                    " From ORDR T0 JOIN RDR1 T1 On T0.DocEntry = T1.DocEntry " & _
                                    " JOIN OITM T2 On T1.ItemCode = T2.ItemCode " & _
                                    " JOIN OCRD T3 On T0.CardCode = T3.CardCode " & _
                                    " JOIN [@Z_ITEMTYPE] T4 On T4.U_TypeCode = T2.U_Z_ItemType " & _
                                    " Where T0.DocEntry = '" + num.ToString + "'" & _
                                    " And Convert(VarChar(8),T1.ShipDate,112) = '" + tmpdate.ToString("yyyyMMdd") + "'"
                        oRecordHeader.DoQuery(strQuery)
                        If Not oRecordHeader.EoF Then
                            While Not oRecordHeader.EoF
                                Try
                                    'Dim query As String = "select U_sequencetype from OCRD where CardCode = '" & oRecordHeader.Fields.Item("cardcode").Value & "'"
                                    'Dim oRs As SAPbobsCOM.Recordset = oTCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    'oRs.DoQuery(query)
                                    'Dim tmpseries As String = oRs.Fields.Item("U_sequencetype").Value & Today.Year

                                    'Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
                                    'oRs.DoQuery(query2)
                                    'Dim series As Integer = oRs.Fields.Item("Series").Value

                                    oInvoice = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    'oInvoice.Series = series
                                    oInvoice.ReserveInvoice = BoYesNoEnum.tYES


                                    '============================================

                                    'Dim query As String = "select U_sequencetype from OCRD where CardCode = '" & oRecordHeader.Fields.Item("Cardcode").Value & "'"
                                    Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    'oRs.DoQuery(query)

                                    Dim tmpseries As String = oRecordHeader.Fields.Item("U_sequencetype").Value & Today.Year
                                    Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
                                    oRs.DoQuery(query2)
                                    Dim series As Integer = oRs.Fields.Item("Series").Value
                                    oInvoice.Series = series

                                    '===========================================
                                    oInvoice.CardCode = oRecordHeader.Fields.Item("Cardcode").Value
                                    oInvoice.DocDate = System.DateTime.Now
                                    'Dim dtDueDate As Date = GetDueDate(oRecordHeader.Fields.Item("CardCode").Value, System.DateTime.Now)
                                    'oInvoice.DocDueDate = dtDueDate
                                    '=========================================================
                                    Dim weekly As String = oRecordHeader.Fields.Item("U_weekly").Value
                                    If weekly = "Y" Then
                                        oInvoice.TaxDate = GetNextDate(Now)
                                        oInvoice.UserFields.Fields.Item("U_deliverydate").Value = oRecordHeader.Fields.Item("ShipDate").Value
                                    Else
                                        oInvoice.UserFields.Fields.Item("U_deliverydate").Value = oRecordHeader.Fields.Item("ShipDate").Value
                                        oInvoice.TaxDate = oRecordHeader.Fields.Item("ShipDate").Value
                                    End If

                                    oInvoice.UserFields.Fields.Item("U_TypeRoute").Value = oRecordHeader.Fields.Item("U_BatchNumberPrefix").Value
                                    oInvoice.UserFields.Fields.Item("U_typepayment").Value = oRecordHeader.Fields.Item("U_typepayment").Value

                                    If oRecordHeader.Fields.Item("U_BatchNumberPrefix").Value.ToString().ToUpper = "Frozen".ToUpper Then
                                        oRecordSet = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strQuery = "Select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1  " & _
                                            " on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode " & _
                                           " where t1.U_CardCode = '" & oRecordHeader.Fields.Item("Cardcode").Value & "'" & _
                                           " And T0.U_TypeRoute = 'Frozen' " & _
                                           " And T0.U_Active = 'Y' " & _
                                           " And T1.U_Active = 'Y' "
                                        oRecordSet.DoQuery(strQuery)
                                        If Not oRecordSet.EoF Then
                                            oInvoice.UserFields.Fields.Item("U_frozenroute").Value = oRecordSet.Fields.Item("U_RouteCode").Value.ToString
                                            oInvoice.UserFields.Fields.Item("U_Dfrozen").Value = oRecordSet.Fields.Item("U_DriverCode").Value.ToString
                                        End If
                                    Else
                                        oRecordSet = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strQuery = "Select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1  " & _
                                            " on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode " & _
                                           " where t1.U_CardCode = '" & oRecordHeader.Fields.Item("Cardcode").Value & "'" & _
                                           " And T0.U_TypeRoute = 'Fresh' " & _
                                            " And T0.U_Active = 'Y' " & _
                                           " And T1.U_Active = 'Y' "
                                        oRecordSet.DoQuery(strQuery)
                                        If Not oRecordSet.EoF Then
                                            oInvoice.UserFields.Fields.Item("U_freshroute").Value = oRecordSet.Fields.Item("U_RouteCode").Value.ToString
                                            oInvoice.UserFields.Fields.Item("U_Dfresh").Value = oRecordSet.Fields.Item("U_DriverCode").Value.ToString
                                        End If
                                    End If
                                    oInvoice.UserFields.Fields.Item("U_datetiming").Value = DateTime.Now.ToString
                                    oInvoice.UserFields.Fields.Item("U_ordertype").Value = oRecordHeader.Fields.Item("U_ordertype").Value
                                    '=================================================================================================

                                    oRecordDetails = oTCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strQuery = "Select T1.ItemCode, T1.Dscription, T1.Quantity, T1.Price, T1.DiscPrcnt,T1.WhsCode,T1.DocEntry,T1.LineNum,T1.TaxCode,T1.U_type  " & _
                                                " From RDR1 T1 " & _
                                                " JOIN OITM T2 On T1.ItemCode = T2.ItemCode " & _
                                                " JOIN [@Z_ITEMTYPE] T4 On T4.U_TypeCode = T2.U_Z_ItemType " & _
                                                " Where T1.DocEntry = '" & num.ToString() & "' And T1.LineStatus = 'O' " & _
                                                " And T4.U_BatchNumberPrefix = '" + oRecordHeader.Fields.Item("U_BatchNumberPrefix").Value + "'" & _
                                                " And Convert(VarChar(8),T1.ShipDate,112) = '" + tmpdate.ToString("yyyyMMdd") + "'"

                                    oRecordDetails.DoQuery(strQuery)
                                    Dim intRow As Integer = 0

                                    While Not oRecordDetails.EoF

                                        If intRow > 0 Then
                                            oInvoice.Lines.Add()
                                        End If

                                        oInvoice.Lines.SetCurrentLine(intRow)
                                        oInvoice.Lines.ItemCode = oRecordDetails.Fields.Item("ItemCode").Value 'Drow("ItemCode").ToString().Trim()
                                        'oInvoice.Lines.ItemDescription = oRecordDetails.Fields.Item("Dscription").Value 'Drow("Dscription").ToString().Trim()
                                        oInvoice.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oOrders
                                        oInvoice.Lines.BaseEntry = oRecordDetails.Fields.Item("DocEntry").Value
                                        oInvoice.Lines.BaseLine = oRecordDetails.Fields.Item("LineNum").Value
                                        oInvoice.Lines.UserFields.Fields.Item("U_type").Value = oRecordDetails.Fields.Item("U_type").Value
                                        'oInvoice.Lines.Quantity = oRecordDetails.Fields.Item("Quantity").Value ' Drow("Quantity")
                                        'oInvoice.Lines.UnitPrice = oRecordDetails.Fields.Item("Price").Value 'Drow("Price")
                                        'oInvoice.Lines.DiscountPercent = oRecordDetails.Fields.Item("DiscPrcnt").Value 'Drow("DiscPrcnt")
                                        'oInvoice.Lines.WarehouseCode = oRecordDetails.Fields.Item("WhsCode").Value '("WhsCode").ToString().Trim()
                                        'oInvoice.Lines.TaxCode = oRecordDetails.Fields.Item("TaxCode").Value

                                        intRow = intRow + 1

                                        oRecordDetails.MoveNext()

                                    End While

                                    Dim intStatus As Integer = oInvoice.Add()
                                    If intStatus = 0 Then
                                        Dim strDocNum As String
                                        oApplication.Company.GetNewObjectCode(strDocNum)
                                        oInvoice.GetByKey(CInt(strDocNum))
                                        'oApplication.Utilities.Trace_Process("Order No : " & oRecordHeader.Fields.Item("DocNum").Value.ToString & " -->Converted to AR Reserve Invoice : " & oInvoice.DocNum, strFile)

                                        Dim strMessage As String = "Order No : " & oRecordHeader.Fields.Item("DocNum").Value.ToString & " -->Converted to AR Reserve Invoice : " & oInvoice.DocNum
                                        strInsertStr = " INSERT INTO Z_RILG (OrderNo,Message,InvoiceNo) Values("
                                        strInsertStr &= "'" & oRecordHeader.Fields.Item("DocNum").Value.ToString & "'"
                                        strInsertStr &= ",'" & strMessage & "'"
                                        strInsertStr &= ",'" & oInvoice.DocNum & "'"
                                        strInsertStr &= " ) "
                                        oRecordSet_M.DoQuery(strInsertStr)
                                        'Console.WriteLine(strMessage)
                                    Else
                                        'MessageBox.Show(oTCompany.GetLastErrorDescription())
                                        'oApplication.Utilities.Trace_Process("Order No : " & oRecordHeader.Fields.Item("DocNum").Value.ToString & " -->Converted to AR Reserve Invoice  Failed . Error  : " & oApplication.Company.GetLastErrorDescription, strFile)

                                        Dim strMessage As String = "Order No : " & oRecordHeader.Fields.Item("DocNum").Value.ToString & " -->Converted to AR Reserve Invoice  Failed . Error  : " & oApplication.Company.GetLastErrorDescription

                                        strInsertStr = " INSERT INTO Z_RILG (OrderNo,Message,InvoiceNo) Values("
                                        strInsertStr &= "'" & oRecordHeader.Fields.Item("DocNum").Value.ToString & "'"
                                        strInsertStr &= ",'" & strMessage & "'"
                                        strInsertStr &= ",''"
                                        strInsertStr &= " ) "
                                        oRecordSet_M.DoQuery(strInsertStr)
                                        'Console.WriteLine(strMessage)
                                    End If

                                    oRecordHeader.MoveNext()
                                Catch ex As Exception
                                    'oApplication.Utilities.Trace_Process("Order No : " + oRecordHeader.Fields.Item("DocNum").Value.ToString + "ERRORDESC : " + ex.Message, strFile)

                                    Dim strMessage As String = "Order No : " + oRecordHeader.Fields.Item("DocNum").Value.ToString + "ERRORDESC : " + ex.Message
                                    strInsertStr = " INSERT INTO Z_RILG (OrderNo,Message,InvoiceNo) Values("
                                    strInsertStr &= "'" & oRecordHeader.Fields.Item("DocNum").Value.ToString & "'"
                                    strInsertStr &= ",'" & strMessage & "'"
                                    strInsertStr &= ",''"
                                    strInsertStr &= " ) "
                                    oRecordSet_M.DoQuery(strInsertStr)
                                    'Console.WriteLine(strMessage)
                                End Try
                            End While
                        End If
                    Next
                End If
            End If

            Dim oRecord_M As SAPbobsCOM.Recordset
            oRecord_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord_M.DoQuery("Select * From Z_RILG")
            If Not oRecord_M.EoF Then
                While Not oRecord_M.EoF
                    oApplication.Utilities.Trace_Process(oRecord_M.Fields.Item("Message").Value.ToString, strFile)
                    oRecord_M.MoveNext()
                End While
            End If
            oApplication.Utilities.Trace_Process("Completed Creating Reserve Invoice : " + System.DateTime.Now, strFile)

            Dim strPath As String = Path.GetTempPath().ToString() + strFile
            If (File.Exists(strPath)) Then
                System.Diagnostics.Process.Start(strPath)
            End If

            oApplication.SBO_Application.SetStatusBarMessage("Generation of Reserve invoices Completed:" & System.DateTime.Now, SAPbouiCOM.BoMessageTime.bmt_Long, False)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Function GetDueDate(ByVal cardCode As String, ByVal refDate As Date) As Date
        Dim _retVal As Date
        Dim vObj As SAPbobsCOM.SBObob
        Dim oRecordSet As SAPbobsCOM.Recordset
        vObj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet = vObj.GetDueDate(cardCode, refDate)
        If Not oRecordSet.EoF Then
            _retVal = oRecordSet.Fields.Item(0).Value
        End If
    End Function

    Private Function GetNextDate(dating As DateTime) As DateTime
        Dim today As Integer = CInt(dating.DayOfWeek)
        Dim delta As Integer = (8 - today)
        If today = 0 Then ' sunday
            Return dating.AddDays(1)
        End If
        Return dating.AddDays(delta)
    End Function

End Class

Public Class clsDelivery
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
    Private ReserveInvoice As ReverseInvoice
    Private thread As System.Threading.Thread
    Private aList_C As ArrayList
    Private aList_R As ArrayList
    Private intThreadRecords As Integer = 10

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Delivery Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Dim docnum As Integer
                                If pVal.ItemUID = "9" Then

                                    aList_C = New ArrayList()

                                    oGrid = oForm.Items.Item("7").Specific

                                    Dim i As Integer = 1
                                    Dim aListHash As New ArrayList
                                    Dim acount As Integer
                                    intThreadRecords = oGrid.DataTable.Rows.Count
                                    For i = 0 To oGrid.DataTable.Rows.Count - 1

                                        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)

                                        If oGrid.DataTable.GetValue("Select", i) = "Y" Then
                                            acount += 1
                                            Dim aHash As Hashtable = New Hashtable
                                            If Not aListHash.Contains(oGrid.DataTable.GetValue("DocEntry", i)) Then
                                                'MsgBox(oGrid.DataTable.GetValue("DocEntry", i))
                                                aListHash.Add(oGrid.DataTable.GetValue("DocEntry", i))
                                            End If
                                        End If

                                        If aListHash.Count = intThreadRecords Then
                                            aList_C.Add(aListHash)
                                        ElseIf oGrid.Rows.Count - 1 = i Then
                                            aList_C.Add(aListHash)
                                        End If
                                        If aListHash.Count = intThreadRecords Then
                                            aListHash = New ArrayList
                                        End If

                                    Next

                                    oApplication.SBO_Application.SetStatusBarMessage("Generating " & acount & " invoice(s). This may take up to one minute please do not turn off your system", SAPbouiCOM.BoMessageTime.bmt_Long, False)

                                    ' make the add !
                                    ' Use For Each loop over the Hashtable.
                                    'AddOrderToDatabase(aHash)
                                    'aHash = New Hashtable

                                    Dim ThreadCollections(aList_C.Count) As System.Threading.Thread
                                    Dim tmpdate As Date = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    Dim strFile As String = "\Reserve_Inv_Creation_" + System.DateTime.Now.ToString("yyyyMMddmmss") + ".txt"

                                    Dim strQuery As String = "Delete From Z_RILG"
                                    Dim oRecord_M As SAPbobsCOM.Recordset
                                    oRecord_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    oRecord_M.DoQuery(strQuery)

                                    oApplication.Utilities.Trace_Process("Started Creating Reserve Invoice : " + System.DateTime.Now, strFile)

                                    For index As Integer = 0 To aList_C.Count - 1
                                        Dim oMultiTask As New MultiTask(oForm, tmpdate, "Thread : " & index.ToString, strFile)
                                        thread = New Thread(AddressOf oMultiTask.ThreadProcess)
                                        thread.Priority = ThreadPriority.Highest
                                        thread.IsBackground = True
                                        thread.ApartmentState = Threading.ApartmentState.STA
                                        thread.Start(aList_C(index))
                                        ThreadCollections(index) = thread
                                    Next


                                    'For index As Integer = 0 To aList_C.Count - 1
                                    '    ThreadCollections(index).Join()
                                    'Next


                                    'For index As Integer = 0 To aList_C.Count - 1
                                    '    If Not ThreadCollections(index).IsAlive Then
                                    '        ThreadCollections(index).Abort()
                                    '    End If
                                    'Next

                                    'oRecord_M = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    'oRecord_M.DoQuery("Select * From Z_RILG")
                                    'If Not oRecord_M.EoF Then
                                    '    While Not oRecord_M.EoF
                                    '        oApplication.Utilities.Trace_Process(oRecord_M.Fields.Item("Message").Value.ToString, strFile)
                                    '        oRecord_M.MoveNext()
                                    '    End While
                                    'End If

                                    'Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
                                    'If (File.Exists(strPath)) Then
                                    '    System.Diagnostics.Process.Start(strPath)
                                    'End If

                                    Application.DoEvents()
                                    CType(oForm.Items.Item("Item_0").Specific, SAPbouiCOM.StaticText).Caption = "Reserve Invoice"
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                    oApplication.SBO_Application.SetStatusBarMessage("Generation of Reserve invoices just Started : " & System.DateTime.Now, SAPbouiCOM.BoMessageTime.bmt_Long, False)

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
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

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

                        End Select
                End Select


            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Public Shared Sub threadcallback(intthreadno As Integer)
        messagebox.show(String.format("independent thread {0} completed.", intthreadno.tostring))
    End Sub

    Private Function GetNextDate(dating As DateTime) As DateTime
        Dim today As Integer = CInt(dating.DayOfWeek)
        Dim delta As Integer = (8 - today)
        If today = 0 Then ' sunday
            Return dating.AddDays(1)
        End If
        Return dating.AddDays(delta)
    End Function


#Region "Add Sales Order"

    Private Sub AddOrderToDatabase(aHashtable As Hashtable)
        Dim thisLock As New Object

        SyncLock thisLock
            Dim oOrder As SAPbobsCOM.Documents  ' Order object

            Dim err As String = "You got errors in those Invoices: \n" '
            Dim flag As Boolean = False
            ' Init the Order object
            ' Use For Each loop over the Hashtable.

            For Each element As DictionaryEntry In aHashtable
                oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                'SyncLock oOrder


                oOrder.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tYES
                oOrder.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                oOrder.DocDate = Now

                Console.WriteLine(element.Key) 'type
                Console.WriteLine(element.Value) ' array of ReserveInvoice

                'Dim datefrom As String = oApplication.Utilities.getEdittextvalue(oForm, "4")

                Dim anArray As ArrayList = element.Value  ' array of ReserveInvoice
                'SyncLock anArray
                Dim j As Integer
                Dim count As Integer = 0


                For j = 0 To anArray.Count - 1
                    ReserveInvoice = anArray(j)
                    If j = 0 Then
                        Dim query As String = "select U_sequencetype from OCRD where CardCode = '" & ReserveInvoice.CardCode & "'"
                        Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs.DoQuery(query)
                        Dim tmpseries As String = oRs.Fields.Item("U_sequencetype").Value & Today.Year

                        Dim query2 As String = "select * from NNM1 where SeriesName ='" & tmpseries & "'"
                        oRs.DoQuery(query2)
                        Dim series As Integer = oRs.Fields.Item("Series").Value


                        oOrder.Series = series
                        oOrder.CardCode = ReserveInvoice.CardCode
                        oOrder.CardName = ReserveInvoice.CardName

                        Dim query3 As String = "select U_DeliveryDaysSales from [@Z_ITEMTYPE]  where U_TypeCode = '" & ReserveInvoice.Type & "'"
                        Dim oRs3 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRs3.DoQuery(query3)

                        Dim daystoexpire As Integer = oRs3.Fields.Item("U_DeliveryDaysSales").Value

                        Dim query4 As String = "select t1.ExtraDays,t1.ExtraMonth From ocrd t0 inner join OCTG t1 on t1.GroupNum = t0.Groupnum where t0.cardcode ='" & ReserveInvoice.CardCode & "'"
                        oRs3.DoQuery(query4)
                        Dim ExtraDays = oRs3.Fields.Item("ExtraDays").Value
                        Dim ExtraMonth = oRs3.Fields.Item("ExtraMonth").Value

                        Dim extratmp As Integer = ExtraMonth * 30 + ExtraDays


                        ' oOrder.DocDueDate = Now.AddDays(extratmp)
                        oOrder.DocDate = Now


                        query = "select * from OCRD where cardcode = '" & ReserveInvoice.CardCode & "'"
                        oRs.DoQuery(query)
                        Dim weekly As String = oRs.Fields.Item("U_weekly").Value
                        If weekly = "Y" Then
                            oOrder.TaxDate = GetNextDate(Now)
                            oOrder.UserFields.Fields.Item("U_deliverydate").Value = ReserveInvoice.ShipDate
                            'oOrder.UserFields.Fields.Item("U_deliverydate").Value = GetNextDate(ReserveInvoice.ShipDate)
                        Else
                            oOrder.UserFields.Fields.Item("U_deliverydate").Value = ReserveInvoice.ShipDate
                            oOrder.TaxDate = ReserveInvoice.ShipDate
                        End If
                        oOrder.UserFields.Fields.Item("U_TypeRoute").Value = ReserveInvoice.Type
                        oOrder.UserFields.Fields.Item("U_typepayment").Value = oRs.Fields.Item("U_typepayment").Value


                        query = "select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode where t1.U_CardCode = '" & ReserveInvoice.CardCode & "'"
                        oRs.DoQuery(query)
                        Dim i As Integer
                        For i = 0 To oRs.RecordCount - 1
                            If oRs.Fields.Item("U_TypeRoute").Value = "Frozen" Then

                                oOrder.UserFields.Fields.Item("U_frozenroute").Value = oRs.Fields.Item("U_RouteCode").Value.ToString

                                oOrder.UserFields.Fields.Item("U_Dfrozen").Value = oRs.Fields.Item("U_DriverCode").Value.ToString


                            ElseIf oRs.Fields.Item("U_TypeRoute").Value = "Fresh" Then

                                oOrder.UserFields.Fields.Item("U_freshroute").Value = oRs.Fields.Item("U_RouteCode").Value.ToString

                                oOrder.UserFields.Fields.Item("U_Dfresh").Value = oRs.Fields.Item("U_DriverCode").Value.ToString


                            End If

                            oRs.MoveNext()
                        Next

                        Dim thisDay As DateTime = DateTime.Today
                        oOrder.UserFields.Fields.Item("U_datetiming").Value = DateTime.Now.ToString

                    End If


                    If count > 0 Then
                        oOrder.Lines.Add()
                        oOrder.Lines.SetCurrentLine(count)
                    End If

                    oOrder.Lines.ItemCode = ReserveInvoice.ItemCode
                    oOrder.Lines.ItemDescription = ReserveInvoice.ItemName
                    oOrder.Lines.Quantity = ReserveInvoice.Quantity
                    oOrder.Lines.TaxCode = ReserveInvoice.TaxCode


                    oOrder.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oOrders
                    oOrder.Lines.BaseEntry = ReserveInvoice.DocEntry
                    oOrder.Lines.BaseLine = ReserveInvoice.BaseLine
                    count = count + 1

                    ' Add lines to the Orer object from the table

                Next
                'End SyncLock



                Dim lRetCode As Integer = oOrder.Add ' Try to add the orer to the database ' Try to add the orer to the database


                If lRetCode <> 0 Then
                    flag = True
                    err &= "\n -" & ReserveInvoice.CardName & " " & ReserveInvoice.DocEntry & " Error : " & oApplication.Company.GetLastErrorDescription
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    'oApplication.Utilities.Message("Delivery created successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
                'End SyncLock
            Next

            ' to change to make the mapping
            '  Dim query1 As String = "update ORDR set DocStatus = 'C' where DocEntry = " & docnum
            ' Dim oRs1 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
            ' oRs1.DoQuery(query1)
            If flag = True Then
                oApplication.SBO_Application.MessageBox(err, , "Ok", "Cancel")
                flag = False
                oApplication.Utilities.Message("Reserve invoice got error", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.Utilities.Message("Reserve invoice created successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        End SyncLock


    End Sub

#End Region

    Private Sub DataBind(aform As SAPbouiCOM.Form)
        Dim strFrmCardCode, strToCardCode As String
        Dim strCondition As String
        Dim strFrmCustomer As String
        Dim strToCustomer As String

        ' strFrmCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strFrmCustomer = oApplication.Utilities.getEdittextvalue(aform, "Item_1")
        strToCustomer = oApplication.Utilities.getEdittextvalue(aform, "Item_2")
        Dim tmpdate As Date = oApplication.Utilities.getEdittextvalue(aform, "4")
        strFrmCardCode = Format(tmpdate, "MM/dd/yyyy")

        If strFrmCardCode <> "" Then

            strCondition = " t1.ShipDate = '" & strFrmCardCode & "'"
        Else
            strCondition = " 1=1"
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

        Dim query As String = "select distinct ord.DocNum,ord.DocEntry,ord.DocDate as DocDate,ord.CardCode as 'Card Code',ord.CardName as 'Card Name',Convert(VarChar(1),'N') As 'Select' from ORDR ord inner join RDR1 t1 on t1.DocEntry = ord.DocEntry inner join OCRD ocrd on ocrd.cardcode = ord.cardcode  where " & strCondition
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


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Delivery
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

        oForm = oApplication.Utilities.LoadForm(xml_Delivery, frm_Delivery)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        'select *  from RDR1 where BaseEntry = '2'
        ' oGrid = oForm.Items.Item("3").Specific
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
