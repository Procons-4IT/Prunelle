Imports SAPbobsCOM

Public Class batchessetup
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
    Private batchnumber As Integer
    Private todaydate As String
    Public Shared flagFIFO As Boolean = True

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_batch_setup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)




                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)


                                'add the button 
                                Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("36").Specific


                                Dim oNewItem As SAPbouiCOM.Item = oForm.Items.Add("createb", SAPbouiCOM.BoFormItemTypes.it_BUTTON)


                                oNewItem.Left = obutton.Item.Left - obutton.Item.Width
                                oNewItem.Width = obutton.Item.Width
                                oNewItem.Height = obutton.Item.Height
                                oNewItem.Top = obutton.Item.Top

                                Dim obutton1 As SAPbouiCOM.Button
                                obutton1 = oNewItem.Specific
                                obutton1.Caption = "Create Batches"
                                obutton1.Item.Visible = False




                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.Columns.Item("7").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                If flagBatchsetup = False Then
                                    obutton1.Item.Visible = True

                                    oMatrix.Columns.Item("2").Editable = False
                                    oMatrix.Columns.Item("5").Editable = False
                                    oMatrix.Columns.Item("10").Editable = False

                                Else


                                    oMatrix.Columns.Item("2").Editable = True
                                    oMatrix.Columns.Item("5").Editable = True
                                    oMatrix.Columns.Item("10").Editable = True


                                End If

                                flagBatchsetup = True

                            Case SAPbouiCOM.BoEventTypes.et_ALL_EVENTS

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                flagBatchsetup = True

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


                                If pVal.ItemUID = "createb" Then
                                    Dim mHash As New Hashtable

                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim thisDay As DateTime = DateTime.Today
                                    ' Display the date in the default (general) format.
                                    todaydate = thisDay.ToString("dd-MM-yyyy").Replace("-", "")
                                    Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    Dim oRS1 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    ' ------------------- take the max for each one 

                                    Dim query As String
                                    query = "select distinct(U_BatchNumberPrefix) From [@Z_ITEMTYPE]"
                                    oRS.DoQuery(query)

                                    Dim u As Integer

                                    For u = 0 To oRS.RecordCount - 1
                                        Dim batchnumberprefix As String = oRS.Fields.Item("U_BatchNumberPrefix").Value

                                        query = "Select max(distnumber) as max from obtn where distnumber like '" & batchnumberprefix & thisDay.ToString("dd-MM-yyyy").Replace("-", "") & "%'"
                                        oRS1.DoQuery(query)

                                        Dim max As String = ""
                                        Dim tmp As Integer = 0
                                        If oRS1.RecordCount = 0 Then
                                            tmp = 0
                                        ElseIf oRS1.RecordCount > 0 Then
                                            max = oRS1.Fields.Item("max").Value
                                            If max <> "" Then
                                                tmp = CInt(max.Substring(max.Length - 3, 3))
                                            Else
                                                tmp = 0
                                            End If
                                        Else
                                            tmp = 0
                                        End If
                                        mHash.Add(batchnumberprefix, tmp)
                                        oRS.MoveNext()
                                    Next


                                    ' ---------------------------------------------
                                    oMatrix = oForm.Items.Item("3").Specific
                                    oMatrix.Columns.Item("2").Editable = True
                                    oMatrix.Columns.Item("5").Editable = True
                                    oMatrix.Columns.Item("10").Editable = True






                                    Dim i As Integer
                                    Dim oDocumentsRow As SAPbouiCOM.Matrix
                                    oDocumentsRow = oForm.Items.Item("35").Specific
                                    oDocumentsRow.Columns.Item("2").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)



                                    oDocumentsRow.Columns.Item("2").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                                    Dim oItemHash As New Hashtable
                                    blnBatchMessage = True

                                    

                                    For i = 1 To oDocumentsRow.VisualRowCount
                                        Dim tmp As Integer = 0

                                        'get item code..

                                      


                                        Dim tmpquery As String = CType(oDocumentsRow.Columns.Item("5").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value


                                        query = "select z.U_TypeCode as TypeCode,z.U_BatchNumberPrefix as prefix, U_Numberofdays as Numberofdays, U_bymonth as bymonth "
                                        query += " from oitm i inner join [@Z_ITEMTYPE] z on z.U_TypeCode = i.U_Z_ItemType where itemcode = '" & tmpquery & "'"
                                        oRS.DoQuery(query)
                                        Dim type As String = oRS.Fields.Item("TypeCode").Value
                                        Dim prefix As String = oRS.Fields.Item("prefix").Value
                                        Dim Numberofdays As String = oRS.Fields.Item("Numberofdays").Value
                                        Dim U_bymonth As String = oRS.Fields.Item("bymonth").Value
                                        Dim ExpiryDate As Date = thisDay '.ToString("dd-MM-yyyy")


                                        If U_bymonth = "Y" Then

                                            ExpiryDate = ExpiryDate.AddMonths(Numberofdays)

                                        Else
                                            ExpiryDate = ExpiryDate.AddDays(Numberofdays)
                                        End If


                                        oMatrix = oForm.Items.Item("3").Specific



                                        If mHash.Contains(prefix) Then
                                            tmp = mHash.Item(prefix)
                                        Else
                                            tmp = 0
                                        End If

                                        'Added by Madhu For Same Batch to be Allocated for Same Item.
                                        Dim tmpString As String
                                        Dim oCombination As String = tmpquery.ToString() & "-" & oProHash.Item(i)
                                        If Not oItemHash.ContainsKey(oCombination) Then
                                            'generation of automatic batch number
                                            tmp = tmp + 1
                                            batchnumber = tmp
                                            tmpString = CStr(tmp)


                                            If tmpString.Length = 1 Then
                                                tmpString = "00" & tmpString
                                            ElseIf tmpString.Length = 2 Then
                                                tmpString = "0" & tmpString
                                            End If
                                            oItemHash.Add(oCombination, tmpString)
                                        Else
                                            tmpString = oItemHash.Item(oCombination)
                                        End If


                                        Dim tmpDisplay As String = prefix & thisDay.ToString("dd-MM-yyyy").Replace("-", "") & tmpString
                                        CType(oMatrix.Columns.Item("2").Cells().Item(1).Specific, SAPbouiCOM.EditText).Value = tmpDisplay
                                        CType(oMatrix.Columns.Item("10").Cells().Item(1).Specific, SAPbouiCOM.EditText).Value = ExpiryDate.ToString("yyyy-MM-dd").Replace("-", "")
                                        Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("1").Specific


                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                                            obutton.Item.Click()
                                        End If

                                        obutton = oForm.Items.Item("createb").Specific
                                        obutton.Item.Enabled = False
                                        If oDocumentsRow.VisualRowCount <> i Then
                                            oDocumentsRow.Columns.Item("2").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        End If

                                        mHash.Remove(prefix)
                                        mHash.Add(prefix, tmp)
                                    Next

                                    blnBatchMessage = False
                                    oProHash = Nothing


                                    For i = 1 To oDocumentsRow.VisualRowCount
                                        oDocumentsRow.Columns.Item("2").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oMatrix.Columns.Item("2").Editable = False
                                        oMatrix.Columns.Item("5").Editable = False
                                        oMatrix.Columns.Item("10").Editable = False
                                    Next


                                End If

                        End Select
                End Select
            ElseIf pVal.FormTypeEx = "65214" Then

                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    flagFIFO = False
                                    flagBatchsetup = False

                                    

                                End If

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    flagFIFO = False
                                    flagBatchsetup = False

                                    Dim oBaseMatrix As SAPbouiCOM.Matrix
                                    oProHash = New Hashtable
                                    oBaseMatrix = oForm.Items.Item("13").Specific
                                    For intRow As Integer = 1 To oBaseMatrix.VisualRowCount
                                        Dim strProOrder As String = CType(oBaseMatrix.Columns.Item("61").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                                        If strProOrder <> "" Then
                                            oProHash.Add(intRow, strProOrder)
                                        End If
                                    Next

                                End If


                        End Select
                End Select


            End If


        Catch ex As Exception
            blnBatchMessage = False
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_InvSO
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
End Class
