Imports SAPbobsCOM

Public Class clsProduction
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
    Private blnFlag As Boolean = False '
    Private strItem As String
    Private Shared mArray As Hashtable
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If (pVal.FormTypeEx = frm_Issue_For_Production Or pVal.FormTypeEx = frm_Issue_Inventory Or pVal.FormTypeEx = frm_InventoryTransfer) Then

                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD



                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    batchessetup.flagFIFO = True
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then


                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    oMatrix = oForm.Items.Item("13").Specific
                                    Dim docnum As String = oApplication.Utilities.getMatrixValues(oMatrix, "61", 0)
                                    MsgBox(docnum)
                                End If

                        End Select

                End Select

            ElseIf pVal.FormTypeEx = frm_batch_number_selection Then



                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If batchessetup.flagFIFO = True Then
                                    Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("btnFIFOB").Specific
                                    obutton.Item.Enabled = True
                                ElseIf batchessetup.flagFIFO = False Then
                                    Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("btnFIFOB").Specific
                                    obutton.Item.Enabled = False
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "16" Then
                                    If batchessetup.flagFIFO = False Then
                                        Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("16").Specific
                                        obutton.Item.Enabled = False
                                        obutton = oForm.Items.Item("btnFIFOB").Specific
                                        obutton.Item.Enabled = False
                                    End If
                                End If
                                If pVal.ItemUID = "btnFIFOB" Then
                                    oMatrix = oForm.Items.Item("3").Specific
                                    Dim i As Integer
                                    Dim oRecSet As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    Dim oRecSet1 As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("1").Specific
                                    Dim arrowbutton As SAPbouiCOM.Button = oForm.Items.Item("48").Specific

                                    For i = 1 To oMatrix.VisualRowCount
                                        If i > 1 Then
                                            oMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        End If

                                        Dim availablebatches As SAPbouiCOM.Matrix = oForm.Items.Item("4").Specific
                                        Dim oitem As String = CType(oMatrix.Columns.Item("1").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value
                                        Dim Qtyneeded As Double = CType(oMatrix.Columns.Item("55").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value
                                        oRecSet.DoQuery("select itemcode, ExpDate as exp_date, BatchNum,Quantity, WhsCode from oibt where Quantity<>0 and ItemCode = '" & oitem & "' order by exp_date ")


                                        While Qtyneeded <> 0


                                            Dim j As Integer
                                            For j = 1 To availablebatches.VisualRowCount
                                                Dim batchnum As String = oRecSet.Fields.Item("BatchNum").Value
                                                Dim Quantityavailableforbatches As Double = oRecSet.Fields.Item("Quantity").Value

                                                If batchnum = oApplication.Utilities.getMatrixValues(availablebatches, "0", j) Then
                                                    If Qtyneeded > oApplication.Utilities.getMatrixValues(availablebatches, "3", j) Then
                                                        oApplication.Utilities.SetMatrixValues(availablebatches, "4", j, oApplication.Utilities.getMatrixValues(availablebatches, "3", j))
                                                        availablebatches.Columns.Item("1").Cells.Item(j).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                        arrowbutton.Item.Click()
                                                        Qtyneeded = CType(oMatrix.Columns.Item("55").Cells().Item(i).Specific, SAPbouiCOM.EditText).Value
                                                        obutton.Item.Click()
                                                        Exit For

                                                    ElseIf Qtyneeded <= oApplication.Utilities.getMatrixValues(availablebatches, "3", j) Then
                                                        oApplication.Utilities.SetMatrixValues(availablebatches, "4", j, Qtyneeded)
                                                        availablebatches.Columns.Item("1").Cells.Item(j).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                        arrowbutton.Item.Click()
                                                        Qtyneeded = 0
                                                        obutton.Item.Click()
                                                        Exit For
                                                    End If



                                                End If

                                            Next

                                            'availablebatches.Columns.Item("1").Cells.Item(col).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                            oRecSet.MoveNext()
                                        End While

                                    Next

                                End If



                        End Select


                    Case False

                        Select Case pVal.EventType


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "16" Then
                                    If batchessetup.flagFIFO = False Then
                                        Dim obutton As SAPbouiCOM.Button = oForm.Items.Item("16").Specific
                                        obutton.Item.Enabled = False
                                        obutton = oForm.Items.Item("btnFIFOB").Specific
                                        obutton.Item.Enabled = False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'add the label in the sales order form.
                                Dim obutton1 As SAPbouiCOM.Button = oForm.Items.Item("16").Specific
                                Dim oNewItem As SAPbouiCOM.Item = oForm.Items.Add("btnFIFOB", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                                oNewItem.Left = obutton1.Item.Left
                                oNewItem.Width = obutton1.Item.Width
                                oNewItem.Height = obutton1.Item.Height
                                oNewItem.Top = obutton1.Item.Top - obutton1.Item.Height
                                Dim obutton As SAPbouiCOM.Button
                                obutton = oNewItem.Specific
                                obutton.Caption = "FIFO selection"

                                If batchessetup.flagFIFO = True Then
                                    obutton.Item.Enabled = True
                                Else
                                    obutton.Item.Enabled = False
                                End If


                        End Select
                End Select
            ElseIf pVal.FormTypeEx = "720" Then

                Select Case pVal.BeforeAction
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                mArray = New Hashtable()


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then






                                End If



                        End Select

                    Case True
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
