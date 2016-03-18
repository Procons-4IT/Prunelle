Public Class clsBP
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
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BPMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD





                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "1" Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_typepayment", 0).ToString().Trim = "None" Or oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_sequencetype", 0).ToString().Trim = "None" Then
                                            oApplication.Utilities.Message("UDF payment type and sequencetype must not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                        End If
                                      


                                        'here we are 

                                        Dim oDataTable As SAPbobsCOM.UserTable
                                        oDataTable = oApplication.Company.UserTables.Item("TableDelivery")
                                        oGrid = oForm.Items.Item("oGridTD").Specific
                                        Dim i As Integer



                                        For i = 0 To oGrid.DataTable.Rows.Count - 1
                                            If oDataTable.GetByKey(oGrid.DataTable.GetValue("Code", i)) Then

                                                oDataTable.UserFields.Fields.Item("U_CardCode").Value = oApplication.Utilities.getEdittextvalue(oForm, "5")
                                                oDataTable.UserFields.Fields.Item("U_CardName").Value = oApplication.Utilities.getEdittextvalue(oForm, "7")
                                                oDataTable.UserFields.Fields.Item("U_Type").Value = oGrid.DataTable.GetValue("U_Type", i)
                                                oDataTable.UserFields.Fields.Item("U_Monday").Value = oGrid.DataTable.GetValue("U_Monday", i)
                                                oDataTable.UserFields.Fields.Item("U_Tuesday").Value = oGrid.DataTable.GetValue("U_Tuesday", i)
                                                oDataTable.UserFields.Fields.Item("U_Wednesday").Value = oGrid.DataTable.GetValue("U_Wednesday", i)
                                                oDataTable.UserFields.Fields.Item("U_Thursday").Value = oGrid.DataTable.GetValue("U_Thursday", i)
                                                oDataTable.UserFields.Fields.Item("U_Friday").Value = oGrid.DataTable.GetValue("U_Friday", i)
                                                oDataTable.UserFields.Fields.Item("U_Saturday").Value = oGrid.DataTable.GetValue("U_Saturday", i)
                                                oDataTable.UserFields.Fields.Item("U_Sunday").Value = oGrid.DataTable.GetValue("U_Sunday", i)
                                                Dim intStatus As Integer = oDataTable.Update()
                                            Else
                                                oDataTable.Code = getIdCodeDriver()
                                                oDataTable.Name = getIdCodeDriver()
                                                oDataTable.UserFields.Fields.Item("U_CardCode").Value = oApplication.Utilities.getEdittextvalue(oForm, "5")
                                                oDataTable.UserFields.Fields.Item("U_CardName").Value = oApplication.Utilities.getEdittextvalue(oForm, "7")
                                                oDataTable.UserFields.Fields.Item("U_Type").Value = oGrid.DataTable.GetValue("U_Type", i)
                                                oDataTable.UserFields.Fields.Item("U_Monday").Value = oGrid.DataTable.GetValue("U_Monday", i)
                                                oDataTable.UserFields.Fields.Item("U_Tuesday").Value = oGrid.DataTable.GetValue("U_Tuesday", i)
                                                oDataTable.UserFields.Fields.Item("U_Wednesday").Value = oGrid.DataTable.GetValue("U_Wednesday", i)
                                                oDataTable.UserFields.Fields.Item("U_Thursday").Value = oGrid.DataTable.GetValue("U_Thursday", i)
                                                oDataTable.UserFields.Fields.Item("U_Friday").Value = oGrid.DataTable.GetValue("U_Friday", i)
                                                oDataTable.UserFields.Fields.Item("U_Saturday").Value = oGrid.DataTable.GetValue("U_Saturday", i)
                                                oDataTable.UserFields.Fields.Item("U_Sunday").Value = oGrid.DataTable.GetValue("U_Sunday", i)
                                                Dim intStatus As Integer = oDataTable.Add()


                                            End If
                                        Next

                                    End If
                                End If



                        End Select

                    Case False
                        Select Case pVal.EventType
                          
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'add the button 
                                Dim oStatic As SAPbouiCOM.StaticText = oForm.Items.Item("2013").Specific


                                Dim oNewItem As SAPbouiCOM.Item = oForm.Items.Add("freshst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                                Dim oNewItem1 As SAPbouiCOM.Item = oForm.Items.Add("freshedit", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                Dim oNewItem2 As SAPbouiCOM.Item = oForm.Items.Add("frznst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                                Dim oNewItem3 As SAPbouiCOM.Item = oForm.Items.Add("frznedit", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)


                                oNewItem.Left = oStatic.Item.Left
                                oNewItem.Width = oStatic.Item.Width
                                oNewItem.Height = oStatic.Item.Height
                                oNewItem.Top = oStatic.Item.Top + oStatic.Item.Height

                                Dim freshstatic As SAPbouiCOM.StaticText
                                freshstatic = oNewItem.Specific
                                freshstatic.Caption = "fresh"
                                freshstatic.Item.FromPane = 1
                                freshstatic.Item.ToPane = 1


                                oNewItem1.Left = freshstatic.Item.Left + freshstatic.Item.Width
                                oNewItem1.Width = freshstatic.Item.Width
                                oNewItem1.Height = freshstatic.Item.Height
                                oNewItem1.Top = freshstatic.Item.Top

                                Dim freshedit As SAPbouiCOM.EditText
                                freshedit = oNewItem1.Specific
                                freshedit.DataBind.SetBound(True, "OCRD", "U_freshroute")
                                freshedit.Item.FromPane = 1
                                freshedit.Item.ToPane = 1


                                oNewItem2.Left = freshstatic.Item.Left
                                oNewItem2.Width = freshstatic.Item.Width
                                oNewItem2.Height = freshstatic.Item.Height
                                oNewItem2.Top = freshstatic.Item.Top + freshstatic.Item.Height

                                Dim frozenstatic As SAPbouiCOM.StaticText
                                frozenstatic = oNewItem2.Specific
                                frozenstatic.Caption = "frozen"
                                frozenstatic.Item.FromPane = 1
                                frozenstatic.Item.ToPane = 1


                                oNewItem3.Left = frozenstatic.Item.Left + frozenstatic.Item.Width
                                oNewItem3.Width = frozenstatic.Item.Width
                                oNewItem3.Height = frozenstatic.Item.Height
                                oNewItem3.Top = frozenstatic.Item.Top

                                Dim frozenedit As SAPbouiCOM.EditText
                                frozenedit = oNewItem3.Specific
                                frozenedit.DataBind.SetBound(True, "OCRD", "U_frozenroute")
                                frozenedit.Item.FromPane = 1
                                frozenedit.Item.ToPane = 1

                                oForm.DataSources.DataTables.Add("DT_5")
                                oApplication.Utilities.AddControls(oForm, "oGridTD", "136", SAPbouiCOM.BoFormItemTypes.it_GRID, "4", 4, 4, "", "", 300, 150, 50)
                                oGrid = oForm.Items.Item("oGridTD").Specific
                                oForm.Items.Item("oGridTD").Left = oForm.Items.Item("136").Left + oForm.Items.Item("136").Width + 1
                                oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_5")

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region



    Private Function getIdCodeDriver() As Integer
        'Calling a Query . we select the database to get the minorderquantity that is not displayed on the form.
        Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim query As String = "select count(*) + 1 as max from [@TableDelivery]"
        oRS.DoQuery(query)
        Dim max As Integer = oRS.Fields.Item("max").Value
        Return max
    End Function


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_InvSO
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                  

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) And
                BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then


                DataBind(oForm)
            ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                DataBind(oForm)
            End If
        Catch ex As Exception
            ' MsgBox(ex.Message)
        End Try
    End Sub




    Private Sub DataBind(aform As SAPbouiCOM.Form)
        Dim strFrmCardCode As String
        Dim oDataTable As SAPbouiCOM.DataTable
        'strFrmCardCode = oApplication.Utilities.getEdittextvalue(aform, "5")
        oEditText = aform.Items.Item("5").Specific
        strFrmCardCode = oEditText.Value


        Dim query As String = "select Code,Name,U_CardCode,U_CardName,U_Type,U_Monday,U_Tuesday,U_Wednesday,U_Thursday,U_Friday,U_Saturday,U_Sunday "
        query = query & " from [@TableDelivery] where U_CardCode= '" & strFrmCardCode & "'"
        oDataTable = oForm.DataSources.DataTables.Item("DT_5")
        oDataTable.ExecuteQuery(query)
        oGrid = oForm.Items.Item("oGridTD").Specific
        Dim oRs As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        If oRs.RecordCount = 0 Then

            'oDataTable.Rows.Add(1)
            oDataTable.SetValue("U_Type", 0, "Fresh")
            oDataTable.Rows.Add(1)
            oDataTable.SetValue("U_Type", 1, "Frozen")
            oDataTable.Rows.Add(1)
            oDataTable.SetValue("U_Type", 2, "Dessert")
        Else
            'oDataTable.ExecuteQuery(query)
        End If
        'Dim query As String = "select rd.ItemCode as ItemCode,rd.Dscription as Dscription,ord.DocDueDate as DocDueDate,ord.CardCode as 'Card Code', rd.Quantity as Quantity ,ord.CardName as 'Card Name',Convert(VarChar(1),'N') As 'Select' , o.U_Z_ItemType as Type from ORDR ord inner join RDR1 rd on rd.DocEntry = ord.DocEntry inner join OITM o on o.ItemCode = rd.ItemCode where " & strCondition & " order by o.U_Z_ItemType"
        oGrid.DataTable = oDataTable
        oGrid.AutoResizeColumns()
        oGrid.Columns.Item("U_Monday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("U_Tuesday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("U_Wednesday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("U_Thursday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("U_Friday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("U_Saturday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("U_Sunday").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox


        oGrid.Columns.Item("U_CardCode").Visible = False
        oGrid.Columns.Item("U_CardName").Visible = False
        oGrid.Columns.Item("Code").Visible = False
        oGrid.Columns.Item("Name").Visible = False
        oGrid.Columns.Item("U_Type").Editable = False


    End Sub
End Class
