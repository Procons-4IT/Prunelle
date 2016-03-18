Public Class clsDeliverySystem
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
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_DeliverySystem Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                                    Dim oForm2 As SAPbouiCOM.Form = oApplication.SBO_Application.Forms.GetForm("-140", 0)
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)



                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oEditText = oForm.Items.Item("4").Specific
                                    Dim cardcode As String = oEditText.Value
                                    Dim query As String = "select * From OCRD where CardCode = '" & cardcode & "'"
                                    Dim oRS As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRS.DoQuery(query)



                                    oEditText = oForm2.Items.Item("U_typepayment").Specific
                                    oEditText.Value = oRS.Fields.Item("U_typepayment").Value

                                    query = "select t0.U_TypeRoute, t0.U_RouteCode,t2.U_DriverCode from [@Z_OCURT] t0 inner join [@Z_CURT1] t1 on t1.DocEntry = t0.DocEntry inner join [@Z_ORUT] t2 on t2.U_RouteCode = t0.U_RouteCode where t1.U_CardCode = '" & cardcode & "'"
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

                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oForm.Items.Item("1").Click()
                                    End If



                                    oEditText = oForm2.Items.Item("U_datetiming").Specific
                                    Dim thisDay As DateTime = DateTime.Today
                                    oEditText.Value = DateTime.Now

                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oForm2 As SAPbouiCOM.Form = oApplication.SBO_Application.Forms.GetForm("-140", 0)
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

                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oForm.Items.Item("1").Click()
                                        End If


                                    End If
                                Catch ex As Exception

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
