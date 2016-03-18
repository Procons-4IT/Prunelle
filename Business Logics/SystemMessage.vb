Public Class SystemMessage
    Inherits clsBase
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = "0" Then
                If pVal.Before_Action Then

                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            If blnBatchMessage Then
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
