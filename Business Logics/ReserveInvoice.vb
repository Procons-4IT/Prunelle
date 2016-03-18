Public Class ReverseInvoice

    Public ItemCode As String
    Public ItemName As String
    Public DocDueDate As Date
    Public CardCode As String
    Public CardName As String
    Public Quantity As Integer
    Public Type As String
    Public TaxCode As String
    Public DocEntry As String
    Public BaseLine As String
    Public ShipDate As Date
    Public WhsCode As String


    Public Sub New(ByVal Itemcode As String, ByVal ItemName As String, DocDueDate As Date, CardCode As String, CardName As String, Quantity As Integer,
                    Type As String, TaxCode As String, DocEntry As String, BaseLine As String, ShipDate As Date, WhsCode As String)
        Me.ItemCode = Itemcode
        Me.ItemName = ItemName
        Me.DocDueDate = DocDueDate
        Me.CardCode = CardCode
        Me.CardName = CardName
        Me.Quantity = Quantity
        Me.Type = Type
        Me.TaxCode = TaxCode
        Me.DocEntry = DocEntry
        Me.BaseLine = BaseLine
        Me.ShipDate = ShipDate
        Me.WhsCode = WhsCode

    End Sub


End Class
