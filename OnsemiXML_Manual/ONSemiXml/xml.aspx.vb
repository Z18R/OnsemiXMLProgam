Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient

Partial Class MRP_xml
    'Inherits System.Web.UI.Page
    Dim xws As XmlWriterSettings = New XmlWriterSettings()
    Dim xw As XmlWriter
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)


        xws.Indent = True
        xw = XmlWriter.Create("c:\\TEMP\" & Date.Now().ToString("MMddYYYY") & ".xml", xws)
        xws.Indent = True
        xws.NewLineOnAttributes = True
        Dim X = 1
        Dim Y = 2

        xw.WriteStartDocument()

        xw.WriteStartElement("dataexchange")
        'subcon
        xw.WriteStartElement("subcon")
        xw.WriteElementString("inv_org", X)
        xw.WriteElementString("subcon_name", Y)
        xw.WriteEndElement() ' END NG SUBCON

        Dim dt As New DataTable
        dt = dtReport()

        For Each row In dt.Rows
            xw.WriteStartElement("transaction") ' START NG TRANSACTION
            xw.WriteElementString("transaction_code", row("TransID"))


            xw.WriteStartElement("invoice") 'START NG transaction_date
            xw.WriteElementString("invoice_number", row("InvoiceNo"))
            xw.WriteElementString("target_location", "")
            xw.WriteEndElement()
            xw.WriteEndElement() ' END NG TRANSACTION
        Next




        'END NG DATA EXCHANGE
        xw.WriteEndElement()
        xw.WriteEndDocument()
        xw.Flush()
        xw.Close()

    End Sub

    Public Sub CreateChildNode(ByVal rows As String)
        'xw.WriteStartElement("transaction_date") 'START NG transaction_date
        'xw.WriteElementString("tran_month", row("tran_month"))
        'xw.WriteElementString("tran_day", row("tran_day"))
        'xw.WriteElementString("tran_year", row("tran_year"))
        'xw.WriteElementString("tran_time", row("tran_time"))
        'xw.WriteEndElement() 'END NG transaction_date


    End Sub

    Public Shared Function dtReport() As DataTable
        Dim dt As New DataTable
        Dim ds As New DataSet

        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim strSQL As String = "SELECT TOP 10  'INVRCPTS' AS TransID, B.InvoiceNo,A.LotID,A.MaterialQty,C.MaterialID AS DEVICE, 'NORMAL' AS Lot_Category,		 " & _
                                "MONTH(A.ReceivedDate) AS tran_month,DAY(A.ReceivedDate) AS tran_day,YEAR(A.ReceivedDate) AS tran_year,					 " & _
                                "convert(varchar, A.ReceivedDate, 108) AS tran_time,A.ReceivedDate,'PS' cust_lot_type FROM TRN_Receive_Material A 		 " & _
                                "LEFT JOIN TRN_Receive B ON A.ReceiveCode = B.ReceiveCode																 " & _
                                "LEFT JOIN PS_Material C ON A.MaterialCode = C.MaterialCode																 "
        sql_handler.CreateParameter(3)

        sql_handler.SetToAXDB()
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(strSQL, ds, CommandType.Text) Then
                dt = ds.Tables(0)
            End If
        End If

        sql_handler.CloseConnection()

        Return dt
    End Function
End Class
