Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System



Public Class main

    Dim xws As XmlWriterSettings = New XmlWriterSettings()

    Dim xw As XmlWriter

    Dim filename As String = "CU_" & Date.Now().ToString("yyyyMMddHHmm")
    Dim folder As String = "xml_" & Date.Now().ToString("yyyyMMdd")
    Dim localbackup As String = Application.StartupPath & "\report\" & folder & "\"
    Dim TempLocation = Application.StartupPath & "\temp\"
    Dim backup As String = "\\192.168.5.20\onsemi_cap\On Semi Xml backup\" & folder & "\"
    'Dim sftp As String = "\\192.168.5.20\sftp\FSC\test\de_xml\" ' TEST
    Dim sftp As String = "\\192.168.5.20\sftp\FSC\de_xml\" ' LIVE
    Dim xmlfile As String = TempLocation & filename & ".xml"
    Dim trgfile As String = TempLocation & filename & ".trg"
    Dim subject As String = "OnSemi B2B Transaction"

    Public Shared TranstimeStart As String = ""
    Public Shared TranstimeEnd As String = ""
    Dim dt As New DataTable


    Private Sub main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim strPath As String = Directory.GetDirectoryRoot(Application.StartupPath)
        Dim Drive As New IO.DriveInfo(strPath)
        Dim Freespace As ULong = Drive.AvailableFreeSpace
        If Freespace >= 1073741824 Then
            CreateXML()
        End If
        Me.Close()


    End Sub

    Private Sub CreateXML()
        Try
            Dim cnt As Integer = 1
attemp:

            If Not Directory.Exists(localbackup) Then
                Directory.CreateDirectory(localbackup)
            End If

            If Not Directory.Exists(backup) Then
                Directory.CreateDirectory(backup)
            End If
            Dim settings As XmlWriterSettings = New XmlWriterSettings
            xws.Encoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
            xws.OmitXmlDeclaration = True

            dt = dtReport()
            If dt.Rows.Count = 0 Then


                If cnt >= 3 Then
                    SendEmail(subject, "No File Created", "", "")
                    GoTo skip
                End If
                cnt += 1
                GoTo attemp
            End If

            Dim xml_string, xml_update As String
            xml_string = ConvertDatatableToXML(dt)
            Dim dt1 As New DataTable
            dt1 = saveXMLLogs(xml_string, filename, TranstimeStart, TranstimeEnd)
            dt.Clear()
            dt = dt1

            If dt1.Rows.Count = 0 Then
                SendEmail(subject, "No File Created", "", "")
                GoTo skip
            End If

            xml()
            trg()
            PostToSFTP()

            'xml_update = ConvertDatatableToXML(dt1)
            SendEmail(subject, CreateMsgBody(), localbackup & filename & ".xml", localbackup & filename & ".trg")
skip:

        Catch ex As Exception
            'MsgBox(ex.Message)
            'errormsg:
            SendEmailerror(subject & " ERROR", ex.Message)
        End Try



    End Sub


    Public Shared Function ConvertDatatableToXML(ByVal dt As DataTable) As String
        If dt.Rows.Count > 0 Then
            dt.TableName = "XMLLogs"
            Dim writer As New System.IO.StringWriter()
            dt.WriteXml(writer, XmlWriteMode.WriteSchema, False)
            Dim result As String = writer.ToString()
            Return result
        End If
        Return Nothing
    End Function

    Public Function saveXMLLogs(ByVal xml As String, ByVal Filename As String, ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataTable
        Dim sql_handler As New SQLHandler
        Dim ds As New DataSet
        Dim dt As New DataTable
        'Dim strSQL As String = "usp_TRN_ONSemi_xml_insert"
        Dim strSQL As String = "usp_TRN_ONSemi_xml_insert_JOE_2024"

        sql_handler.CreateParameter(4)
        sql_handler.SetParameterValues(0, "@XMLdoc", SqlDbType.Xml, xml)
        sql_handler.SetParameterValues(1, "@Filename", SqlDbType.NVarChar, Filename)
        sql_handler.SetParameterValues(2, "@FromDate", SqlDbType.DateTime, DateFrom)
        sql_handler.SetParameterValues(3, "@ToDate", SqlDbType.DateTime, DateTo)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(strSQL, ds, CommandType.StoredProcedure) Then
                dt = ds.Tables(0)
            End If
        End If

        Return dt
    End Function

    Private Sub xml()


        xws.Indent = True
        'xw = XmlWriter.Create((backup & filename & ".xml"), xws)
        xw = XmlWriter.Create((xmlfile), xws)

        xws.Indent = True
        xws.NewLineOnAttributes = True
        data()


    End Sub


    Private Sub trg()
        System.IO.File.Copy(xmlfile, trgfile, True)

    End Sub

    Private Sub PostToSFTP()
        Dim cnt As Integer = 0
        Try
attemp:

            cnt += 1
            Dim fileList() As String = Nothing

            fileList = Directory.GetFiles(TempLocation, "*.*", SearchOption.TopDirectoryOnly)

            For Each files As String In fileList

                Dim fileName As String = Path.GetFileName(files)

                System.IO.File.Copy(files, sftp & fileName, True)
                System.IO.File.Copy(files, backup & fileName, True)
                System.IO.File.Copy(files, localbackup & fileName, True)

                If File.Exists(sftp & fileName) Then
                    System.IO.File.Delete(files)
                End If

            Next


            'If Not File.Exists(sftp & filename & ".xml") And Not File.Exists(sftp & filename & ".trg") Then
            '    If File.Exists(xmlfile) And File.Exists(trgfile) Then
            '        System.IO.File.Copy(xmlfile, sftp & filename & ".xml", True)
            '        System.IO.File.Copy(xmlfile, sftp & filename & ".trg", True)

            '        System.IO.File.Copy(xmlfile, backup & filename & ".xml", True)
            '        System.IO.File.Copy(xmlfile, backup & filename & ".trg", True)
            '    Else
            '        CreateXML()
            '    End If
            'Else
            '    'Email Success Confirmation
            'End If
        Catch ex As Exception
            If cnt >= 3 Then
                SendEmailerror(subject & " ERROR", ex.Message)
                GoTo skip
            End If
            GoTo attemp
        End Try
skip:
    End Sub

    Private Sub data()

        xw.WriteStartDocument()

        xw.WriteStartElement("dataexchange")
        'SUBCON
        xw.WriteStartElement("subcon")
        xw.WriteElementString("inv_org", "PH4")
        xw.WriteElementString("subcon_name", "ATEC")
        xw.WriteEndElement() ' END of SUBCON



        For Each row In dt.Rows
            xw.WriteStartElement("transaction") ' START of TRANSACTION
            xw.WriteElementString("transaction_code", row("transaction_code"))

            transaction_date(row)
            invoice(row)
            customer_lot(row)
            source_lot(row)
            assembly_step(row)
            test_insertion(row)
            target_lot(row)
            comp_lot(row)
            date_code(row)
            reject(row)
            bonus(row)
            commit_date(row)
            hold_lot(row)

            xw.WriteEndElement() ' END of TRANSACTION
        Next

        xw.WriteEndElement()
        xw.WriteEndDocument()
        xw.Flush()
        xw.Close()
    End Sub

    Public Sub transaction_date(ByRef row As DataRow)

        xw.WriteStartElement("transaction_date")
        Dim test As String = row("tran_month")
        xw.WriteElementString("tran_month", row("tran_month"))
        xw.WriteElementString("tran_day", row("tran_day"))
        xw.WriteElementString("tran_year", row("tran_year"))
        xw.WriteElementString("tran_time", row("tran_time"))
        xw.WriteEndElement()
    End Sub

    Public Sub invoice(ByRef row As DataRow)
        xw.WriteStartElement("invoice")
        xw.WriteElementString("invoice_number", row("invoice_number"))
        xw.WriteElementString("target_location", row("target_location"))
        xw.WriteEndElement()
    End Sub

    Public Sub customer_lot(ByRef row As DataRow)
        xw.WriteStartElement("customer_lot")
        xw.WriteElementString("cust_lot_type", row("cust_lot_type"))
        xw.WriteElementString("cust_lot_num", row("cust_lot_num"))
        xw.WriteEndElement()
    End Sub

    Public Sub source_lot(ByRef row As DataRow)
        xw.WriteStartElement("source_lot")
        xw.WriteElementString("src_lot_type", row("src_lot_type"))
        xw.WriteElementString("src_lot_number", row("src_lot_number"))
        xw.WriteElementString("src_lot_category", row("src_lot_category"))
        xw.WriteElementString("src_lot_qty", row("src_lot_qty"))
        xw.WriteElementString("src_device", row("src_device"))
        xw.WriteElementString("src_location", row("src_location"))
        xw.WriteEndElement()
    End Sub

    Public Sub assembly_step(ByRef row As DataRow)
        xw.WriteStartElement("assembly_step")
        xw.WriteElementString("step_name", row("step_name"))
        xw.WriteElementString("step_out_qty", row("step_out_qty"))
        xw.WriteEndElement()
    End Sub

    Public Sub test_insertion(ByRef row As DataRow)
        xw.WriteStartElement("test_insertion")
        xw.WriteElementString("insertion_name", row("insertion_name"))
        xw.WriteElementString("insertion_qty_out", row("insertion_qty_out"))
        xw.WriteEndElement()
    End Sub

    Public Sub target_lot(ByRef row As DataRow)
        xw.WriteStartElement("target_lot")
        xw.WriteElementString("tgt_lot_type", row("tgt_lot_type"))
        xw.WriteElementString("tgt_lot_number", row("tgt_lot_number"))
        xw.WriteElementString("tgt_lot_category", row("tgt_lot_category"))
        xw.WriteElementString("tgt_lot_qty", row("tgt_lot_qty"))
        xw.WriteElementString("tgt_device", row("tgt_device"))
        xw.WriteElementString("tgt_location", row("tgt_location"))
        xw.WriteEndElement()
    End Sub

    Public Sub comp_lot(ByRef row As DataRow)
        xw.WriteStartElement("comp_lot")
        xw.WriteElementString("comp_lot_type", row("comp_lot_type"))
        xw.WriteElementString("comp_lot_number", row("comp_lot_number"))
        xw.WriteElementString("comp_lot_category", row("comp_lot_category"))
        xw.WriteElementString("comp_lot_qty", row("comp_lot_qty"))
        xw.WriteElementString("comp_device", row("comp_device"))
        xw.WriteElementString("comp_location", row("comp_location"))
        xw.WriteEndElement()
    End Sub

    Public Sub date_code(ByRef row As DataRow)
        xw.WriteStartElement("date_code")
        xw.WriteElementString("date_code1", row("date_code1"))
        xw.WriteElementString("date_code2", row("date_code2"))
        xw.WriteEndElement()
    End Sub

    Public Sub reject(ByRef row As DataRow)
        xw.WriteStartElement("reject")
        xw.WriteElementString("reject_code", row("reject_code"))
        xw.WriteElementString("reject_location", row("reject_location"))
        xw.WriteElementString("reject_qty", row("reject_qty"))
        xw.WriteElementString("reject_comment", row("reject_comment"))
        xw.WriteEndElement()
    End Sub

    Public Sub bonus(ByRef row As DataRow)
        xw.WriteStartElement("bonus")
        xw.WriteElementString("bonus_code", row("bonus_code"))
        xw.WriteElementString("bonus_location", row("bonus_location"))
        xw.WriteElementString("bonus_qty", row("bonus_qty"))
        xw.WriteElementString("bonus_comment", row("bonus_comment"))
        xw.WriteEndElement()
    End Sub

    Public Sub commit_date(ByRef row As DataRow)
        xw.WriteStartElement("commit_date")
        xw.WriteElementString("commit_month", row("commit_month"))
        xw.WriteElementString("commit_day", row("commit_day"))
        xw.WriteElementString("commit_year", row("commit_year"))
        xw.WriteElementString("commit_time", row("commit_time"))
        xw.WriteEndElement()
    End Sub

    Public Sub hold_lot(ByRef row As DataRow)
        xw.WriteStartElement("hold_lot")
        xw.WriteElementString("held_reason", row("held_reason"))
        xw.WriteElementString("held_location", row("held_location"))
        xw.WriteElementString("held_comment", row("held_comment"))
        xw.WriteEndElement()
    End Sub

    Public Shared Function dtReport() As DataTable
        ' Dim getdate As Date = Date.Now()

        'Dim getdate As Date = "11/16/2017 9:00:05 PM"
        Dim dt As New DataTable
        Dim ds As New DataSet

        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        Dim GetTransactionRageSql = "EXEC usp_TRN_OnSemi_GetTransaction_RageTime_Xml"

        Dim GetTransactionRageDS As DataSet = FillDataset(GetTransactionRageSql)
        TranstimeStart = GetTransactionRageDS.Tables(0).Rows(0).Item("TranstimeStart").ToString()
        TranstimeEnd = GetTransactionRageDS.Tables(0).Rows(0).Item("TranstimeEnd").ToString()



        'Dim strSQL As String = "usp_TRN_ONSemi_xml '" & TranstimeStart & "' , '" & TranstimeEnd & "'"

        ''manual sending
        'Dim strSQL As String = "usp_TRN_ONSemi_LotMerge_test '2021-01-22 00:19:50.000','2021-01-22 00:20:00.999'"
        Dim strSQL As String = "usp_TRN_ONSemi_xml '2024-02-19 13:00:01.000','2024-02-19 17:00:01.000'"
        'Dim strSQL As String = "usp_TRN_ONSemi_xml_manual"
        'Dim strSQL As String = "usp_TRN_ONSemi_xml_manual_JOE_2024"
        'Dim strSQL As String = "usp_TRN_ONSemi_xml_manual_per_lot"
        'Dim strSQL As String = "SELECT * FROM tbl_OnSemi_Manual " & _
        '                       "WHERE(tran_date BETWEEN '2018-07-27 11:00:01.000' AND '2018-07-27 13:00:00.999') " & _
        '                       "ORDER BY tran_date, CONVERT(int, tgt_lot_qty)"

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(strSQL, ds, CommandType.Text) Then
                dt = ds.Tables(0)
            End If
        End If

        sql_handler.CloseConnection()

        Return dt
    End Function

    Private Sub SendEmail(ByVal strSubject As String, ByVal strMessage As String, ByVal xmlfile As String, ByVal trgfile As String)
        '8/12/2010
        'ATECPHIL MAILHOST
        'MULTIPLE RECIPIENTS
        'ONE FILE ATTACHMENT

        ' Try
        Dim MailMsg As New MailMessage()
        MailMsg.IsBodyHtml = True
        MailMsg.Subject = strSubject.Trim()
        MailMsg.Body = strMessage.Trim() & vbCrLf
        MailMsg.Priority = MailPriority.High
        MailMsg.IsBodyHtml = True


        If Not xmlfile = "" Then
            Dim MsgAttach As New Attachment(xmlfile)
            MailMsg.Attachments.Add(MsgAttach)
        End If

        If Not trgfile = "" Then
            Dim MsgAttach As New Attachment(trgfile)
            MailMsg.Attachments.Add(MsgAttach)
        End If

        'Email Recipients
        Dim ds As New DataSet
        Dim conEmail As New SqlClient.SqlConnection
        Dim Sqlemail As String = "usp_SPT_AutoEmail_GetRecipients 16"

        ConnectToMES_ATEC(conEmail)
        ' Try
        If (ExecuteQuery(Sqlemail, conEmail, ds, "EmailList")) Then
            Dim i As Integer
            If ds.Tables("EmailList").Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If Trim(ds.Tables(0).Rows(i).Item("EmailTo")) = "True" Then
                        MailMsg.To.Add(New MailAddress(Trim(ds.Tables(0).Rows(i).Item("Email_Address"))))
                    ElseIf Trim(ds.Tables(0).Rows(i).Item("EmailCC")) = "True" Then
                        MailMsg.CC.Add(New MailAddress(Trim(ds.Tables(0).Rows(i).Item("Email_Address"))))
                    ElseIf Trim(ds.Tables(0).Rows(i).Item("EmailFrom")) = "True" Then
                        MailMsg.From = New MailAddress(Trim(ds.Tables(0).Rows(i).Item("Email_Address")))
                    End If
                Next
                i = Nothing
            End If
        End If
        'Catch ex As Exception
        '    'MessageBox.Show(ex.Message)
        '    Exit Sub
        'Finally
        '    ConnectToMSDYNAMICS(conEmail)
        'End Try

        '--ATECPHIL--
        'Dim SmtpMail As New SmtpClient
        'SmtpMail.Host = "192.168.1.7"
        'SmtpMail.Port = 25
        'SmtpMail.UseDefaultCredentials = True
        'SmtpMail.Credentials = New System.Net.NetworkCredential("administrator", "trator#$0809")
        Dim Sql As String = "usp_Get_ATEC_EmailServer_V2"
        ServicePointManager.SecurityProtocol = CType(48 Or 192 Or 768 Or 3072, SecurityProtocolType)
        ds = FillDataset(Sql, ConnectionString)
        Dim Username, Password As String
        Dim SmtpMail As New SmtpClient
        SmtpMail.Host = ds.Tables(0).Rows(0).Item("Host").ToString() '"Atec-mail"
        SmtpMail.Port = ds.Tables(0).Rows(0).Item("Port").ToString()
        Username = ds.Tables(0).Rows(0).Item("Username").ToString()
        Password = ds.Tables(0).Rows(0).Item("Password").ToString()
        SmtpMail.UseDefaultCredentials = True
        SmtpMail.Credentials = New System.Net.NetworkCredential(Username, Password)
        SmtpMail.EnableSsl = True
        '--ATECPHIL--

        SmtpMail.Send(MailMsg)
        MailMsg = Nothing
        SmtpMail = Nothing

        Threading.Thread.Sleep(5000)
        'Catch exEmail As Exception
        '    'Message Error
        '    Me.Cursor = Cursors.Default
        '    'MessageBox.Show(exEmail.ToString, "Error Dialog MWO Checker Email Sending", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Exit Sub
        '    'MsgBox(exEmail.Message)
        'End Try
    End Sub
    Private Function CreateMsgBody() As String

        CreateMsgBody = ""
        CreateMsgBody = "<html><body><pre>"

        CreateMsgBody &= "<pre>Hello Team </pre>"

        CreateMsgBody &= "<br><pre>See attached file for " & subject & " is successfully posted in SFTP</pre>"

        CreateMsgBody &= "Transaction Time From " & TranstimeStart & " to " & TranstimeEnd



        CreateMsgBody &= "</pre></body></html>"
    End Function
    Private Sub SendEmailerror(ByVal strSubject As String, ByVal strMessage As String)
        '8/12/2010
        'ATECPHIL MAILHOST
        'MULTIPLE RECIPIENTS
        'ONE FILE ATTACHMENT

        Dim MailMsg As New MailMessage()
        MailMsg.IsBodyHtml = True
        MailMsg.Subject = strSubject.Trim()
        MailMsg.Body = strMessage.Trim() & vbCrLf
        MailMsg.Priority = MailPriority.High
        MailMsg.IsBodyHtml = True

        'Email Recipients
        Dim ds As New DataSet
        Dim conEmail As New SqlClient.SqlConnection
        Dim Sqlemail As String = "usp_SPT_AutoEmail_GetRecipients 16"

        ConnectToMES_ATEC(conEmail)
        ' Try
        If (ExecuteQuery(Sqlemail, conEmail, ds, "EmailList")) Then
            Dim i As Integer
            If ds.Tables("EmailList").Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If Trim(ds.Tables(0).Rows(i).Item("EmailTo")) = "True" Then
                        MailMsg.To.Add(New MailAddress(Trim(ds.Tables(0).Rows(i).Item("Email_Address"))))
                    ElseIf Trim(ds.Tables(0).Rows(i).Item("EmailCC")) = "True" Then
                        MailMsg.CC.Add(New MailAddress(Trim(ds.Tables(0).Rows(i).Item("Email_Address"))))
                    ElseIf Trim(ds.Tables(0).Rows(i).Item("EmailFrom")) = "True" Then
                        MailMsg.From = New MailAddress(Trim(ds.Tables(0).Rows(i).Item("Email_Address")))
                    End If
                Next



                i = Nothing
            End If
        End If

        '--ATECPHIL--
        'Dim SmtpMail As New SmtpClient
        'SmtpMail.Host = "192.168.1.7"
        'SmtpMail.Port = 25
        'SmtpMail.UseDefaultCredentials = True
        'SmtpMail.Credentials = New System.Net.NetworkCredential("administrator", "trator#$0809")
        Dim Sql As String = "SELECT * FROM  ATEC_EmailServer WHERE ID = 1"
        ds = FillDataset(Sql, ConnectionString)
        Dim Username, Password As String
        Dim SmtpMail As New SmtpClient
        SmtpMail.Host = ds.Tables(0).Rows(0).Item("Host").ToString() '"Atec-mail"
        SmtpMail.Port = ds.Tables(0).Rows(0).Item("Port").ToString()
        Username = ds.Tables(0).Rows(0).Item("Username").ToString()
        Password = ds.Tables(0).Rows(0).Item("Password").ToString()
        SmtpMail.UseDefaultCredentials = True
        SmtpMail.Credentials = New System.Net.NetworkCredential(Username, Password)

        '--ATECPHIL--

        SmtpMail.Send(MailMsg)
        MailMsg = Nothing
        SmtpMail = Nothing


    End Sub
End Class
