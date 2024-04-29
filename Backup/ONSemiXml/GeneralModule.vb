Imports System.Threading
Imports System.Data.OleDb
Module GeneralModule
    'Access 2007
    'Public ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|WaterBillingSystemDatabase.mdb"
    'Access 2010
    'Public ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|CustomersDatabase.accdb"
    'Public ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.1.10\Public_Users\AtecCustomerDatabase\CustomersDatabase.accdb"
    'Public ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|IDMaker.accdb"
    'Public ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|ITTechPartsDatabase.accdb"
    'Public ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|TonerDatabase.accdb"
    'Public ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|ITMSDatabase.accdb"
    'Public ConnectionString As String = "Provider=SQLNCLI10;Server=;Database=SmartTrack 3.25;Uid=SA;Pwd=25hpw2k30304$"
    'Public Const ConnectionString As String = "Provider=SQLOLEDB.1;Data Source=192.168.1.3;Initial Catalog=SmartTrack 3.25;User ID=SA;Password=25hpw2k30304$"
    Public Const ConnectionString As String = "Provider=SQLOLEDB.1;Data Source=192.168.1.58\AXDB;Initial Catalog=MES_ATEC;User ID=sa;Password=p@ssw0rd"
    'Public ConnectionString As String = "Provider=SQLOLEDB.1;Data Source=192.168.101.85;Initial Catalog=MES_ATEC;User ID=sa;Password=dnhk0723$%"
    Public ConnectionString1 As String = "Provider=SQLOLEDB.1;Data Source=192.168.1.3;Initial Catalog=SmartTrack 3.25;User ID=sa;Password=25hpw2k30304$"
    'Public TEST_AXConnectionString As String = "Provider=SQLOLEDB.1;Data Source=MSDYNAMICS-DB\AXDB;Initial Catalog=TestDB;User ID=sa;Password=p@ssw0rd"
    Public TEST_AXConnectionString As String = "Provider=SQLOLEDB.1;Data Source=MSDYNAMICS-DB\AXDB;Initial Catalog=TestDB;User ID=sa;Password=p@ssw0rd"
    Public tempSQL As String
    Public CurrentUser As String
    Public tempPress As String
    Public Port As String = "LPT1"
    Public PM_By, userid, username, admin, eng, fname, HardwareID, HardwareStatus, HardwareName, HardwareNo, HardwareCode, TesterCode, HandlerCode, CabinetCode, LocationCode, Remarks, tempress, HardwareRemarks, LeadCount, PackageType, CorStatus, RequestCorelationID, CorelationID, Program As String

    Public Sub NumberOnly(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        'Allow Backspace (8), numeric keys (48 to 57), comma (44), 
        'decimal point (46)and dollar sign (36)
        Select Case KeyAscii
            Case 8, 48 To 57
                e.Handled = False  'Allow the key
            Case Else
                e.Handled = True   'Ignore the key
        End Select
    End Sub

    Public Sub FillComboBox(ByVal ComboBox As ComboBox, ByVal SQL_String As String)
        ComboBox.Items.Clear()
        Dim DS As DataSet
        DS = FillDataset(SQL_String)
        For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
            ComboBox.Items.Add(DS.Tables(0).Rows(i).Item(0).ToString())
        Next
    End Sub
    Public Sub FillComboBox_V2(ByVal ComboBox As ComboBox, ByVal SQL_String As String)
        ComboBox.Items.Clear()
        Dim DS As DataSet
        DS = FillDataset1(SQL_String)
        For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
            ComboBox.Items.Add(DS.Tables(0).Rows(i).Item(0).ToString())
        Next
    End Sub

    Public Sub FillToolStripComboBox(ByVal ComboBox As ToolStripComboBox, ByVal SQL_String As String)
        ComboBox.Items.Clear()
        Dim DS As DataSet
        DS = FillDataset(SQL_String)
        For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
            ComboBox.Items.Add(DS.Tables(0).Rows(i).Item(0).ToString())
        Next
    End Sub

    Public Sub IntelText(ByVal Combobox As ToolStripComboBox, ByVal SQL_String As String)
        Combobox.AutoCompleteCustomSource.Clear()
        Dim DS As DataSet
        DS = FillDataset(SQL_String)
        For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
            Combobox.AutoCompleteCustomSource.Add(DS.Tables(0).Rows(i).Item(0).ToString())
        Next
    End Sub

    Public Function GetIfExist(ByVal TableName As String, ByVal Column As String, ByVal Value As String)
        GetIfExist = True
        Dim DS As DataSet
        DS = FillDataset("select " & Column & " from " & TableName & " where " & Column & " = '" & Value & "'")
        If DS.Tables(0).Rows.Count <= 0 Then
            GetIfExist = False
        End If
    End Function
    Public Function GetIfExistV2(ByVal SQL As String) As Boolean
        GetIfExistV2 = True
        Dim DS As DataSet
        DS = FillDataset(SQL)
        If DS.Tables(0).Rows.Count <= 0 Then
            GetIfExistV2 = False
        End If
    End Function

    Public Function RowCounter(ByVal SQL As String) As Integer
        RowCounter = 0
        Dim DS As DataSet
        DS = FillDataset(SQL)
        If DS.Tables(0).Rows.Count > 0 Then
            RowCounter = DS.Tables(0).Rows.Count
        Else
            Return 0
        End If
    End Function
    Public Function RowCounter1(ByVal SQL As String) As Integer
        RowCounter1 = 0
        Dim DS As DataSet
        DS = FillDataset1(SQL, ConnectionString1)
        If DS.Tables(0).Rows.Count > 0 Then
            RowCounter1 = DS.Tables(0).Rows.Count
        Else
            Return 0
        End If
    End Function

    Public Function FillDataset(ByVal SQLQuery As String) As DataSet
        Dim Connection As New OleDb.OleDbConnection(ConnectionString)
        Dim DataAdapter As New OleDb.OleDbDataAdapter(SQLQuery, Connection)
        Dim DS As New DataSet
        Try
            DataAdapter.Fill(DS)
            Connection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Connection.Close()
        End Try
        Return DS
    End Function
    Public Function FillDataset(ByVal SQLQuery As String, ByVal ConnectionString As String, Optional ByVal TimeOutLimit As Integer = 60) As DataSet
        'DSError = ""
        Dim Connection As New OleDb.OleDbConnection(ConnectionString)
        Dim DataAdapter As New OleDb.OleDbDataAdapter(SQLQuery, ConnectionString)
        Dim DS As New DataSet
        Try
            DataAdapter.SelectCommand.CommandTimeout = TimeOutLimit
            DataAdapter.Fill(DS)
            Connection.Close()
        Catch ex As Exception
            'DSError = ex.Message
            MessageBox.Show(ex.Message)
            Connection.Close()
        End Try
        Return DS
    End Function

    Public Function FillDataset1(ByVal SQLQuery As String) As DataSet
        Dim Connection As New OleDb.OleDbConnection(ConnectionString1)
        Dim DataAdapter As New OleDb.OleDbDataAdapter(SQLQuery, Connection)
        Dim DS As New DataSet
        Try
            DataAdapter.Fill(DS)
            Connection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Connection.Close()
        End Try
        Return DS
    End Function
 
    Public Function FillDataset1(ByVal SQLQuery As String, ByVal ConnectionString As String, Optional ByVal TimeOutLimit As Integer = 60) As DataSet
        'DSError = ""
        Dim Connection As New OleDb.OleDbConnection(ConnectionString1)
        Dim DataAdapter As New OleDb.OleDbDataAdapter(SQLQuery, ConnectionString)
        Dim DS As New DataSet
        Try
            DataAdapter.SelectCommand.CommandTimeout = TimeOutLimit
            DataAdapter.Fill(DS)
            Connection.Close()
        Catch ex As Exception
            'DSError = ex.Message
            MessageBox.Show(ex.Message)
            Connection.Close()
        End Try
        Return DS
    End Function



    Public Sub connectDB_DataGrid(ByVal SQL As String, ByVal DataGridViewTemp As DataGridView)
        Dim myConnection As OleDbConnection = New OleDbConnection
        Try
            myConnection.ConnectionString = ConnectionString

            Dim da As OleDbDataAdapter = New OleDbDataAdapter(SQL, myConnection)

            Dim ds As DataSet = New DataSet
            Threading.Thread.Sleep(5000)
            da.Fill(ds, 0)

            DataGridViewTemp.DataSource = ds.Tables(0)

        Catch ex As OleDbException
            MessageBox.Show("Error " & ex.ToString, "Loading Records")
        Finally
            myConnection.Close()
        End Try

    End Sub

    Public Function FillDataGrid(ByVal datagrid As DataGridView, ByVal SQL_String As String) As Boolean
        datagrid.Rows.Clear()
        datagrid.Columns.Clear()

        'FillDataSet
        Dim DS As DataSet
        DS = FillDataset(SQL_String)

        'Check if Table is created
        If DS.Tables.Count > 0 Then
            If DS.Tables(0).Rows.Count > 0 Then
                'Insert Columns to DataGridView
                For Column As Integer = 0 To DS.Tables(0).Columns.Count - 1
                    datagrid.Columns.Add(DS.Tables(0).Columns(Column).ColumnName.ToString, DS.Tables(0).Columns(Column).ColumnName.ToString)
                    datagrid.Columns(Column).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Next

                'Insert data to DataGridView
                For Row As Integer = 0 To DS.Tables(0).Rows.Count - 1
                    datagrid.Rows.Add()
                    For Column As Integer = 0 To DS.Tables(0).Columns.Count - 1
                        datagrid.Item(Column, Row).Value = DS.Tables(0).Rows(Row).Item(Column).ToString
                    Next
                Next
            End If
        End If
    End Function

    Public Sub FillDataGridV3(ByVal datagrid As DataGridView, ByVal SQL_String As String)
        datagrid.DataSource = Nothing
        Dim DS As New DataSet
        DS = FillDataset(SQL_String)
        datagrid.DataSource = DS.Tables(0)
    End Sub

    'Public Function FillDataGridV2(ByVal datagrid As DataGridView, ByVal SQL_String As String) As Boolean
    '    Try
    '        FillDataGridV2 = True
    '        Dim DS As DataSet
    '        DS = FillDataset(SQL_String)
    '        datagrid.DataSource = DS.Tables(0)
    '        For Column As Integer = 0 To DS.Tables(0).Columns.Count - 1
    '            datagrid.Columns(Column).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
    '        Next
    '    Catch ex As Exception
    '        FillDataGridV2 = False
    '    End Try

    'End Function

    Public Function FillDataGridV2(ByVal datagrid As DataGridView, ByVal SQL_String As String, ByVal ConnectionString As String, Optional ByVal TimeOutLimit As Integer = 60) As Boolean
        Try
            FillDataGridV2 = True
            Dim DS As DataSet
            DS = FillDataset(SQL_String, ConnectionString, TimeOutLimit)
            datagrid.DataSource = DS.Tables(0)
            datagrid.AllowUserToAddRows = False
            datagrid.AllowUserToDeleteRows = False
            datagrid.AllowUserToOrderColumns = False
            datagrid.AllowUserToResizeRows = False
            datagrid.RowHeadersVisible = False
            datagrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing

            'For Column As Integer = 0 To DS.Tables(0).Columns.Count - 1
            '    datagrid.Columns(Column).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Next
        Catch ex As Exception
            FillDataGridV2 = False
        Finally
            datagrid.Refresh()
        End Try
    End Function
    Public Function SaveToDB(ByVal SQLQuery As String) As String
        SaveToDB = True
        Dim Connection As New OleDb.OleDbConnection(ConnectionString)
        Connection.Open()
        Try
            Dim Cmd As New OleDb.OleDbCommand(SQLQuery, Connection)
            Cmd.CommandType = CommandType.Text
            Cmd.ExecuteNonQuery()
            ' MessageBox.Show("Transaction Completed")
            Connection.Close()
        Catch ex As Exception
            SaveToDB = False
            MessageBox.Show(ex.ToString)
            Connection.Close()
        End Try
    End Function

    Public Sub CloseForm(ByVal _form As Form)
        'For i As Decimal = 1 To 0.1 Step -0.001
        '    _form.Height = _form.Height * i
        '    _form.Width = _form.Width * i
        'Next
    End Sub

    Public Sub MoneyKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal _TextBox As TextBox)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        'Allow Backspace (8), numeric keys (48 to 57), comma (44), 
        'decimal point (46)and dollar sign (36)

        Select Case KeyAscii
            Case 46
                If Not _TextBox.Text.Contains(".") Then
                    e.Handled = False  'Allow the key
                Else
                    e.Handled = True   'Ignore the key
                End If
            Case 48 To 57
                'error
                e.Handled = False  'Allow the key
                If _TextBox.Text.Contains(".") Then
                    Dim start As Integer = InStr(_TextBox.Text, ".")
                    Dim str As String = Mid(_TextBox.Text, start + 1, _TextBox.Text.Length - start)
                    If str.Length < 2 Then
                        e.Handled = False  'Allow the key
                    Else
                        e.Handled = True   'Ignore the key
                    End If
                    Dim test As String = ""
                Else
                    e.Handled = False  'Allow the key
                End If
            Case 8
                e.Handled = False  'Allow the key
            Case Else
                e.Handled = True   'Ignore the key
        End Select
    End Sub

    Public Sub LoadIntellisense(ByVal TBox As TextBox, ByVal SQLQuery As String)
        Dim sqlConnection As New OleDb.OleDbConnection(ConnectionString)
        Dim DataAdapter As New OleDb.OleDbDataAdapter(SQLQuery, sqlConnection)
        Dim DS As New DataSet
        DataAdapter.Fill(DS)
        TBox.AutoCompleteCustomSource.Clear()
        TBox.AutoCompleteMode = AutoCompleteMode.Suggest
        TBox.AutoCompleteSource = AutoCompleteSource.CustomSource
        For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
            TBox.AutoCompleteCustomSource.Add(DS.Tables(0).Rows(i).Item(0).ToString)
        Next

        'TBox.AutoCompleteCustomSource = DS.Tables(0)

    End Sub

    Public Sub FirstCharacterCap(ByRef TextBox)
        If TextBox.text.ToString.Length = 1 Then
            TextBox.Text = TextBox.Text.ToUpper
            TextBox.SelectionStart = 2
        End If
    End Sub

    Private Function GetNextDate(ByVal day As DayOfWeek) As DateTime
        Dim now As DateTime = DateTime.Today
        Dim today As Integer = CInt(now.DayOfWeek)
        Dim find As Integer = CInt(day)

        Dim delta As Integer = find - today
        If delta > 0 Then
            Return now.AddDays(delta)
        Else
            Return now.AddDays(7 - delta)
        End If
    End Function
    Public Sub LoadIntellisense(ByVal TBox As ToolStripComboBox, ByVal SQLQuery As String)
        Dim sqlConnection As New OleDb.OleDbConnection(ConnectionString)
        Dim DataAdapter As New OleDb.OleDbDataAdapter(SQLQuery, sqlConnection)
        Dim DS As New DataSet
        DataAdapter.Fill(DS)
        TBox.AutoCompleteCustomSource.Clear()
        TBox.AutoCompleteMode = AutoCompleteMode.Suggest
        TBox.AutoCompleteSource = AutoCompleteSource.CustomSource
        For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
            TBox.AutoCompleteCustomSource.Add(DS.Tables(0).Rows(i).Item(0).ToString)
        Next
    End Sub
    Private Function GetMonth(ByVal DateValue As DateTime) As Integer
        Dim now As DateTime = DateTime.Today
        Dim month As Integer = CInt(now.Month)
        Dim find As Integer = CInt(month)
        Dim delta As Integer = find - month
    End Function

    Public Sub ClearTextBox(ByVal root As Control)
        For Each ctrl As Control In root.Controls
            ClearTextBox(ctrl)
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Text = ""
            End If
        Next ctrl
    End Sub

    Public Sub TextBoxNotApplicable(ByVal root As Control)

        For Each ctrl As Control In root.Controls
            TextBoxNotApplicable(ctrl)
            If TypeOf ctrl Is TextBox And ctrl.Text = "" Then
                CType(ctrl, TextBox).Text = "N/A"
            End If

        Next ctrl

    End Sub

    Public Sub ComboboxNotApplicable(ByVal root As Control)

        For Each ctrl As Control In root.Controls
            ComboboxNotApplicable(ctrl)
            If TypeOf ctrl Is ComboBox And ctrl.Text = "" Then
                CType(ctrl, ComboBox).Text = "N/A"
            End If
        Next ctrl

    End Sub

    Public Sub EnableTextBox(ByVal root As Control, ByVal t As Boolean)

        For Each ctrl As Control In root.Controls
            EnableTextBox(ctrl, t)
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Enabled = t
            End If
        Next ctrl
    End Sub

    Public Sub EnablecomboBox(ByVal root As Control, ByVal t As Boolean)
        For Each ctrl As Control In root.Controls
            EnablecomboBox(ctrl, t)
            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).Enabled = t
            End If
        Next ctrl
    End Sub

    Public Sub EmptyAllComboBox(ByVal root As Control)
        For Each ctrl As Control In root.Controls
            EmptyAllComboBox(ctrl)
            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).Text = Nothing
            End If
        Next ctrl
    End Sub

    Public Function SaveDB(ByVal strSQL As String)
        Dim cn As OleDbConnection
        cn = New OleDbConnection

        Try
            With cn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = ConnectionString
                .Open()
            End With

            Dim cmd As OleDbCommand = New OleDbCommand(strSQL, cn)

            cmd.ExecuteNonQuery()

            Return True
        Catch ex As OleDbException
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ATEC")
            Return False
        Finally
            cn.Close()
        End Try
    End Function
    Public Function ConvertToLetter(ByRef iCol As Integer) As String

        'Dim Reminder_Part As Integer = iCol Mod 26
        'Dim Integer_Part As Integer = Int(iCol / 26)
        ''Dim ConvertToLetter As String = ""

        'If Integer_Part = 0 Then
        '    ConvertToLetter = Chr(Reminder_Part + 64)
        'ElseIf Integer_Part > 0 And Reminder_Part <> 0 Then
        '    ConvertToLetter = Chr(Integer_Part + 64) + Chr(Reminder_Part + 64)
        'ElseIf Integer_Part > 0 And Reminder_Part = 0 Then
        '    ConvertToLetter = Chr(Integer_Part * 26) + Chr(Reminder_Part + 64)
        'End If

        Dim iAlpha As Integer
        Dim iRemainder As Integer
        iAlpha = Int(iCol / 27)
        iRemainder = iCol - (iAlpha * 26)
        If iAlpha > 0 Then
            ConvertToLetter = Chr(iAlpha + 64)
        End If
        If iRemainder > 0 Then
            ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
        End If

        'Return ConvertToLetter
    End Function
End Module


