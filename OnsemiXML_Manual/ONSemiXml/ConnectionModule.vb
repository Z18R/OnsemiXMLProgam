Imports System.Data.SqlClient
Imports System.Data.OleDb

Module ConnectionModule

    Public Sub ConnectToSmartTrack(ByRef conn As SqlConnection)
        'Try
        Dim strConn As String = "Initial Catalog = SmartTrack 3.25; Data Source = 192.168.1.3;  User ID = sa; Password = 25hpw2k30304$; Connection TimeOut = 99999999"
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.ConnectionString = strConn
        conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Public Sub ConnectToMPS(ByRef conn As SqlConnection)
        'Try
        Dim strConn As String = "Initial Catalog = MPS; Data Source = 192.168.1.15;  User ID = sa; Password = enola845&*; Connection TimeOut = 99999999"
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.ConnectionString = strConn
        conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Public Sub ConnectToMPXData(ByRef conn As SqlConnection)
        ' Try
        Dim strConn As String = "Initial Catalog = MPXData; Data Source = 192.168.1.15;  User ID = sa; Password = enola845&*; Connection TimeOut = 99999999"
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.ConnectionString = strConn
        conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub
    Public Sub ConnectToMSDYNAMICS(ByRef conn As SqlConnection)
        ' Try
        Dim strConn As String = "Initial Catalog =AX2009DB; Data Source = 192.168.1.58\AXDB;  User ID = sa; Password = p@ssw0rd; Connection TimeOut = 99999999"
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.ConnectionString = strConn
        conn.Open()
        ' Catch ex As Exception
        'MessageBox.Show(ex.Message, "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ' End Try
    End Sub

    Public Sub ConnectToMES_ATEC(ByRef conn As SqlConnection)
        'Try
        Dim strConn As String = "Initial Catalog =MES_ATEC; Data Source = 192.168.1.58\AXDB;  User ID = sa; Password = p@ssw0rd; Connection TimeOut = 99999999"
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.ConnectionString = strConn
        conn.Open()
        'Catch ex As Exception
        ' MessageBox.Show(ex.Message, "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Public Sub ConnectToMSAccessData(ByRef conn As OleDbConnection)
        Try
            Dim strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\MAPICS.accdb;Jet OLEDB:Database Password=MAPICS5.0"
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.ConnectionString = strConn
            conn.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub CloseConnection(ByRef conn As SqlConnection)
        If conn.State = ConnectionState.Open Then
            conn.Close()
            conn = Nothing
        End If
    End Sub

    Public Sub CloseConnection(ByRef conn As OleDbConnection)
        If conn.State = ConnectionState.Open Then
            conn.Close()
            conn = Nothing
        End If
    End Sub


    Public Function ExecuteQuery(ByVal strSQL As String, ByVal conn As SqlConnection, ByRef dr As SqlDataReader) As Boolean
        ExecuteQuery = False

        Dim cmd As New SqlCommand(strSQL, conn)
        Try
            dr = cmd.ExecuteReader()
            ExecuteQuery = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Failed to Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ExecuteQuery = False
        End Try
    End Function

    Public Function ExecuteNonQuery(ByVal strSQL As String, ByVal conn As SqlConnection) As Boolean
        ExecuteNonQuery = False

        Dim cmd As New SqlCommand(strSQL, conn)
        Try
            cmd.ExecuteNonQuery()
            cmd = Nothing
            ExecuteNonQuery = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Failed to Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ExecuteNonQuery = False
        End Try
    End Function

    Public Function ExecuteQuery(ByVal strSQL As String, ByVal conn As SqlConnection, ByRef ds As DataSet, ByVal strName As String) As Boolean
        ExecuteQuery = True
        Try
            Dim cmd As New SqlCommand(strSQL, conn)
            cmd.CommandTimeout = 99999999
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(ds, strName)
            da = Nothing
            cmd = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Failed to Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ExecuteQuery = False
        End Try
    End Function


    Public Function ExecuteQuery(ByVal strSQL As String, ByVal conn As OleDbConnection, ByRef dr As OleDbDataReader) As Boolean
        ExecuteQuery = False

        Dim cmd As New OleDbCommand(strSQL, conn)
        Try
            dr = cmd.ExecuteReader()
            ExecuteQuery = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Failed to Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ExecuteQuery = False
        End Try
    End Function

    Public Function ExecuteNonQuery(ByVal strSQL As String, ByVal conn As OleDbConnection) As Boolean
        ExecuteNonQuery = False

        Dim cmd As New OleDbCommand(strSQL, conn)
        Try
            cmd.ExecuteNonQuery()
            cmd = Nothing
            ExecuteNonQuery = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Failed to Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ExecuteNonQuery = False
        End Try
    End Function

    Public Function ExecuteQuery(ByVal strSQL As String, ByVal conn As OleDbConnection, ByRef ds As DataSet, ByVal dsName As String) As Boolean
        ExecuteQuery = True
        Try
            Dim cmd As New OleDbCommand(strSQL, conn)
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, dsName)
            da = Nothing
            cmd = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Failed to Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ExecuteQuery = False
        End Try
    End Function

End Module
