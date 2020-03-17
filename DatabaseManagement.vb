Imports System.Xml.Serialization
'Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System.Data
Imports MySql.Data.MySqlClient

Public Class DatabaseManagement

    Public Shared Function ConvertToDBDateFormat(_date As Date, Optional isDatetime As Boolean = False)
        If isDatetime Then
            Return _date.ToString("yyyy-MM-dd HH:mm:ss")
        Else : Return _date.ToString("yyyy-MM-dd")
        End If
    End Function

#Region "MySQL"
    Public Const MySQLConfigFileExtension = ".mysql.config.xml"
    <XmlRoot("MysqlConfiguration")> Public Class MysqlConfiguration
        Implements IDisposable

        <XmlIgnore> Public Connection As New MySqlConnection
        Public Server As String, DatabaseName As String, UserID As String, Password As String, Port As String

        Public ReadOnly Property HasSelectedDatabase As Boolean
            Get
                Return Not DatabaseName = ""
            End Get
        End Property

        Public WriteOnly Property OpenConnection As Boolean
            Set(value As Boolean)
                If value Then
                    If Not Connection.State = ConnectionState.Open Then Connection.Open()
                Else
                    If Connection.State = ConnectionState.Open Then Connection.Close()
                End If
            End Set
        End Property

        Public Function CloneConnection() As MySqlConnection
            Dim con As MySqlConnection = Nothing
            con = Connection.Clone
            Return con
        End Function

        Public Function SetupConnection(Optional _connectionString As String = Nothing) As Boolean
            Try
                Dim connectionString As String = "Server={0};Uid={1};Pwd={2};port={3}{4};Convert Zero Datetime=True;command Timeout=20000;SslMode=None"
                If _connectionString IsNot Nothing And _connectionString IsNot "" Then
                    connectionString = _connectionString
                End If

                Connection = New MySqlConnection
                Connection.ConnectionString = String.Format(connectionString _
                                                             , Server, UserID, Password, Port, IIf(DatabaseName = "", "", ";database=" & DatabaseName))
                OpenConnection = True
                Return True
            Catch ex As Exception
                MsgBox(ex.Message)
                Return False
            End Try
        End Function

        Public Sub SetupConnection(ByRef con As MySqlConnection)
            Try
                con = New MySqlConnection
                con.ConnectionString = String.Format("Server={0};Uid={1};Pwd={2};port={3}{4};Convert Zero Datetime=True;command Timeout=20000;SslMode=None;" _
                                                     , Server, UserID, Password, Port, IIf(DatabaseName = "", "", ";database=" & DatabaseName))
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Public Sub SetupConnection(_databaseName As String, ByRef con As MySqlConnection)
            Try
                con = New MySqlConnection
                con.ConnectionString = String.Format("Server={0};Uid={1};Pwd={2};port={3};database={4};Convert Zero Datetime=True;command Timeout=20000;SslMode=None;" _
                                                     , Server, UserID, Password, Port, _databaseName)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Public Sub Close()
            OpenConnection = False
            Connection.Dispose()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub

        Public Function ExecuteDataReader(sql As String) As MySqlDataReader
            OpenConnection = True
            Return New MySqlCommand(sql, Connection).ExecuteReader
        End Function

        Public Shared Function ExecuteDataReader(sql As String, _con As MySqlConnection) As MySqlDataReader
            Return New MySqlCommand(sql, _con).ExecuteReader
        End Function

        Public Sub ExecuteQuery(ByVal sql As String)
            Try
                OpenConnection = True
                Dim cmd As MySqlCommand
                cmd = New MySqlCommand(sql, Connection)
                cmd.ExecuteNonQuery()
                OpenConnection = False
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Public Shared Sub ExecuteQuery(ByVal sql As String, _con As MySqlConnection)
            Dim cmd As MySqlCommand
            cmd = New MySqlCommand(sql, _con)
            cmd.ExecuteNonQuery()
        End Sub

        Public Function ExecuteScalar(ByVal sql As String) As String
            Dim res As String = ""
            Dim cmd As MySqlCommand
            cmd = New MySqlCommand(sql, Connection)
            res = cmd.ExecuteScalar()
            Return res
        End Function

        Public Shared Sub AlterTablename(tablename As String, newTablename As String, con As MySqlConnection)
            MysqlConfiguration.ExecuteDataReader("Alter Table `" & tablename & "` Rename To `" & newTablename & "`", con)
        End Sub

        Public Sub AlterTablename(tablename As String, newTablename As String)
            MysqlConfiguration.AlterTablename(tablename, newTablename, Connection)
        End Sub

        Public Function ToDT(sql As String) As DataTable
            Dim dt As DataTable = New DataTable()
            Dim dAdapter As New MySqlDataAdapter(sql, Connection)
            dAdapter.Fill(dt)
            Return dt
        End Function

        Public Shared Function ToDT(sql As String, _con As MySqlConnection) As DataTable
            Dim dt As DataTable = New DataTable()
            Dim dAdapter As New MySqlDataAdapter(sql, _con)
            dAdapter.Fill(dt)
            Return dt
        End Function

        Public Function CreateSchema(ByVal dbname As String) As String
            Dim res As String = ""
            res = "CREATE SCHEMA " & dbname & ";"
            Return res
        End Function

        Public Sub TryCreateTable(ByVal tbl As String, ByVal flds As String(), Optional overwrite As Boolean = False)
            If CheckTable(tbl) Then
                If Not overwrite Then Exit Sub
                ExecuteQuery("DROP TABLE `" & tbl & "`")
            End If

            CreateTable(tbl, flds)
        End Sub

        Public Sub CreateTable(ByVal tbl As String, ByVal flds As String())
            MysqlConfiguration.CreateTable(tbl, flds, Connection)
        End Sub
        Public Shared Sub CreateTable(ByVal tbl As String, ByVal flds As String(), ByVal con As MySqlConnection)
            Dim qry As String = String.Format("CREATE TABLE `{0}`(", tbl)
            For i As Integer = 0 To flds.Length - 1
                qry &= IIf(i = 0, flds(i), "," & flds(i))
            Next
            qry &= ")"

            MysqlConfiguration.ExecuteQuery(qry, con)
        End Sub

        'SQL COMMANDS
        Public Function CheckTable(tbl As String) As Boolean
            Return MysqlConfiguration.CheckTable(DatabaseName, tbl, Connection)
        End Function
        Public Shared Function CheckTable(DatabaseName As String, tbl As String, _con As MySqlConnection) As Boolean
            Return MysqlConfiguration.GetTables(DatabaseName, _con).Contains(tbl.ToLower)
        End Function

        Public Shared Function GetTables(DatabaseName As String, _con As MySqlConnection) As List(Of String)
            Dim lst As New List(Of String)
            Using rdr As MySqlDataReader = MysqlConfiguration.ExecuteDataReader("SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_SCHEMA='" & DatabaseName & "';", _con)
                While rdr.Read
                    lst.Add(rdr.Item(0).ToString.ToLower)
                End While
            End Using
            Return lst
        End Function
        Public Function GetTables() As List(Of String)
            Return MysqlConfiguration.GetTables(DatabaseName, Connection)
        End Function
        'Public Shared Sub CreateTable(ByVal tbl As String, ByVal flds As String(), ByVal con As MySqlConnection)
        '    Dim qry As String = String.Format("CREATE TABLE `{0}`(", tbl)
        '    For i As Integer = 0 To flds.Length - 1
        '        qry &= IIf(i = 0, flds(i), "," & flds(i))
        '    Next
        '    qry &= ")"

        '    MysqlConfiguration.ExecuteNonQuery(qry, con)
        'End Sub

        Public Shared Function CreateTableGenerator(ByVal table As String, ByVal fields As String(), ByVal dataTypes As String(), ByVal withPrimary As Boolean, Optional ByVal type As String = "") As String
            Dim res As String = ""
            res = "CREATE TABLE " & table & " ("

            If withPrimary Then
                res += fields(0) & " " & type & ", "

                For i As Integer = 1 To fields.Length - 1
                    res += fields(i) & " " & dataTypes(i) & ", "
                Next

                res += "PRIMARY KEY (" & fields(0) & "));"
            Else
                res += fields(0) & " " & dataTypes(0) & ","

                For i As Integer = 1 To fields.Length - 1
                    If i = fields.Length - 1 Then
                        res += fields(i) & " " & dataTypes(i) & ");"
                    Else
                        res += fields(i) & " " & dataTypes(i) & ","
                    End If
                Next

            End If

            Return res
        End Function

        Public Sub Insert(ByVal schema As String, ByVal tbl As String, ByVal fld As String(), ByVal val As Object())
            MysqlConfiguration.Insert(schema, tbl, fld, val, Connection)
        End Sub

        Public Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object())
            MysqlConfiguration.Insert(DatabaseName, tbl, fld, val, Connection)
        End Sub

        Public Shared Sub Insert(schema As String, ByVal table As String, ByVal fields As String(), ByVal values As Object(), con As MySqlConnection)
            Dim qry As String = String.Format("INSERT INTO `{0}` (", table)
            Dim valtype As String = ""

            For i As Integer = 0 To fields.Length - 1
                Dim f As String = fields(i)
                If f = fields(0) Then
                    qry &= String.Format("`{0}`", f)
                Else
                    qry &= String.Format(",`{0}`", f)
                End If
            Next

            qry &= ") VALUES("

            For i As Integer = 0 To values.Length - 1
                Dim v = values(i)
                valtype = TypeName(v)
                Select Case valtype
                    Case "String"
                        qry &= String.Format("'{0}',", v)
                    Case "Date"
                        qry &= String.Format("'{0}',", Date.Parse(v).ToString("yyyy-MM-dd HH:mm:ss"))
                    Case Else
                        qry &= String.Format("{0},", v)
                End Select
            Next
           qry = System.Text.RegularExpressions.Regex.Replace(qry, ",+$", "")
            qry &= ")"


            MysqlConfiguration.ExecuteQuery(qry, con)
        End Sub

        Public Shared Function GenerateSearch(ByVal columns() As String, ByVal table As String, ByVal ID As String, ByVal IDval As String) As String
            Dim sql As String = "Select "

            For i As Integer = 0 To columns.Length - 1
                If i = columns.Length - 1 Then
                    sql += "[" & columns(i) & "] "
                Else
                    sql += "[" & columns(i) & "], "
                End If
            Next

            sql += " from " & table & " where " & ID & " = '" & IDval & "'"

            Return sql.Trim
        End Function

        Public Sub Update(schema As String, ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As SQLCondition())
            MysqlConfiguration.Update(schema, tbl, fld, val, condition, Connection)
        End Sub
        Public Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As SQLCondition())
            MysqlConfiguration.Update(DatabaseName, tbl, fld, val, condition, Connection)
        End Sub
        Public Shared Sub Update(schema As String, ByVal table As String, ByVal fields As String(), ByVal values As Object(), ByVal condition As SQLCondition(), con As MySqlConnection)
            Dim qry As String = String.Format("UPDATE {0} SET ", table)
            Dim valtype As String = ""

            If fields.Length = values.Length Then
                For i As Integer = 0 To fields.Length - 1
                    Dim v = values(i)
                    valtype = TypeName(v)
                    Select Case valtype
                        Case "String"
                            qry &= String.Format("`{0}`='{1}',", fields(i), v)
                        Case "Date"
                            qry &= String.Format("`{0}`='{1}',", fields(i), Date.Parse(v).ToString("yyyy-MM-dd HH:mm:ss"))
                        Case Else
                            qry &= String.Format("`{0}`={1},", fields(i), v)
                    End Select
                Next
            End If

            qry = System.Text.RegularExpressions.Regex.Replace(qry, ",+$", "")
            If Not condition Is Nothing Then
                qry &= " WHERE"
                For i As Integer = 0 To condition.Length - 1
                    qry &= condition(i).ToString
                Next
            End If


            MysqlConfiguration.ExecuteQuery(qry, con)
        End Sub

        Public Shared Function Update(ByVal table As String, ByVal fields As String(), ByVal values As String(), ByVal ID As String, ByVal IDVal As String)
            Dim sql As String = ""

            sql = "UPDATE " & table & " SET "
            For i As Integer = 1 To values.Length - 1
                If values(i) Is Nothing Then values(i) = ""
                If i < values.Length - 1 Then
                    sql += fields(i) & "='" & values(i).Replace("'", "''") & "',"
                Else
                    sql += fields(i) & "='" & values(i).Replace("'", "''") & "'"
                End If
            Next
            sql += " WHERE " & ID & " = '" & IDVal & "'"


            Return sql
        End Function

        Public Function SchemaExist(ByVal dbname As String) As Boolean
            Dim cmd As MySqlCommand = New MySqlCommand("SELECT IF(EXISTS (SELECT SCHEMA_NAME " & _
                 "FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & dbname & "'), 'Y','N')", Connection)
            'cmd.Parameters.AddWithValue("@DbName", dbname)

            Dim exists As String = cmd.ExecuteScalar().ToString()
            Return If(exists = "Y", True, False)
        End Function

        Public Shared Function SchemaExist(ByVal dbname As String, con As MySqlConnection) As Boolean
            Dim cmd As MySqlCommand = New MySqlCommand("SELECT IF(EXISTS (SELECT SCHEMA_NAME " & _
                 "FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & dbname & "'), 'Y','N')", con)
            'cmd.Parameters.AddWithValue("@DbName", dbname)

            Dim exists As String = cmd.ExecuteScalar().ToString()
            Return If(exists = "Y", True, False)
        End Function

        Public Function TableExist(ByVal table As String) As Boolean
            Dim restrictions(4) As String
            restrictions(2) = table

            Dim dbTbl As DataTable = Connection.GetSchema("Tables", restrictions)
            If dbTbl.Rows.Count = 0 Then
                Return False
            Else
                Return True
            End If
        End Function

#Region "Other Methods"
        Public Shared Function OpenEditor(_config As MysqlConfiguration, Optional appLocation As String = "", Optional appName As String = "Local") As MysqlConfiguration
            Dim filePath As String = Path.Combine(appLocation, appName & MySQLConfigFileExtension)
            Dim newConnection As MysqlConfiguration
            Dim settingsEditor As New SEAN.MySQL_Configuration(_config, appName, filePath)
            newConnection = settingsEditor.Config
            settingsEditor.ShowDialog()
            settingsEditor.Dispose()
            Return newConnection
        End Function

        Public Shared Function StartDefaultSetup(appLocation As String, Optional appName As String = "Local")
            Dim filePath As String = Path.Combine(appLocation, appName & MySQLConfigFileExtension)
            If File.Exists(filePath) Then
                Return ConfigurationStoring.XmlSerialization.ReadFromFile(filePath, New MysqlConfiguration)
            Else
                Return OpenEditor(New MysqlConfiguration, appLocation, appName)
            End If
        End Function
#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    Connection.Dispose()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class


#End Region












    '#Region "SQLite"
    '    Public Const SQLiteConfigFileExtension = ".sqlite.config.xml"
    '    <XmlRoot("SQLiteConfiguration")> Public Class SQLiteConfiguration
    '        <XmlIgnore> Public Connection As New SQLite.SQLiteConnection
    '        Public DBPath As String
    '        Public Password As String

    '        Sub New(_dbpath As String, Optional _password As String = "", Optional openNow As Boolean = False)
    '            DBPath = _dbpath
    '            Password = _password
    '            If openNow Then
    '                Open()
    '            End If
    '        End Sub

    '        Public Function CloneConnection() As SQLite.SQLiteConnection
    '            Dim con As SQLite.SQLiteConnection = Nothing
    '            con = Connection.Clone
    '            Return con
    '        End Function

    '        Public Sub Open()
    '            SQLiteConfiguration.Open(DBPath, Connection, Password)
    '        End Sub

    '        Public Shared Sub Open(_dbpath As String, ByRef _con As SQLite.SQLiteConnection, Optional _password As String = "")
    '            Try
    '                _con = New SQLite.SQLiteConnection("Data Source=""" & _dbpath & """;Password=" & _password & "")
    '                _con.Open()
    '            Catch ex As Exception
    '                MessageBox.Show("Can't connect right now, Please try again.", "Error: SQL Database Connection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '            End Try
    '        End Sub

    '        Public Sub Close()
    '            Connection.Close()
    '            Connection.Dispose()
    '        End Sub

    '        Public Sub CreateTable(ByVal tbl As String, ByVal flds As String())
    '            SQLiteConfiguration.CreateTable(tbl, flds, Connection)
    '        End Sub
    '        Public Shared Sub CreateTable(ByVal tbl As String, ByVal flds As String(), ByVal con As SQLite.SQLiteConnection)
    '            Dim qry As String = String.Format("CREATE TABLE `{0}`(", tbl)
    '            For i As Integer = 0 To flds.Length - 1
    '                qry &= IIf(i = 0, flds(i), "," & flds(i))
    '            Next
    '            qry &= ")"

    '            SQLiteConfiguration.ExecuteQuery(qry, con)
    '        End Sub

    '        Public Function ExecuteDataReader(query As String) As SQLite.SQLiteDataReader
    '            Return SQLiteConfiguration.ExecuteDataReader(query, Connection)
    '        End Function

    '        Public Shared Function ExecuteDataReader(query As String, _con As SQLite.SQLiteConnection) As SQLite.SQLiteDataReader
    '            Return New SQLite.SQLiteCommand(query, _con).ExecuteReader
    '        End Function

    '        Public Function ExecuteQuery(query As String) As Boolean
    '            Return SQLiteConfiguration.ExecuteQuery(query, Connection)
    '        End Function

    '        Public Shared Function ExecuteQuery(query As String, _con As SQLite.SQLiteConnection) As Boolean
    '            Try
    '                Dim command As New SQLite.SQLiteCommand(query, _con)
    '                command.ExecuteNonQuery()
    '                command.Dispose()
    '                Return True
    '            Catch ex As Exception
    '                MsgBox(ex.Message & " ExecuteQuery")
    '                Return False
    '            End Try
    '        End Function

    '        Public Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object())
    '            SQLiteConfiguration.Insert(tbl, fld, val, Connection)
    '        End Sub
    '        Public Shared Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), con As SQLite.SQLiteConnection)
    '            Dim qry As String = String.Format("INSERT INTO `{0}` (", tbl)
    '            Dim valtype As String = ""

    '            For i As Integer = 0 To fld.Length - 1
    '                Dim f As String = fld(i)
    '                If f = fld(0) Then
    '                    qry &= String.Format("`{0}`", f)
    '                Else
    '                    qry &= String.Format(",`{0}`", f)
    '                End If
    '            Next

    '            qry &= ") VALUES("

    '            For i As Integer = 0 To val.Length - 1
    '                Dim v = val(i)
    '                valtype = TypeName(v)
    '                If i = 0 Then
    '                    If valtype = "String" Then
    '                        qry &= String.Format("'{0}'", v)
    '                    Else
    '                        qry &= String.Format("{0}", v)
    '                    End If
    '                Else
    '                    If valtype = "String" Then
    '                        qry &= String.Format(",'{0}'", v)
    '                    Else
    '                        qry &= String.Format(",{0}", v)
    '                    End If
    '                End If
    '            Next
    '            qry &= ")"

    '            SQLiteConfiguration.ExecuteQuery(qry, con)
    '        End Sub
    '        Public Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object())
    '            SQLiteConfiguration.Update(tbl, fld, val, condition, Connection)
    '        End Sub
    '        Public Shared Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object(), ByVal con As SQLite.SQLiteConnection)
    '            Dim qry As String = String.Format("UPDATE {0} SET ", tbl)
    '            Dim valtype As String = ""

    '            If fld.Length = val.Length Then
    '                For f As Integer = 0 To fld.GetUpperBound(0)
    '                    valtype = TypeName(val(f))
    '                    If f = 0 Then
    '                        If valtype = "String" Then
    '                            qry &= String.Format("[{0}]='{1}'", fld(f), val(f))
    '                        Else
    '                            qry &= String.Format("[{0}]={1}", fld(f), val(f))
    '                        End If
    '                    Else
    '                        If valtype = "String" Then
    '                            qry &= String.Format(",[{0}]='{1}'", fld(f), val(f))
    '                        Else
    '                            qry &= String.Format(",[{0}]={1}", fld(f), val(f))
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            If Not condition Is Nothing Then
    '                If TypeName(condition(1)) = "String" Then
    '                    qry &= String.Format(" WHERE {0} = '{1}'", condition(0), condition(1))
    '                Else
    '                    qry &= String.Format(" WHERE {0} = {1}", condition(0), condition(1))
    '                End If
    '            End If

    '            SQLiteConfiguration.ExecuteQuery(qry, con)
    '        End Sub

    '        Public Function ToDT(qry As String) As DataTable
    '            Return SQLiteConfiguration.ToDT(qry, Connection)
    '        End Function
    '        Public Shared Function ToDT(qry As String, _con As SQLite.SQLiteConnection) As DataTable
    '            Try
    '                Dim dt As New DataTable
    '                Dim command As New SQLite.SQLiteDataAdapter(qry, _con)
    '                command.Fill(dt)
    '                command.Dispose()
    '                Return dt
    '            Catch ex As Exception
    '                MsgBox(ex.Message & " ExecuteQuery")
    '                Return Nothing
    '            End Try
    '        End Function

    '        Public Function CheckTable(tbl As String) As Boolean
    '            Return SQLiteConfiguration.CheckTable(tbl, Connection)
    '        End Function
    '        Public Function GetTables() As List(Of String)
    '            Return SQLiteConfiguration.GetTables(Connection)
    '        End Function
    '        Public Shared Function CheckTable(tbl As String, _con As SQLite.SQLiteConnection) As Boolean
    '            Return SQLiteConfiguration.GetTables(_con).Contains(tbl)
    '        End Function
    '        Public Shared Function GetTables(_con As SQLite.SQLiteConnection) As List(Of String)
    '            Dim lst As New List(Of String)
    '            Using rdr As SQLite.SQLiteDataReader = SQLiteConfiguration.ExecuteDataReader("SELECT `name` FROM sqlite_master WHERE type='table';", _con)
    '                While rdr.Read
    '                    lst.Add(rdr.Item(0))
    '                End While
    '            End Using
    '            Return lst
    '        End Function
    '    End Class
    '#End Region

#Region "SQLite"
    Public Const SQLiteConfigFileExtension = ".sqlite.config.xml"
    <XmlRoot("SQLiteConfiguration")> Public Class SQLiteConfiguration
        <XmlIgnore> Public Connection As New SQLite.SQLiteConnection
        Public Password As String
        Public DatabaseDirectory As String
        Public DatabaseName As String
        Public ReadOnly Property DatabaseFullPath As String
            Get
                Return Path.Combine(DatabaseDirectory, DatabaseName) & ".db"
            End Get
        End Property

        Sub New()

        End Sub

        Sub New(_databaseDirectory As String, _databaseName As String, Optional _password As String = "", Optional openNow As Boolean = False)
            DatabaseDirectory = _DatabaseDirectory
            DatabaseName = _databaseName
            Password = _password
            If openNow Then
                Open()
            End If
        End Sub

        Public Function CloneConnection() As SQLite.SQLiteConnection
            Dim con As SQLite.SQLiteConnection = Nothing
            con = Connection.Clone
            Return con
        End Function

        Public Function Open(Optional addOns As String = "") As Boolean
            Return SQLiteConfiguration.Open(DatabaseFullPath, Connection, Password, addOns)
        End Function

        Public Shared Function Open(_dbpath As String, ByRef _con As SQLite.SQLiteConnection, Optional _password As String = "", Optional addOns As String = "") As Boolean
            Try
                _con = New SQLite.SQLiteConnection(String.Format("Data Source={0};Password={1}{2}", _dbpath, _password, addOns))
                _con.Open()
                Return True
            Catch ex As Exception
                MessageBox.Show("Can't connect right now, Please try again.", "Error: SQL Database Connection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
            Return False
        End Function

        Public Sub Close()
            Connection.Close()
            Connection.Dispose()
        End Sub

#Region "Other Methods"
        Public Shared Function OpenEditor(_config As SQLiteConfiguration, Optional _connectionName As String = "Local", Optional _configFilePath As String = "") As SQLiteConfiguration
            Dim filePath As String = Path.Combine(_configFilePath, _connectionName & SQLiteConfigFileExtension)
            Dim newConnection As SQLiteConfiguration
            Dim settingsEditor As New SEAN.SQLite_Configuration(_config, _connectionName, filePath)
            newConnection = settingsEditor.Config
            settingsEditor.ShowDialog()
            settingsEditor.Dispose()
            Return newConnection
        End Function

        Public Shared Function StartDefaultSetup(appLocation As String, Optional appName As String = "Local")
            Dim filePath As String = Path.Combine(appLocation, appName & SQLiteConfigFileExtension)
            If File.Exists(filePath) Then
                Return ConfigurationStoring.XmlSerialization.ReadFromFile(filePath, New SQLiteConfiguration)
            Else
                Return OpenEditor(New SQLiteConfiguration, appName, appLocation)
            End If
        End Function
#End Region


        Public Shared Function CreateSchema(filename As String)
            Try
                SQLite.SQLiteConnection.CreateFile(filename)
                Return True
            Catch : Return False
            End Try
        End Function

        Public Sub TryCreateTable(ByVal tbl As String, ByVal flds As String(), Optional overwrite As Boolean = False)
            If CheckTable(tbl) Then
                If Not overwrite Then Exit Sub
                ExecuteQuery("DROP TABLE `" & tbl & "`")
            End If

            CreateTable(tbl, flds)
        End Sub
        Public Sub CreateTable(ByVal tbl As String, ByVal flds As String())
            SQLiteConfiguration.CreateTable(tbl, flds, Connection)
        End Sub
        Public Shared Sub CreateTable(ByVal tbl As String, ByVal flds As String(), ByVal con As SQLite.SQLiteConnection)
            Dim qry As String = String.Format("CREATE TABLE `{0}`(", tbl)
            For i As Integer = 0 To flds.Length - 1
                qry &= IIf(i = 0, flds(i), "," & flds(i))
            Next
            qry &= ")"

            SQLiteConfiguration.ExecuteQuery(qry, con)
        End Sub

        Public Shared Sub AlterTablename(tablename As String, newTablename As String, con As SQLite.SQLiteConnection)
            SQLiteConfiguration.ExecuteDataReader("Alter Table `" & tablename & "` Rename To `" & newTablename & "`", con)
        End Sub

        Public Sub AlterTablename(tablename As String, newTablename As String)
            SQLiteConfiguration.AlterTablename(tablename, newTablename, Connection)
        End Sub

        Public Function ExecuteDataReader(query As String) As SQLite.SQLiteDataReader
            Return SQLiteConfiguration.ExecuteDataReader(query, Connection)
        End Function

        Public Shared Function ExecuteDataReader(query As String, _con As SQLite.SQLiteConnection) As SQLite.SQLiteDataReader
            Return New SQLite.SQLiteCommand(query, _con).ExecuteReader
        End Function

        Public Function ExecuteQuery(query As String) As Boolean
            Return SQLiteConfiguration.ExecuteQuery(query, Connection)
        End Function

        Public Shared Function ExecuteQuery(query As String, _con As SQLite.SQLiteConnection) As Boolean
            Try
                Dim command As New SQLite.SQLiteCommand(query, _con)
                command.ExecuteNonQuery()
                command.Dispose()
                Return True
            Catch ex As Exception
                MsgBox(ex.Message & " ExecuteQuery")
                Return False
            End Try
        End Function

        Public Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object())
            SQLiteConfiguration.Insert(tbl, fld, val, Connection)
        End Sub
        Public Shared Sub Insert(ByVal tbl As String, ByVal fields As String(), ByVal values As Object(), con As SQLite.SQLiteConnection)
            Dim qry As String = String.Format("INSERT INTO `{0}` (", tbl)
            Dim valtype As String = ""

            For i As Integer = 0 To fields.Length - 1
                Dim f As String = fields(i)
                If f = fields(0) Then
                    qry &= String.Format("`{0}`", f)
                Else
                    qry &= String.Format(",`{0}`", f)
                End If
            Next

            qry &= ") VALUES("

            For i As Integer = 0 To values.Length - 1
                Dim v = values(i)
                valtype = TypeName(v)
                Select Case valtype
                    Case "String"
                        qry &= String.Format("'{0}',", v)
                    Case "Date"
                        qry &= String.Format("'{0}',", Date.Parse(v).ToString("yyyy-MM-dd HH:mm:ss"))
                    Case Else
                        qry &= String.Format("{0},", v)
                End Select
            Next
            qry &= ")"
            qry = System.Text.RegularExpressions.Regex.Replace(qry, ",\)", ")")


            SQLiteConfiguration.ExecuteQuery(qry, con)
        End Sub
        Public Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object())
            SQLiteConfiguration.Update(tbl, fld, val, condition, Connection)
        End Sub
        Public Shared Sub Update(ByVal tbl As String, ByVal fields As String(), ByVal values As Object(), ByVal condition As Object(), ByVal con As SQLite.SQLiteConnection)
            Dim qry As String = String.Format("UPDATE {0} SET ", tbl)
            Dim valtype As String = ""

            If fields.Length = values.Length Then
                For f As Integer = 0 To fields.GetUpperBound(0)
                    valtype = TypeName(values(f))
                    If f = 0 Then
                        If valtype = "String" Then
                            qry &= String.Format("[{0}]='{1}'", fields(f), values(f))
                        Else
                            qry &= String.Format("[{0}]={1}", fields(f), values(f))
                        End If
                    Else
                        If valtype = "String" Then
                            qry &= String.Format(",[{0}]='{1}'", fields(f), values(f))
                        Else
                            qry &= String.Format(",[{0}]={1}", fields(f), values(f))
                        End If
                    End If
                Next
            End If

            If Not condition Is Nothing Then
                If TypeName(condition(1)) = "String" Then
                    qry &= String.Format(" WHERE {0} = '{1}'", condition(0), condition(1))
                Else
                    qry &= String.Format(" WHERE {0} = {1}", condition(0), condition(1))
                End If
            End If

            SQLiteConfiguration.ExecuteQuery(qry, con)
        End Sub

        Public Function ToDT(qry As String) As DataTable
            Return SQLiteConfiguration.ToDT(qry, Connection)
        End Function
        Public Shared Function ToDT(qry As String, _con As SQLite.SQLiteConnection) As DataTable
            Try
                Dim dt As New DataTable
                Dim command As New SQLite.SQLiteDataAdapter(qry, _con)
                command.Fill(dt)
                command.Dispose()
                Return dt
            Catch ex As Exception
                MsgBox(ex.Message & " ExecuteQuery")
                Return Nothing
            End Try
        End Function

        Public Function CheckTable(tbl As String) As Boolean
            Return SQLiteConfiguration.CheckTable(tbl, Connection)
        End Function
        Public Function GetTables() As List(Of String)
            Return SQLiteConfiguration.GetTables(Connection)
        End Function
        Public Shared Function CheckTable(tbl As String, _con As SQLite.SQLiteConnection) As Boolean
            Return SQLiteConfiguration.GetTables(_con).Contains(tbl.ToLower)
        End Function
        Public Shared Function GetTables(_con As SQLite.SQLiteConnection) As List(Of String)
            Dim lst As New List(Of String)
            Using rdr As SQLite.SQLiteDataReader = SQLiteConfiguration.ExecuteDataReader("SELECT `name` FROM sqlite_master WHERE type='table';", _con)
                While rdr.Read
                    lst.Add(rdr.Item(0).ToString.ToLower)
                End While
            End Using
            Return lst
        End Function
    End Class
#End Region
#Region "MDB/DBF"
    Public Class MDBConfiguration

        '    Public Const MDBConfigFileExtension = ".mdb.config.xml"

        Public Connection As OleDb.OleDbConnection
        Public DBPath As String
        Public UserID As String
        Public Password As String

        Sub New(_dbpath As String, Optional _userid As String = "", Optional _password As String = "", Optional openNow As Boolean = False)
            DBPath = _dbpath
            UserID = _userid
            Password = _password
            If openNow Then
                Open()
            End If
        End Sub


        Public Sub Open()
            MDBConfiguration.Open(DBPath, Connection, UserID, Password)
        End Sub
        Public Sub CreateTable(ByVal tbl As String, ByVal flds As String())
            MDBConfiguration.CreateTable(tbl, flds, Connection)
        End Sub
        Public Sub ExecuteQuery(ByVal Qry As String)
            MDBConfiguration.ExecuteQuery(Qry, Connection)
        End Sub
        Public Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object())
            MDBConfiguration.Insert(tbl, fld, val, Connection)
        End Sub
        Public Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object())
            MDBConfiguration.Update(tbl, fld, val, condition, Connection)
        End Sub
        Public Function ToDT(ByVal qry As String) As DataTable
            Return MDBConfiguration.ToDT(qry, Connection)
        End Function
        Public Function CheckTable(tbl As String) As Boolean
            Return MDBConfiguration.CheckTable(tbl, Connection)
        End Function

        Public Sub Close()
            Connection.Close()
            Connection.Dispose()
        End Sub

        Public Shared Sub Open(ByVal _dbpath As String, ByRef _con As OleDb.OleDbConnection, Optional _userid As String = "", Optional _password As String = "")
            Try
                _con = New System.Data.OleDb.OleDbConnection(String.Format("Provider=Microsoft.JET.OLEDB.4.0;Data Source={0};User Id={1};Password={2};", _dbpath, _userid, _password))
                _con.Open()
            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        'Public Sub SaveToDBF(dbfPath As String, flds As String(),values As List(Of clspay)
        '    Dim ExportedDbf = New SocialExplorer.IO.FastDBF.DbfFile(Text.Encoding.GetEncoding(1252))
        '    ExportedDbf.Open(dbfPath, FileMode.Create)


        '    For Each col In flds
        '        ExportedDbf.Header.AddColumn(New SocialExplorer.IO.FastDBF.DbfColumn(col.ToString, SocialExplorer.IO.FastDBF.DbfColumn.DbfColumnType.Character, 50, 0))
        '    Next

        '    Dim Counter As Integer = 0

        '    For Each row In dt.Rows

        '        Dim ColumnCounter As Integer = 0
        '        Dim NewRec = New SocialExplorer.IO.FastDBF.DbfRecord(ExportedDbf.Header)

        '        For Each col In flds

        '            NewRec(ColumnCounter) = dt.Rows(Counter)(col).ToString
        '            ColumnCounter = ColumnCounter + 1
        '        Next

        '        ExportedDbf.Write(NewRec, True)
        '        Counter = Counter + 1

        '    Next

        '    ExportedDbf.Close()
        'End Sub

        Public Shared Sub Create(ByVal _dbpath As String, Optional _userid As String = "", Optional _password As String = "")
            Try
                Dim cat As New ADOX.Catalog
                cat.Create(String.Format("Provider=Microsoft.JET.OLEDB.4.0;Data Source={0};User Id={1};Password={2};", _dbpath, _userid, _password))
            Catch Ex As System.Exception
            End Try
        End Sub

        Public Shared Sub CreateTable(ByVal tbl As String, ByVal flds As String(), ByVal con As OleDb.OleDbConnection)
            Dim qry As String = String.Format("CREATE TABLE {0}(", tbl)
            For i As Integer = 0 To flds.Length - 1
                qry &= IIf(i = 0, flds(i), "," & flds(i))
            Next
            qry &= ")"

            MDBConfiguration.ExecuteQuery(qry, con)
        End Sub
        Public Shared Function ToDT(ByVal qry As String, ByVal con As OleDb.OleDbConnection) As DataTable
            Try
                Dim dt As New DataTable
                Dim da As New OleDb.OleDbDataAdapter(qry, con)
                da.Fill(dt)
                Return dt
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Shared Sub ExecuteQuery(ByVal Qry As String, ByVal con As OleDb.OleDbConnection)
            Try
                Dim com As New OleDb.OleDbCommand(Qry, con)
                com.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Public Shared Sub Insert(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), con As OleDb.OleDbConnection)
            Dim qry As String = String.Format("INSERT INTO {0} (", tbl)
            Dim valtype As String = ""

            For i As Integer = 0 To fld.Length - 1
                Dim f As String = fld(i)
                If f = fld(0) Then
                    qry &= String.Format("[{0}]", f)
                Else
                    qry &= String.Format(",[{0}]", f)
                End If
            Next

            qry &= ") VALUES("

            For i As Integer = 0 To val.Length - 1
                Dim v = val(i)
                valtype = TypeName(v)
                If i = 0 Then
                    If valtype = "String" Then
                        qry &= String.Format("'{0}'", v)
                    Else
                        qry &= String.Format("{0}", v)
                    End If
                Else
                    If valtype = "String" Then
                        qry &= String.Format(",'{0}'", v)
                    Else
                        qry &= String.Format(",{0}", v)
                    End If
                End If
            Next
            qry &= ")"

            MDBConfiguration.ExecuteQuery(qry, con)
        End Sub

        Public Shared Sub Update(ByVal tbl As String, ByVal fld As String(), ByVal val As Object(), ByVal condition As Object(), ByVal con As OleDb.OleDbConnection)
            Dim qry As String = String.Format("UPDATE {0} SET ", tbl)
            Dim valtype As String = ""

            If fld.Length = val.Length Then
                For f As Integer = 0 To fld.GetUpperBound(0)
                    valtype = TypeName(val(f))
                    If f = 0 Then
                        If valtype = "String" Then
                            qry &= String.Format("[{0}]='{1}'", fld(f), val(f))
                        Else
                            qry &= String.Format("[{0}]={1}", fld(f), val(f))
                        End If
                    Else
                        If valtype = "String" Then
                            qry &= String.Format(",[{0}]='{1}'", fld(f), val(f))
                        Else
                            qry &= String.Format(",[{0}]={1}", fld(f), val(f))
                        End If
                    End If
                Next
            End If

            If Not condition Is Nothing Then
                If TypeName(condition(1)) = "String" Then
                    qry &= String.Format(" WHERE {0} = '{1}'", condition(0), condition(1))
                Else
                    qry &= String.Format(" WHERE {0} = {1}", condition(0), condition(1))
                End If
            End If

            MDBConfiguration.ExecuteQuery(qry, con)
        End Sub

        Public Shared Function CheckTable(tbl As String, con As OleDb.OleDbConnection) As Boolean
            Return getTables(con).Contains(tbl)
        End Function

        Public Shared Function getTables(ByVal con As OleDb.OleDbConnection) As List(Of String)
            getTables = New List(Of String)
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            Dim dt As DataTable = con.GetSchema("Tables", restrictions)
            For i As Integer = 0 To dt.Rows.Count - 1
                getTables.Add(dt.Rows(i)(2).ToString)
            Next
            Return getTables
        End Function

    End Class
#End Region
#Region "Misc"
    Public Class SQLCondition
        Public Field As String
        Public Value As Object
        Public Conjunction As String

        Sub New(_field As String, _value As Object, Optional _conjunction As String = "")
            Field = _field
            Value = _value
            Conjunction = _conjunction
        End Sub

        Public Overrides Function ToString() As String
            Dim v = Value
            Dim valtype As String = TypeName(v)
            Select Case valtype
                Case "String"
                    Return String.Format(" `{0}` = '{1}' {2}", Field, Value, Conjunction)
                Case "Date"
                    Return String.Format(" `{0}` = '{1}' {2}", Field, Date.Parse(Value).ToString("yyyy-MM-dd HH:mm:ss"), Conjunction)
                Case Else
                    Return String.Format(" {0} = {1} {2}", Field, Value, Conjunction)
            End Select
        End Function
    End Class
#End Region
End Class
