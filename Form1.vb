Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim allSQL As String = TextBox1.Text
        Dim L As New List(Of String)
        L.AddRange(Split(allSQL, ";"))
        L = L.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()

        For Each SQL As String In L
            Try
                ExecuteNonQuery(InfoDB, SQL)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next
        TextBox1.Text = ""
    End Sub

    Property InfoDB As String
        Get
            Dim DBPAth As String = System.AppDomain.CurrentDomain.BaseDirectory & "\Database3.mdb"
            Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & DBPAth & "'"
        End Get
        Set(value As String)

        End Set
    End Property

    Sub ExecuteNonQuery(ByVal DBConnection As String, ByVal SQL As String)
        Dim Dcom As New Data.OleDb.OleDbCommand
        Dim Dcon As New Data.OleDb.OleDbConnection

        Dcon.ConnectionString = DBConnection
        Dcom.Connection = Dcon
        Dcom.CommandText = SQL

        'MsgBox(System.AppDomain.CurrentDomain.BaseDirectory)

        Try
            Dcom.Connection.Open()
        Catch ex As Exception

        End Try

        Try
            Dcom.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            Dcom.Connection.Close()
        Catch ex As Exception

        End Try

    End Sub

End Class