Imports System.Data.OleDb

Public Class CustomerTable

    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\.\.\..\College.mdb;")
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    ' ================= FORM LOAD =================
    Private Sub CustomerTable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()

        ' Disable buttons for Student role
        If LMS.userRole = "Student" Then
            Button1.Enabled = False ' Insert
            Button2.Enabled = False ' Update
            Button3.Enabled = False ' Delete
        End If
    End Sub

    ' ================= LOAD DATA =================
    Sub LoadData(Optional ByVal filter As String = "")
        Try
            con.Open()
            Dim query As String

            ' Dynamic search filter
            If filter <> "" Then
                query = "SELECT * FROM Customer WHERE CustomerName LIKE ? OR BookName LIKE ?"
                da = New OleDbDataAdapter(query, con)
                da.SelectCommand.Parameters.AddWithValue("?", "%" & filter & "%")
                da.SelectCommand.Parameters.AddWithValue("?", "%" & filter & "%")
            ElseIf LMS.userRole = "Student" Then
                query = "SELECT * FROM Customer WHERE CustomerName=?"
                da = New OleDbDataAdapter(query, con)
                da.SelectCommand.Parameters.AddWithValue("?", LMS.currentUser)
            Else
                query = "SELECT * FROM Customer"
                da = New OleDbDataAdapter(query, con)
            End If

            dt = New DataTable()
            da.Fill(dt)
            DataGridView1.DataSource = dt

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ================= INSERT =================
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            con.Open()
            Dim cmd As New OleDbCommand("INSERT INTO Customer (CustomerName, Phone, [Address], BookID, BookName) VALUES (?,?,?,?,?)", con)

            cmd.Parameters.Add("?", OleDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = Val(TextBox3.Text)
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = TextBox4.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = Val(TextBox5.Text)
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = TextBox6.Text

            cmd.ExecuteNonQuery()
            con.Close()

            MessageBox.Show("Record Inserted Successfully")
            LoadData()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ================= UPDATE =================
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If TextBox1.Text = "" Then
                MessageBox.Show("Select a record to update")
                Exit Sub
            End If

            con.Open()
            Dim cmd As New OleDbCommand("UPDATE Customer SET CustomerName=?, Phone=?, [Address]=?, BookID=?, BookName=? WHERE CustomerID=?", con)

            cmd.Parameters.Add("?", OleDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = Val(TextBox3.Text)
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = TextBox4.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = Val(TextBox5.Text)
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = TextBox6.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = Val(TextBox1.Text)

            cmd.ExecuteNonQuery()
            con.Close()

            MessageBox.Show("Record Updated Successfully")
            LoadData()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ================= DELETE =================
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If TextBox1.Text = "" Then
                MessageBox.Show("Select a record to delete")
                Exit Sub
            End If

            con.Open()
            Dim cmd As New OleDbCommand("DELETE FROM Customer WHERE CustomerID=?", con)
            cmd.Parameters.AddWithValue("?", Val(TextBox1.Text))
            cmd.ExecuteNonQuery()
            con.Close()

            MessageBox.Show("Record Deleted Successfully")
            LoadData()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ================= SEARCH =================
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        ' TextBox7 is the search box
        LoadData(TextBox7.Text)
    End Sub

    ' ================= DATAGRIDVIEW ROW SELECT =================
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            TextBox1.Text = row.Cells("CustomerID").Value.ToString()
            TextBox2.Text = row.Cells("CustomerName").Value.ToString()
            TextBox3.Text = row.Cells("Phone").Value.ToString()
            TextBox4.Text = row.Cells("Address").Value.ToString()
            TextBox5.Text = row.Cells("BookID").Value.ToString()
            TextBox6.Text = row.Cells("BookName").Value.ToString()
        End If
    End Sub

End Class