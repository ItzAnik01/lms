Imports System.Data.OleDb

Public Class BooksTable

    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\.\.\..\College.mdb;")
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    Private Sub BooksTable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
        Me.ActiveControl = Nothing

        ' ===== Role-based access =====
        If LMS.userRole = "Student" Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False

            TextBox1.ReadOnly = True
            TextBox2.ReadOnly = True
            TextBox3.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox5.ReadOnly = True
            TextBox6.ReadOnly = True
        End If
    End Sub

    ' ===== LOAD DATA WITH SEARCH =====
    Sub LoadData(Optional searchValue As String = "")
        Try
            con.Open()

            Dim query As String

            If searchValue <> "" Then
                query = "SELECT * FROM Books WHERE BookName LIKE ? OR CustomerName LIKE ?"
                da = New OleDbDataAdapter(query, con)
                da.SelectCommand.Parameters.AddWithValue("?", "%" & searchValue & "%")
                da.SelectCommand.Parameters.AddWithValue("?", "%" & searchValue & "%")

            ElseIf LMS.userRole = "Student" Then
                query = "SELECT * FROM Books WHERE CustomerId = ?"
                da = New OleDbDataAdapter(query, con)
                da.SelectCommand.Parameters.AddWithValue("?", LMS.currentUser)

            Else
                query = "SELECT * FROM Books"
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

    ' ===== INSERT =====
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
                MessageBox.Show("Please fill all fields")
                Exit Sub
            End If

            con.Open()
            Dim cmd As New OleDbCommand("INSERT INTO Books (BookId, BookName, IssueDate, ReturnDate, CustomerId, CustomerName) VALUES (?,?,?,?,?,?)", con)
            cmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox2.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox3.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox4.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox5.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox6.Text.Trim())
            cmd.ExecuteNonQuery()
            con.Close()

            MessageBox.Show("Book Added Successfully")
            LoadData()
            ClearFields()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ===== UPDATE =====
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If TextBox1.Text = "" Then
                MessageBox.Show("Select a book to update")
                Exit Sub
            End If

            con.Open()
            Dim cmd As New OleDbCommand("UPDATE Books SET BookName=?, IssueDate=?, ReturnDate=?, CustomerId=?, CustomerName=? WHERE BookId=?", con)
            cmd.Parameters.AddWithValue("?", TextBox2.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox3.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox4.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox5.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox6.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())
            cmd.ExecuteNonQuery()
            con.Close()

            MessageBox.Show("Book Updated Successfully")
            LoadData()
            ClearFields()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ===== DELETE =====
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If TextBox1.Text = "" Then
                MessageBox.Show("Select a book to delete")
                Exit Sub
            End If

            con.Open()
            Dim cmd As New OleDbCommand("DELETE FROM Books WHERE BookId=?", con)
            cmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())
            cmd.ExecuteNonQuery()
            con.Close()

            MessageBox.Show("Book Deleted Successfully")
            LoadData()
            ClearFields()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    ' ===== LIVE SEARCH FROM TEXTBOX7 =====
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        LoadData(TextBox7.Text.Trim())
    End Sub

    ' ===== FILL TEXTBOXES WHEN ROW SELECTED =====
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            TextBox1.Text = row.Cells("BookId").Value.ToString()
            TextBox2.Text = row.Cells("BookName").Value.ToString()
            TextBox3.Text = row.Cells("IssueDate").Value.ToString()
            TextBox4.Text = row.Cells("ReturnDate").Value.ToString()
            TextBox5.Text = row.Cells("CustomerId").Value.ToString()
            TextBox6.Text = row.Cells("CustomerName").Value.ToString()
        End If
    End Sub

    ' ===== CLEAR =====
    Sub ClearFields()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
    End Sub

End Class