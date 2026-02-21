Imports System.Data.OleDb

Public Class Registration

    'DATABASE CONNECTION WITH YOUR PATH
    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\.\.\..\College.mdb;")
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    'FORM LOAD — SHOW USERS
    Private Sub Registration_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()


    End Sub

    'LOAD DATA INTO GRID
    Sub LoadData()
        Try
            con.Open()
            da = New OleDbDataAdapter("SELECT [Name], [Designation], [UserName], [Password] FROM [Register]", con)
            dt = New DataTable()
            da.Fill(dt)
            DataGridView1.DataSource = dt
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    'REGISTER BUTTON
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Please fill all fields")

            Exit Sub
        End If

        If TextBox3.Text <> TextBox4.Text Then
            MessageBox.Show("Passwords do not match")
            Exit Sub
        End If

        Try
            con.Open()

            Dim cmd As New OleDbCommand("INSERT INTO [Register] ([Name], [Designation], [UserName], [Password], [Confirm_Password]) VALUES (?,?,?,?,?)", con)

            cmd.Parameters.AddWithValue("?", TextBox1.Text)        'Name
            cmd.Parameters.AddWithValue("?", ComboBox1.Text)       'Designation
            cmd.Parameters.AddWithValue("?", TextBox2.Text)        'Username
            cmd.Parameters.AddWithValue("?", TextBox3.Text)        'Password
            cmd.Parameters.AddWithValue("?", TextBox4.Text)        'Confirm Password

            cmd.ExecuteNonQuery()
            MessageBox.Show("Registration Successful")

            con.Close()
            LoadData()

            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            ComboBox1.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox2.Text = "" Then
            MessageBox.Show("Enter UserName to delete")
            Exit Sub
        End If

        Try
            con.Open()

            Dim cmd As New OleDbCommand("DELETE FROM [Register] WHERE [UserName]=?", con)
            cmd.Parameters.AddWithValue("?", TextBox2.Text)

            Dim result As Integer = cmd.ExecuteNonQuery()

            If result > 0 Then
                MessageBox.Show("User Deleted Successfully")
            Else
                MessageBox.Show("User not found")
            End If

            con.Close()
            LoadData()

            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            ComboBox1.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try

    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString()
            ComboBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
            TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString()
            TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString()
        End If
    End Sub


End Class