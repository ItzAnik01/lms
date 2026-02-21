Public Class LMS
    Public userRole As String
    Public currentUser As String

    Private Sub LMS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' If logged in as Student, restrict controls
        If userRole = "Student" Then
            Button1.Enabled = False   ' Registration ❌
            Button3.Enabled = False   ' Students ❌
            Button4.Enabled = False   ' Faculties ❌
            Button5.Enabled = False   ' Book Issue ❌
            Button6.Enabled = False   ' Book Return ❌
        End If

    End Sub

    ' ===== Registration =====
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim registration As New Registration()
        registration.ShowDialog()
    End Sub

    ' ===== Books Table =====
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim bookstable As New BooksTable()
        bookstable.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim customertable As New CustomerTable()
        customertable.ShowDialog()
    End Sub
    ' ===== Book Search (Everyone can open) =====
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim bookstable As New BooksTable()
        bookstable.ShowDialog()
    End Sub

    ' ===== Logout =====
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Hide()
        Login_Page.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim customertable As New CustomerTable()
        customertable.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim bookstable As New BooksTable()
        bookstable.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim bookstable As New BooksTable()
        bookstable.ShowDialog()
    End Sub
End Class