Public Class frmMain

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        ' Declare and get password length
        Dim length As Integer = 0
        Try
            length = CInt(txtSize.Text)
        Catch ex As Exception
            MessageBox.Show("The password length must be an integer.", "Size Argument Error", MessageBoxButtons.OK)
            Exit Sub
        End Try

        ' generate password
        txtPassword.Text = genPassword(length)
    End Sub

    Private Function genPassword(length As Integer)
        Dim rand As Random = New Random()
        Dim password As String = ""
        Dim key As Integer
        For i = 0 To length - 1
            key = rand.Next(74) + 1
            If (key <= 10) Then
                password += key.ToString
            ElseIf (key <= 62) Then
                key -= 10
                password += mapAlphabet(key)
            Else
                key -= 62
                password += mapSpecial(key)
            End If
        Next
        Return password
    End Function ' generates a password of given length
    Private Function mapAlphabet(key As Integer) As String
        Dim myKey As Integer = key
        If myKey > 26 Then
            myKey -= 26
        End If
        Dim c As String = ""
        Select Case myKey
            Case 1
                c = "a"
            Case 2
                c = "b"
            Case 3
                c = "c"
            Case 4
                c = "d"
            Case 5
                c = "e"
            Case 6
                c = "f"
            Case 7
                c = "g"
            Case 8
                c = "h"
            Case 9
                c = "i"
            Case 10
                c = "j"
            Case 11
                c = "k"
            Case 12
                c = "l"
            Case 13
                c = "m"
            Case 14
                c = "n"
            Case 15
                c = "o"
            Case 16
                c = "p"
            Case 17
                c = "q"
            Case 18
                c = "r"
            Case 19
                c = "s"
            Case 20
                c = "t"
            Case 21
                c = "u"
            Case 22
                c = "v"
            Case 23
                c = "w"
            Case 24
                c = "x"
            Case 25
                c = "y"
            Case 26
                c = "z"
        End Select

        If key > 26 Then
            c = c.ToUpper()
        End If

        Return c
    End Function ' get [a-z] [A-Z]
    Private Function mapSpecial(key As Integer) As String
        Dim c As String = ""
        Select Case key
            Case 1
                c = "!"
            Case 2
                c = "@"
            Case 3
                c = "#"
            Case 4
                c = "$"
            Case 5
                c = "%"
            Case 6
                c = "^"
            Case 7
                c = "~"
            Case 8
                c = "_"
            Case 9
                c = "*"
            Case 10
                c = "-"
            Case 11
                c = "+"
            Case 12
                c = "="
        End Select
        Return c
    End Function ' get Special
End Class
