Imports System.Xml

Public Class sideLinkerPopup

    Private Sub sideLinkerPopup_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub sideLinkerButton_Click(sender As Object, e As EventArgs) Handles sideLinkerButton.Click
        If sideBox1.Text = "" Then
            If sideBox2.Text = "" Then
                Dim result As Integer = MsgBox("Are you sure you don't want to include sides for this meal?", MsgBoxStyle.YesNo)
                If result = DialogResult.No Then
                    Exit Sub
                End If
            Else
                MsgBox("Please fill in the first side before you fill in the second side. This will prevent formatting issues when printing menus.")
                Exit Sub
            End If
        ElseIf sideBox1.Text <> "" And sideBox2.Text = "" Then
            Dim result As Integer = MsgBox("Are you sure you only want to include one side for this meal?", MsgBoxStyle.YesNo)
            If result = DialogResult.No Then
                Exit Sub
            End If
        End If
        Me.Close()
    End Sub
End Class