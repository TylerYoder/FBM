Public Class menuPopup

    Private Sub menuPopup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StartTimePicker.Value() = DateTime.Now()
        EndTimePicker.Value() = DateTime.Now()

    End Sub

    Private Sub menuPopupButton_Click(sender As Object, e As EventArgs) Handles menuPopupButton.Click
        If (StartTimePicker.Value() > EndTimePicker.Value()) Then
            MsgBox("Start Time cannot be after End Time. Please try again.")
        Else
            Me.Close()
            main.tempShiftFlag = tempShiftFlagCheckBox.Checked()
            main.generateMenu(StartTimePicker.Value, EndTimePicker.Value)
        End If
    End Sub
End Class