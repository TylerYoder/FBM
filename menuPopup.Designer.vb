﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class menuPopup
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.StartTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.EndTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.tempShiftFlagCheckBox = New System.Windows.Forms.CheckBox()
        Me.menuPopupButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Start Date:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "End Date:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 55)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Temp Shift?:"
        '
        'StartTimePicker
        '
        Me.StartTimePicker.CustomFormat = "ddd', 'MMMdd"
        Me.StartTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.StartTimePicker.Location = New System.Drawing.Point(76, 3)
        Me.StartTimePicker.Name = "StartTimePicker"
        Me.StartTimePicker.Size = New System.Drawing.Size(100, 20)
        Me.StartTimePicker.TabIndex = 1
        Me.StartTimePicker.Value = New Date(2014, 9, 24, 0, 0, 0, 0)
        '
        'EndTimePicker
        '
        Me.EndTimePicker.CustomFormat = "ddd', 'MMMdd"
        Me.EndTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.EndTimePicker.Location = New System.Drawing.Point(76, 29)
        Me.EndTimePicker.Name = "EndTimePicker"
        Me.EndTimePicker.Size = New System.Drawing.Size(100, 20)
        Me.EndTimePicker.TabIndex = 2
        '
        'tempShiftFlagCheckBox
        '
        Me.tempShiftFlagCheckBox.AutoSize = True
        Me.tempShiftFlagCheckBox.Location = New System.Drawing.Point(76, 55)
        Me.tempShiftFlagCheckBox.Name = "tempShiftFlagCheckBox"
        Me.tempShiftFlagCheckBox.Size = New System.Drawing.Size(15, 14)
        Me.tempShiftFlagCheckBox.TabIndex = 3
        Me.tempShiftFlagCheckBox.UseVisualStyleBackColor = True
        '
        'menuPopupButton
        '
        Me.menuPopupButton.Location = New System.Drawing.Point(58, 75)
        Me.menuPopupButton.Name = "menuPopupButton"
        Me.menuPopupButton.Size = New System.Drawing.Size(75, 23)
        Me.menuPopupButton.TabIndex = 4
        Me.menuPopupButton.Text = "Generate"
        Me.menuPopupButton.UseVisualStyleBackColor = True
        '
        'menuPopup
        '
        Me.AcceptButton = Me.menuPopupButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(189, 107)
        Me.Controls.Add(Me.menuPopupButton)
        Me.Controls.Add(Me.tempShiftFlagCheckBox)
        Me.Controls.Add(Me.EndTimePicker)
        Me.Controls.Add(Me.StartTimePicker)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "menuPopup"
        Me.Text = "Generate Menu"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents StartTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents EndTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents tempShiftFlagCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents menuPopupButton As System.Windows.Forms.Button
End Class
