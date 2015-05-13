<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class sideLinkerPopup
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(sideLinkerPopup))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.sideLinkerButton = New System.Windows.Forms.Button()
        Me.sideBox1 = New System.Windows.Forms.ComboBox()
        Me.sideBox2 = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.MaximumSize = New System.Drawing.Size(200, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "First Side: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 113)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Second Side:"
        '
        'sideLinkerButton
        '
        Me.sideLinkerButton.Location = New System.Drawing.Point(74, 132)
        Me.sideLinkerButton.Name = "sideLinkerButton"
        Me.sideLinkerButton.Size = New System.Drawing.Size(75, 23)
        Me.sideLinkerButton.TabIndex = 5
        Me.sideLinkerButton.Text = "Link Sides"
        Me.sideLinkerButton.UseVisualStyleBackColor = True
        '
        'sideBox1
        '
        Me.sideBox1.FormattingEnabled = True
        Me.sideBox1.Location = New System.Drawing.Point(83, 78)
        Me.sideBox1.Name = "sideBox1"
        Me.sideBox1.Size = New System.Drawing.Size(123, 21)
        Me.sideBox1.TabIndex = 6
        '
        'sideBox2
        '
        Me.sideBox2.FormattingEnabled = True
        Me.sideBox2.Location = New System.Drawing.Point(83, 105)
        Me.sideBox2.Name = "sideBox2"
        Me.sideBox2.Size = New System.Drawing.Size(123, 21)
        Me.sideBox2.TabIndex = 6
        '
        'sideLinkerPopup
        '
        Me.AcceptButton = Me.sideLinkerButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(219, 167)
        Me.ControlBox = False
        Me.Controls.Add(Me.sideLinkerButton)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.sideBox1)
        Me.Controls.Add(Me.sideBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "sideLinkerPopup"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Side Linker"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents sideLinkerButton As System.Windows.Forms.Button
    Friend WithEvents sideBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents sideBox2 As System.Windows.Forms.ComboBox
End Class
