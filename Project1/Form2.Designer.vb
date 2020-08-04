<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Timer1
        '
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(312, 36)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(130, 41)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "KIRIM SINYAL :: A"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Timer2
        '
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(312, 83)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(406, 41)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "BERSIH"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(7, 131)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(299, 42)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "CEK LAMPU"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 15
        Me.ListBox1.Location = New System.Drawing.Point(7, 36)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(299, 94)
        Me.ListBox1.TabIndex = 4
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(312, 130)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(406, 43)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "EXIT"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(449, 36)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(130, 41)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "KIRIM SINYAL :: B"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(588, 36)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(130, 41)
        Me.Button6.TabIndex = 7
        Me.Button6.Text = "KIRIM SINYAL :: C"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Copperplate Gothic Bold", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(322, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 33)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "00"
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(728, 240)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
