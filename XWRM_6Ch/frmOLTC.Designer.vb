<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOLTC
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
        Me.components = New System.ComponentModel.Container()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.dgvHRdata = New System.Windows.Forms.DataGridView()
        Me.Label100 = New System.Windows.Forms.Label()
        Me.Label101 = New System.Windows.Forms.Label()
        Me.Label102 = New System.Windows.Forms.Label()
        Me.Label103 = New System.Windows.Forms.Label()
        Me.Label104 = New System.Windows.Forms.Label()
        Me.Label105 = New System.Windows.Forms.Label()
        Me.Label106 = New System.Windows.Forms.Label()
        Me.Label107 = New System.Windows.Forms.Label()
        Me.Label108 = New System.Windows.Forms.Label()
        Me.Label109 = New System.Windows.Forms.Label()
        Me.Label110 = New System.Windows.Forms.Label()
        Me.Label111 = New System.Windows.Forms.Label()
        Me.Label112 = New System.Windows.Forms.Label()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Panel3.SuspendLayout()
        CType(Me.dgvHRdata, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(105, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(158, 20)
        Me.TextBox1.TabIndex = 0
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(292, 12)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(149, 20)
        Me.TextBox2.TabIndex = 1
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(495, 12)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(171, 20)
        Me.TextBox3.TabIndex = 2
        '
        'Timer1
        '
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.dgvHRdata)
        Me.Panel3.Controls.Add(Me.Label100)
        Me.Panel3.Controls.Add(Me.Label101)
        Me.Panel3.Controls.Add(Me.Label102)
        Me.Panel3.Controls.Add(Me.Label103)
        Me.Panel3.Controls.Add(Me.Label104)
        Me.Panel3.Controls.Add(Me.Label105)
        Me.Panel3.Controls.Add(Me.Label106)
        Me.Panel3.Controls.Add(Me.Label107)
        Me.Panel3.Controls.Add(Me.Label108)
        Me.Panel3.Controls.Add(Me.Label109)
        Me.Panel3.Controls.Add(Me.Label110)
        Me.Panel3.Controls.Add(Me.Label111)
        Me.Panel3.Controls.Add(Me.Label112)
        Me.Panel3.Controls.Add(Me.Button9)
        Me.Panel3.Controls.Add(Me.Button10)
        Me.Panel3.Location = New System.Drawing.Point(115, 101)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(598, 468)
        Me.Panel3.TabIndex = 308
        '
        'dgvHRdata
        '
        Me.dgvHRdata.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvHRdata.Location = New System.Drawing.Point(20, 149)
        Me.dgvHRdata.Name = "dgvHRdata"
        Me.dgvHRdata.Size = New System.Drawing.Size(561, 255)
        Me.dgvHRdata.TabIndex = 327
        '
        'Label100
        '
        Me.Label100.AutoSize = True
        Me.Label100.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label100.Location = New System.Drawing.Point(511, 116)
        Me.Label100.Name = "Label100"
        Me.Label100.Size = New System.Drawing.Size(42, 20)
        Me.Label100.TabIndex = 321
        Me.Label100.Text = "10 A"
        '
        'Label101
        '
        Me.Label101.AutoSize = True
        Me.Label101.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label101.Location = New System.Drawing.Point(413, 116)
        Me.Label101.Name = "Label101"
        Me.Label101.Size = New System.Drawing.Size(107, 20)
        Me.Label101.TabIndex = 320
        Me.Label101.Text = "LV Current : - "
        '
        'Label102
        '
        Me.Label102.AutoSize = True
        Me.Label102.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label102.Location = New System.Drawing.Point(154, 115)
        Me.Label102.Name = "Label102"
        Me.Label102.Size = New System.Drawing.Size(33, 20)
        Me.Label102.TabIndex = 319
        Me.Label102.Text = "1 A"
        '
        'Label103
        '
        Me.Label103.AutoSize = True
        Me.Label103.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label103.Location = New System.Drawing.Point(47, 115)
        Me.Label103.Name = "Label103"
        Me.Label103.Size = New System.Drawing.Size(110, 20)
        Me.Label103.TabIndex = 318
        Me.Label103.Text = "HV Current : - "
        '
        'Label104
        '
        Me.Label104.AutoSize = True
        Me.Label104.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label104.Location = New System.Drawing.Point(145, 61)
        Me.Label104.Name = "Label104"
        Me.Label104.Size = New System.Drawing.Size(31, 20)
        Me.Label104.TabIndex = 317
        Me.Label104.Text = " 11"
        '
        'Label105
        '
        Me.Label105.AutoSize = True
        Me.Label105.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label105.Location = New System.Drawing.Point(47, 61)
        Me.Label105.Name = "Label105"
        Me.Label105.Size = New System.Drawing.Size(103, 20)
        Me.Label105.TabIndex = 316
        Me.Label105.Text = "Report No : - "
        '
        'Label106
        '
        Me.Label106.AutoSize = True
        Me.Label106.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label106.Location = New System.Drawing.Point(169, 87)
        Me.Label106.Name = "Label106"
        Me.Label106.Size = New System.Drawing.Size(49, 20)
        Me.Label106.TabIndex = 315
        Me.Label106.Text = " Ynyn"
        '
        'Label107
        '
        Me.Label107.AutoSize = True
        Me.Label107.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label107.Location = New System.Drawing.Point(47, 87)
        Me.Label107.Name = "Label107"
        Me.Label107.Size = New System.Drawing.Size(126, 20)
        Me.Label107.TabIndex = 314
        Me.Label107.Text = "Vector Group : - "
        '
        'Label108
        '
        Me.Label108.AutoSize = True
        Me.Label108.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label108.Location = New System.Drawing.Point(464, 88)
        Me.Label108.Name = "Label108"
        Me.Label108.Size = New System.Drawing.Size(89, 20)
        Me.Label108.TabIndex = 313
        Me.Label108.Text = "HH:MM:SS"
        '
        'Label109
        '
        Me.Label109.AutoSize = True
        Me.Label109.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label109.Location = New System.Drawing.Point(412, 88)
        Me.Label109.Name = "Label109"
        Me.Label109.Size = New System.Drawing.Size(51, 20)
        Me.Label109.TabIndex = 312
        Me.Label109.Text = "Time :"
        '
        'Label110
        '
        Me.Label110.AutoSize = True
        Me.Label110.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label110.Location = New System.Drawing.Point(464, 58)
        Me.Label110.Name = "Label110"
        Me.Label110.Size = New System.Drawing.Size(89, 20)
        Me.Label110.TabIndex = 311
        Me.Label110.Text = "DD/MM/YY"
        '
        'Label111
        '
        Me.Label111.AutoSize = True
        Me.Label111.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label111.Location = New System.Drawing.Point(411, 58)
        Me.Label111.Name = "Label111"
        Me.Label111.Size = New System.Drawing.Size(56, 20)
        Me.Label111.TabIndex = 310
        Me.Label111.Text = "Date : "
        '
        'Label112
        '
        Me.Label112.AutoSize = True
        Me.Label112.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label112.Location = New System.Drawing.Point(184, 6)
        Me.Label112.Name = "Label112"
        Me.Label112.Size = New System.Drawing.Size(299, 31)
        Me.Label112.TabIndex = 309
        Me.Label112.Text = "Heat Run Test Report"
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.Maroon
        Me.Button9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button9.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Button9.Location = New System.Drawing.Point(481, 411)
        Me.Button9.Margin = New System.Windows.Forms.Padding(4)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(100, 40)
        Me.Button9.TabIndex = 253
        Me.Button9.Text = "Back"
        Me.Button9.UseVisualStyleBackColor = False
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.Maroon
        Me.Button10.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Button10.Location = New System.Drawing.Point(20, 411)
        Me.Button10.Margin = New System.Windows.Forms.Padding(4)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(100, 40)
        Me.Button10.TabIndex = 252
        Me.Button10.Text = "Print"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'frmOLTC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(829, 670)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "frmOLTC"
        Me.Text = "frmOLTC"
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.dgvHRdata, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Panel3 As Panel
    Friend WithEvents dgvHRdata As DataGridView
    Friend WithEvents Label100 As Label
    Friend WithEvents Label101 As Label
    Friend WithEvents Label102 As Label
    Friend WithEvents Label103 As Label
    Friend WithEvents Label104 As Label
    Friend WithEvents Label105 As Label
    Friend WithEvents Label106 As Label
    Friend WithEvents Label107 As Label
    Friend WithEvents Label108 As Label
    Friend WithEvents Label109 As Label
    Friend WithEvents Label110 As Label
    Friend WithEvents Label111 As Label
    Friend WithEvents Label112 As Label
    Friend WithEvents Button9 As Button
    Friend WithEvents Button10 As Button
End Class
