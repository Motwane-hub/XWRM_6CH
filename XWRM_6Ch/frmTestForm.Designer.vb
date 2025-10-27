<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTestForm
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblTemp1 = New System.Windows.Forms.Label()
        Me.lblVectorGroup = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.LBL = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.lblTime = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.lblLV3Reading = New System.Windows.Forms.Label()
        Me.lblLV2Reading = New System.Windows.Forms.Label()
        Me.lblLV1Reading = New System.Windows.Forms.Label()
        Me.lblHV3Reading = New System.Windows.Forms.Label()
        Me.lblHV2Reading = New System.Windows.Forms.Label()
        Me.lblHV1Reading = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.lblLVCurrent = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.lblHVCurrent = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.trmTestON = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Maroon
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Button1.Location = New System.Drawing.Point(404, 394)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 40)
        Me.Button1.TabIndex = 264
        Me.Button1.Text = "STOP"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'lblTemp1
        '
        Me.lblTemp1.AutoSize = True
        Me.lblTemp1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemp1.Location = New System.Drawing.Point(624, 116)
        Me.lblTemp1.Name = "lblTemp1"
        Me.lblTemp1.Size = New System.Drawing.Size(25, 16)
        Me.lblTemp1.TabIndex = 263
        Me.lblTemp1.Text = "- - -"
        '
        'lblVectorGroup
        '
        Me.lblVectorGroup.AutoSize = True
        Me.lblVectorGroup.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVectorGroup.Location = New System.Drawing.Point(201, 116)
        Me.lblVectorGroup.Name = "lblVectorGroup"
        Me.lblVectorGroup.Size = New System.Drawing.Size(40, 16)
        Me.lblVectorGroup.TabIndex = 262
        Me.lblVectorGroup.Text = " Ynyn"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(609, 68)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(89, 20)
        Me.Label26.TabIndex = 261
        Me.Label26.Text = "HH:MM:SS"
        '
        'LBL
        '
        Me.LBL.AutoSize = True
        Me.LBL.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBL.Location = New System.Drawing.Point(155, 68)
        Me.LBL.Name = "LBL"
        Me.LBL.Size = New System.Drawing.Size(21, 20)
        Me.LBL.TabIndex = 260
        Me.LBL.Text = "D"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(561, 116)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(56, 16)
        Me.Label22.TabIndex = 259
        Me.Label22.Text = "Temp :- "
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(103, 116)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(102, 16)
        Me.Label21.TabIndex = 258
        Me.Label21.Text = "Vector Group : - "
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.Maroon
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label20.Location = New System.Drawing.Point(350, 98)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(68, 24)
        Me.Label20.TabIndex = 257
        Me.Label20.Text = "Result"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(141, 16)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(500, 25)
        Me.Label19.TabIndex = 256
        Me.Label19.Text = "Winding Resistance Meter Model-XWRM-25A6"
        '
        'lblTime
        '
        Me.lblTime.AutoSize = True
        Me.lblTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTime.Location = New System.Drawing.Point(560, 19)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(51, 20)
        Me.lblTime.TabIndex = 255
        Me.lblTime.Text = "Time :"
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate.Location = New System.Drawing.Point(102, 68)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(56, 20)
        Me.lblDate.TabIndex = 254
        Me.lblDate.Text = "Date : "
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.Maroon
        Me.Button7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Button7.Location = New System.Drawing.Point(588, 394)
        Me.Button7.Margin = New System.Windows.Forms.Padding(4)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(100, 40)
        Me.Button7.TabIndex = 253
        Me.Button7.Text = "Save"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Maroon
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Button6.Location = New System.Drawing.Point(234, 394)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(100, 40)
        Me.Button6.TabIndex = 252
        Me.Button6.Text = "Back"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Maroon
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Button5.Location = New System.Drawing.Point(107, 394)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(100, 40)
        Me.Button5.TabIndex = 251
        Me.Button5.Text = "Print"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label46.Location = New System.Drawing.Point(142, 323)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(37, 20)
        Me.Label46.TabIndex = 250
        Me.Label46.Text = "- - -"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label47.Location = New System.Drawing.Point(143, 285)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(37, 20)
        Me.Label47.TabIndex = 249
        Me.Label47.Text = "- - -"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label48.Location = New System.Drawing.Point(142, 245)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(37, 20)
        Me.Label48.TabIndex = 248
        Me.Label48.Text = "- - -"
        '
        'lblLV3Reading
        '
        Me.lblLV3Reading.AutoSize = True
        Me.lblLV3Reading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLV3Reading.Location = New System.Drawing.Point(585, 323)
        Me.lblLV3Reading.Name = "lblLV3Reading"
        Me.lblLV3Reading.Size = New System.Drawing.Size(37, 20)
        Me.lblLV3Reading.TabIndex = 247
        Me.lblLV3Reading.Text = "- - -"
        '
        'lblLV2Reading
        '
        Me.lblLV2Reading.AutoSize = True
        Me.lblLV2Reading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLV2Reading.Location = New System.Drawing.Point(585, 283)
        Me.lblLV2Reading.Name = "lblLV2Reading"
        Me.lblLV2Reading.Size = New System.Drawing.Size(37, 20)
        Me.lblLV2Reading.TabIndex = 246
        Me.lblLV2Reading.Text = "- - -"
        '
        'lblLV1Reading
        '
        Me.lblLV1Reading.AutoSize = True
        Me.lblLV1Reading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLV1Reading.Location = New System.Drawing.Point(585, 242)
        Me.lblLV1Reading.Name = "lblLV1Reading"
        Me.lblLV1Reading.Size = New System.Drawing.Size(37, 20)
        Me.lblLV1Reading.TabIndex = 245
        Me.lblLV1Reading.Text = "- - -"
        '
        'lblHV3Reading
        '
        Me.lblHV3Reading.AutoSize = True
        Me.lblHV3Reading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHV3Reading.Location = New System.Drawing.Point(311, 323)
        Me.lblHV3Reading.Name = "lblHV3Reading"
        Me.lblHV3Reading.Size = New System.Drawing.Size(37, 20)
        Me.lblHV3Reading.TabIndex = 244
        Me.lblHV3Reading.Text = "- - -"
        '
        'lblHV2Reading
        '
        Me.lblHV2Reading.AutoSize = True
        Me.lblHV2Reading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHV2Reading.Location = New System.Drawing.Point(311, 283)
        Me.lblHV2Reading.Name = "lblHV2Reading"
        Me.lblHV2Reading.Size = New System.Drawing.Size(37, 20)
        Me.lblHV2Reading.TabIndex = 243
        Me.lblHV2Reading.Text = "- - -"
        '
        'lblHV1Reading
        '
        Me.lblHV1Reading.AutoSize = True
        Me.lblHV1Reading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHV1Reading.Location = New System.Drawing.Point(311, 245)
        Me.lblHV1Reading.Name = "lblHV1Reading"
        Me.lblHV1Reading.Size = New System.Drawing.Size(37, 20)
        Me.lblHV1Reading.TabIndex = 242
        Me.lblHV1Reading.Text = "- - -"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.Location = New System.Drawing.Point(489, 323)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(31, 20)
        Me.Label37.TabIndex = 241
        Me.Label37.Text = "wn"
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.Location = New System.Drawing.Point(489, 283)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(27, 20)
        Me.Label38.TabIndex = 240
        Me.Label38.Text = "vn"
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.Location = New System.Drawing.Point(488, 241)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(29, 20)
        Me.Label39.TabIndex = 239
        Me.Label39.Text = "un"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.Location = New System.Drawing.Point(230, 323)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(37, 20)
        Me.Label36.TabIndex = 238
        Me.Label36.Text = "WN"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.Location = New System.Drawing.Point(230, 283)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(33, 20)
        Me.Label35.TabIndex = 237
        Me.Label35.Text = "VN"
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.Location = New System.Drawing.Point(230, 245)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(34, 20)
        Me.Label34.TabIndex = 236
        Me.Label34.Text = "UN"
        '
        'lblLVCurrent
        '
        Me.lblLVCurrent.AutoSize = True
        Me.lblLVCurrent.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLVCurrent.Location = New System.Drawing.Point(585, 209)
        Me.lblLVCurrent.Name = "lblLVCurrent"
        Me.lblLVCurrent.Size = New System.Drawing.Size(37, 20)
        Me.lblLVCurrent.TabIndex = 235
        Me.lblLVCurrent.Text = "- - -"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(470, 209)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(69, 20)
        Me.Label32.TabIndex = 234
        Me.Label32.Text = "Current"
        '
        'lblHVCurrent
        '
        Me.lblHVCurrent.AutoSize = True
        Me.lblHVCurrent.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHVCurrent.Location = New System.Drawing.Point(311, 209)
        Me.lblHVCurrent.Name = "lblHVCurrent"
        Me.lblHVCurrent.Size = New System.Drawing.Size(37, 20)
        Me.lblHVCurrent.TabIndex = 233
        Me.lblHVCurrent.Text = "- - -"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(215, 209)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(69, 20)
        Me.Label30.TabIndex = 232
        Me.Label30.Text = "Current"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(128, 209)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(81, 20)
        Me.Label29.TabIndex = 231
        Me.Label29.Text = "Tap No . "
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(488, 170)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(153, 24)
        Me.Label24.TabIndex = 230
        Me.Label24.Text = "Secondary / LV"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(186, 170)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(127, 24)
        Me.Label23.TabIndex = 229
        Me.Label23.Text = "Primary / HV"
        '
        'trmTestON
        '
        '
        'frmTestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblTemp1)
        Me.Controls.Add(Me.lblVectorGroup)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.LBL)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.lblTime)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label46)
        Me.Controls.Add(Me.Label47)
        Me.Controls.Add(Me.Label48)
        Me.Controls.Add(Me.lblLV3Reading)
        Me.Controls.Add(Me.lblLV2Reading)
        Me.Controls.Add(Me.lblLV1Reading)
        Me.Controls.Add(Me.lblHV3Reading)
        Me.Controls.Add(Me.lblHV2Reading)
        Me.Controls.Add(Me.lblHV1Reading)
        Me.Controls.Add(Me.Label37)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.Label36)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.lblLVCurrent)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.lblHVCurrent)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Label23)
        Me.Name = "frmTestForm"
        Me.Text = "frmTestForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents lblTemp1 As Label
    Friend WithEvents lblVectorGroup As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents LBL As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents lblTime As Label
    Friend WithEvents lblDate As Label
    Friend WithEvents Button7 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Label46 As Label
    Friend WithEvents Label47 As Label
    Friend WithEvents Label48 As Label
    Friend WithEvents lblLV3Reading As Label
    Friend WithEvents lblLV2Reading As Label
    Friend WithEvents lblLV1Reading As Label
    Friend WithEvents lblHV3Reading As Label
    Friend WithEvents lblHV2Reading As Label
    Friend WithEvents lblHV1Reading As Label
    Friend WithEvents Label37 As Label
    Friend WithEvents Label38 As Label
    Friend WithEvents Label39 As Label
    Friend WithEvents Label36 As Label
    Friend WithEvents Label35 As Label
    Friend WithEvents Label34 As Label
    Friend WithEvents lblLVCurrent As Label
    Friend WithEvents Label32 As Label
    Friend WithEvents lblHVCurrent As Label
    Friend WithEvents Label30 As Label
    Friend WithEvents Label29 As Label
    Friend WithEvents Label24 As Label
    Friend WithEvents Label23 As Label
    Friend WithEvents trmTestON As Timer
End Class
