Public Class frmOLTC
    Private WithEvents objAccessXWRM10A As New clsAccessXWRM10A
    Dim objResponse As New Response
    Dim VarTapNO As Integer = objResponse.TapNumber
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox1.Text = objResponse.CURRENT & " " & objResponse.CURRENT_UNIT
        TextBox2.Text = objResponse.READING & " " & objResponse.READING_UNIT
        TextBox3.Text = objResponse.T1Reading
    End Sub

    Private Sub frmOLTC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        dgvHRdata.AllowUserToAddRows = False
        dgvHRdata.RowHeadersVisible = False
        dgvHRdata.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgvHRdata.Columns.Add("colTap", "Tap No")
        dgvHRdata.Columns.Add("colUN", "UN")
        dgvHRdata.Columns.Add("colvn", "VN")
        dgvHRdata.Columns.Add("colwn", "WN")

        Dim rowIndex As Integer = dgvHRdata.Rows.Add()
        With dgvHRdata.Rows(rowIndex)
            .Cells("colTap").Value = objResponse.TapNumber
            If objResponse.TapNumber > VarTapNO Then
                .Cells("colUN").Value = objResponse.READING
                Exit Sub
            End If
            .Cells("colvn").Value = objResponse.READING
            .Cells("colwn").Value = objResponse.READING

        End With


    End Sub
End Class