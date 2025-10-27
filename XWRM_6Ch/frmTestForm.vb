Public Class frmTestForm

    'Dim WithEvents objXWRM10A As New AccessXWRM25.clsAccessXWRM10A
    'Dim objResponse As AccessXWRM25.Response

    Private WithEvents objXWRM10A As New clsAccessXWRM10A
    Dim objResponse As Response


    Private Sub frmTestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        trmTestON.Start()
    End Sub

    Private Sub trmTestON_Tick(sender As Object, e As EventArgs) Handles trmTestON.Tick
        'If Not objXWRM10A.IsCommunicated Then
        '    MsgBox("Device Not Connected")
        '    'testOff()
        '    'PictureBox1.Visible = False
        '    'PictureBox2.Visible = False
        '    'Label10.Visible = False
        '    'Button5.Enabled = False
        '    'pnlTestParameters.Visible = False
        '    'pnlTestResult.Visible = False
        '    Try
        '        objXWRM10A.Disconnect()
        '    Catch ex As Exception

        '    End Try
        '    objXWRM10A.Connect("COM7")
        'End If
        objXWRM10A.Connect("COM7")
        'cntr -= 1
        'lblSec.Text = cntr.ToString
        'If Label10.Visible = True Then
        '    Label10.Visible = False
        'Else
        '    Label10.Visible = True
        'End If

        'If PictureBox1.Visible = True Then
        '    PictureBox1.Visible = False
        '    PictureBox2.Visible = False
        'Else
        '    PictureBox1.Visible = True
        '    PictureBox2.Visible = True
        'End If

        objResponse = objXWRM10A.GetResponse

        'If Not (lblCH1Reading.Text.Trim = "-" Or lblCH1Reading.Text.Trim = "0.0 Ohm") Then
        'If Not IsNothing(objResponse) Then
        '    'Label4.Text = objResponse.ToString
        '    tmrHRTime.Start()
        'End If

        'lblTestMode.Text = objResponse.TestModeStatus
        lblHVCurrent.Text = objResponse.HVTestCurrent
        lblLVCurrent.Text = objResponse.LVTestCurrent

        'lblTotalReportCount.Text = objResponse.TotalReportCount

        'Select Case objResponse.TestModeStatus
        '    Case objXWRM10A.enmTestMode.HR
        '        lblTestMode.Text = "TEST MODE: " & "HEAT RUN"

        '    Case objXWRM10A.enmTestMode.NORMAL
        '        lblTestMode.Text = "TEST MODE: " & "NORMAL"

        '    Case clsAccessXWRM10A.enmTestMode.OLTC
        '        If lblTapCaption.Visible = False Then
        '            lblTapCaption.Visible = True
        '            lblTapNumber.Visible = True

        '            lblTapCaption.Refresh()
        '            lblTapNumber.Refresh()
        '        End If

        '        lblTestMode.Text = "TEST MODE: " & "OLTC"
        'End Select

        'Select Case objResponse.TestCurrentUnit
        '    Case objXWRM10A.enmCurrentsUnit.Ampere

        'lblCorrectedRes.Text = FormatChannelReading(GetCorrectedReading1(Convert.ToDouble(Val(lblCH1Reading.Text)), Convert.ToDouble(Val(lblTemprature1.Text))))

        Select Case objResponse.HVTestCurrentUnit
            Case objXWRM10A.enmCurrentsUnit.Ampere
                lblHVCurrent.Text = FormatCurrentReading(Val(lblHVCurrent.Text)) & " A"
            Case objXWRM10A.enmCurrentsUnit.mAmpere
                lblHVCurrent.Text = FormatCurrentReading(Val(lblHVCurrent.Text)) & " mA"
        End Select

        Select Case objResponse.LVTestCurrentUnit
            Case objXWRM10A.enmCurrentsUnit.Ampere
                lblLVCurrent.Text = FormatCurrentReading(Val(lblLVCurrent.Text)) & " A"
            Case objXWRM10A.enmCurrentsUnit.mAmpere
                lblLVCurrent.Text = FormatCurrentReading(Val(lblLVCurrent.Text)) & " mA"
        End Select

        'Select Case objResponse.CH1Reading

        '    Case 333333344
        '        lblCorrectedRes.Text = "Rev"

        '    Case 555555584
        '        lblCorrectedRes.Text = "???"

        '        lblCorrectedRes.Refresh()
        '        res1.Refresh()
        '    Case 111111112
        '        lblCorrectedRes.Text = "U L"

        '    Case 1000000000
        '        lblCorrectedRes.Text = "O L"

        'End Select

        'If Not (lblCorrectedRes.Text = "Rev" Or lblCorrectedRes.Text = "???" Or lblCorrectedRes.Text = "U L" Or lblCorrectedRes.Text = "O L") Then

        Select Case objResponse.Winding

            Case objXWRM10A.enmWinding.UN

                lblHV1Reading.Text = FormatChannelReading(objResponse.HVReading)

                Select Case objResponse.HVUnit
                    Case objXWRM10A.enmResistanceUnit.KOhm
                        lblHV1Reading.Text = lblHV1Reading.Text + " kOhm"

                    Case objXWRM10A.enmResistanceUnit.mOhm
                        lblHV1Reading.Text = lblHV1Reading.Text + " mOhm"

                    Case objXWRM10A.enmResistanceUnit.Ohm
                        lblHV1Reading.Text = lblHV1Reading.Text + " Ohm"

                    Case objXWRM10A.enmResistanceUnit.uOhm
                        lblHV1Reading.Text = lblHV1Reading.Text + " uOhm"

                End Select
                'End If


                lblLV1Reading.Text = FormatChannelReading(objResponse.LVReading)

                Select Case objResponse.LVUnit
                    Case objXWRM10A.enmResistanceUnit.KOhm
                        lblLV1Reading.Text = lblLV1Reading.Text + " kOhm"

                    Case objXWRM10A.enmResistanceUnit.mOhm
                        lblLV1Reading.Text = lblLV1Reading.Text + " mOhm"

                    Case objXWRM10A.enmResistanceUnit.Ohm
                        lblLV1Reading.Text = lblLV1Reading.Text + " Ohm"

                    Case objXWRM10A.enmResistanceUnit.uOhm
                        lblLV1Reading.Text = lblLV1Reading.Text + " uOhm"

                End Select


            Case objXWRM10A.enmWinding.VN

                lblHV2Reading.Text = FormatChannelReading(objResponse.HVReading)

                Select Case objResponse.HVUnit
                    Case objXWRM10A.enmResistanceUnit.KOhm
                        lblHV2Reading.Text = lblHV2Reading.Text + " kOhm"

                    Case objXWRM10A.enmResistanceUnit.mOhm
                        lblHV2Reading.Text = lblHV2Reading.Text + " mOhm"

                    Case objXWRM10A.enmResistanceUnit.Ohm
                        lblHV2Reading.Text = lblHV2Reading.Text + " Ohm"

                    Case objXWRM10A.enmResistanceUnit.uOhm
                        lblHV2Reading.Text = lblHV2Reading.Text + " uOhm"

                End Select
                'End If


                lblLV2Reading.Text = FormatChannelReading(objResponse.LVReading)

                Select Case objResponse.LVUnit
                    Case objXWRM10A.enmResistanceUnit.KOhm
                        lblLV2Reading.Text = lblLV2Reading.Text + " kOhm"

                    Case objXWRM10A.enmResistanceUnit.mOhm
                        lblLV2Reading.Text = lblLV2Reading.Text + " mOhm"

                    Case objXWRM10A.enmResistanceUnit.Ohm
                        lblLV2Reading.Text = lblLV2Reading.Text + " Ohm"

                    Case objXWRM10A.enmResistanceUnit.uOhm
                        lblLV2Reading.Text = lblLV2Reading.Text + " uOhm"

                End Select


            Case objXWRM10A.enmWinding.WN

                lblHV3Reading.Text = FormatChannelReading(objResponse.HVReading)

                Select Case objResponse.HVUnit
                    Case objXWRM10A.enmResistanceUnit.KOhm
                        lblHV3Reading.Text = lblHV3Reading.Text + " kOhm"

                    Case objXWRM10A.enmResistanceUnit.mOhm
                        lblHV3Reading.Text = lblHV3Reading.Text + " mOhm"

                    Case objXWRM10A.enmResistanceUnit.Ohm
                        lblHV3Reading.Text = lblHV3Reading.Text + " Ohm"

                    Case objXWRM10A.enmResistanceUnit.uOhm
                        lblHV3Reading.Text = lblHV3Reading.Text + " uOhm"

                End Select
                'End If


                lblLV3Reading.Text = FormatChannelReading(objResponse.LVReading)

                Select Case objResponse.LVUnit
                    Case objXWRM10A.enmResistanceUnit.KOhm
                        lblLV3Reading.Text = lblLV3Reading.Text + " kOhm"

                    Case objXWRM10A.enmResistanceUnit.mOhm
                        lblLV3Reading.Text = lblLV3Reading.Text + " mOhm"

                    Case objXWRM10A.enmResistanceUnit.Ohm
                        lblLV3Reading.Text = lblLV3Reading.Text + " Ohm"

                    Case objXWRM10A.enmResistanceUnit.uOhm
                        lblLV3Reading.Text = lblLV3Reading.Text + " uOhm"

                End Select


        End Select

        'Case objXWRM10A.enmCurrentsUnit.mAmpere
        'lblTestCurrent.Text = FormatCurrentReading(Val(lblTestCurrent.Text)) & " mAmp"

        ' End Select

        'If lblTestMode.Text = "TEST MODE: " & "NORMAL" Then
        '    If cmdNormal.Text <> "Normal Readings" Then
        '        lblCH1Reading.Text = FormatChannelReading(objResponse.Corrected_CH1_Reading)

        '    Else
        '        lblCH1Reading.Text = FormatChannelReading(objResponse.CH1Reading)

        '    End If
        'Else
        '    lblCH1Reading.Text = FormatChannelReading(objResponse.CH1Reading)

        'End If

        'lblCH1Reading.Text = FormatChannelReading(objResponse.CH1Reading)

        'Select Case objResponse.CH1Reading
        '    Case 333333344
        '        lblCH1Reading.Text = "Rev"

        '    Case 555555584
        '        lblCH1Reading.Text = "???"

        '    Case 111111112
        '        lblCH1Reading.Text = "U L"

        '    Case 1000000000
        '        lblCH1Reading.Text = "O L"

        'End Select

        'If Not (lblCH1Reading.Text = "Rev" Or lblCH1Reading.Text = "???" Or lblCH1Reading.Text = "U L" Or lblCH1Reading.Text = "O L") Then
        '    Select Case objResponse.CH1Unit
        '        Case objXWRM10A.enmResistanceUnit.KOhm
        '            lblCH1Reading.Text = lblCH1Reading.Text + " kOhm"

        '        Case objXWRM10A.enmResistanceUnit.mOhm
        '            lblCH1Reading.Text = lblCH1Reading.Text + " mOhm"

        '        Case objXWRM10A.enmResistanceUnit.Ohm
        '            lblCH1Reading.Text = lblCH1Reading.Text + " Ohm"

        '        Case objXWRM10A.enmResistanceUnit.uOhm
        '            lblCH1Reading.Text = lblCH1Reading.Text + " Ohm"

        '    End Select
        'End If

        'If lblTestMode.Text = "TEST MODE: " & "NORMAL" Then
        '    If cmdNormal.Text <> "Normal Readings" Then
        '        lblCH2Reading.Text = FormatChannelReading(objResponse.Corrected_CH2_Reading)
        '    Else
        '        lblCH2Reading.Text = FormatChannelReading(objResponse.CH2Reading)
        '    End If
        'Else
        '    lblCH2Reading.Text = FormatChannelReading(objResponse.CH2Reading)
        'End If

        'Select Case objResponse.CH2Reading
        '    Case 333333344
        '        lblCH2Reading.Text = "Rev"

        '    Case 555555584
        '        lblCH2Reading.Text = "???"

        '    Case 111111112
        '        lblCH2Reading.Text = "U L"

        '    Case 1000000000
        '        lblCH2Reading.Text = "O L"

        'End Select

        'If Not (lblCH2Reading.Text = "Rev" Or lblCH2Reading.Text = "???" Or lblCH2Reading.Text = "U L" Or lblCH2Reading.Text = "O L") Then
        '    Select Case objResponse.CH2Unit
        '        Case objXWRM10A.enmResistanceUnit.KOhm
        '            lblCH2Reading.Text = lblCH2Reading.Text + " kOhm"
        '        Case objXWRM10A.enmResistanceUnit.mOhm
        '            lblCH2Reading.Text = lblCH2Reading.Text + " mOhm"
        '        Case objXWRM10A.enmResistanceUnit.Ohm
        '            lblCH2Reading.Text = lblCH2Reading.Text + " Ohm"
        '        Case objXWRM10A.enmResistanceUnit.uOhm
        '            lblCH2Reading.Text = lblCH2Reading.Text + " μOhm"
        '    End Select
        'End If

        'If lblTestMode.Text = "TEST MODE: " & "NORMAL" Then
        '    If cmdNormal.Text <> "Normal Readings" Then
        '        lblCH3Reading.Text = FormatChannelReading(objResponse.Corrected_CH3_Reading)
        '    Else
        '        lblCH3Reading.Text = FormatChannelReading(objResponse.CH3Reading)
        '    End If
        'Else
        '    lblCH3Reading.Text = FormatChannelReading(objResponse.CH3Reading)
        'End If

        'Select Case objResponse.CH3Reading
        '    Case 333333344
        '        lblCH3Reading.Text = "Rev"

        '    Case 555555584
        '        lblCH3Reading.Text = "???"

        '    Case 111111112
        '        lblCH3Reading.Text = "U L"

        '    Case 1000000000
        '        lblCH3Reading.Text = "O L"

        'End Select

        'If Not (lblCH3Reading.Text = "Rev" Or lblCH3Reading.Text = "???" Or lblCH3Reading.Text = "U L" Or lblCH3Reading.Text = "O L") Then
        '    Select Case objResponse.CH3Unit
        '        Case objXWRM10A.enmResistanceUnit.KOhm
        '            lblCH3Reading.Text = lblCH3Reading.Text + " kOhm"
        '        Case objXWRM10A.enmResistanceUnit.mOhm
        '            lblCH3Reading.Text = lblCH3Reading.Text + " mOhm"
        '        Case objXWRM10A.enmResistanceUnit.Ohm
        '            lblCH3Reading.Text = lblCH3Reading.Text + " Ohm"
        '        Case objXWRM10A.enmResistanceUnit.uOhm
        '            lblCH3Reading.Text = lblCH3Reading.Text + " μOhm"
        '    End Select
        'End If

        lblTemp1.Text = objResponse.T1Reading & " °C"
        'lblTemprature2.Text = objResponse.T2Reading & " °C"
        'lblTapNumber.Text = objResponse.TapNumber

        'lblCorrectedRes.Refresh()
        'lblCH1Reading.Refresh()
        'res1.Refresh()
    End Sub

    Private Function FormatCurrentReading(ByVal Value As Double) As String
        Dim strSplit As String() = Value.ToString.Split(".")
        Dim intInteger As Integer = 0

        If strSplit.Length > 1 Then
            Select Case Val(strSplit(0)).ToString.Length
                Case 1
                    Return Format(Value, "0.000")

                Case 2
                    Return Format(Value, "0.00")

                Case 3
                    Return Format(Value, "0.0")

                Case 4
                    Return Value

                Case Else
                    Return Value
            End Select
        Else
            Return Format(Value, "0.0")
        End If

        Return Value
    End Function

    Private Function FormatChannelReading(ByVal Value As Double) As String
        Dim strSplit As String() = Value.ToString.Split(".")
        Dim intInteger As Integer = 0

        If strSplit.Length > 1 Then
            Select Case Val(strSplit(0)).ToString.Length
                Case 1
                    Return Format(Value, "0.0000")

                Case 2
                    Return Format(Value, "0.000")

                Case 3
                    Return Format(Value, "0.00")
                Case 4
                    Return Format(Value, "0.0")

                Case 5
                    Return Value
            End Select
        Else
            Return Format(Value, "0.0")
        End If

        Return Value
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If objXWRM10A.IsConnected = True Then
            objXWRM10A.TestOff()
            System.Threading.Thread.Sleep(1000)

            objXWRM10A.TestOff()
            System.Threading.Thread.Sleep(1000)
        End If

        trmTestON.Stop()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        Dim f As New Form1
        f.ShowDialog()
        Me.Close()
    End Sub
End Class