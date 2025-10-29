Imports System.Drawing.Drawing2D
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports AccessXWRM25

Public Class Form1
    Private WithEvents objAccessXWRM10A As New clsAccessXWRM10A
    'Private WithEvents objAccessXWRM10A As New AccessXWRM25.clsAccessXWRM10A
    Public strComPort As String

    Dim objResponse As Response
    'Dim VarTapNO As Integer = objResponse.TapNumber
    Private currentReadingIndex As Integer = 0
    Private totalTaps As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'RoundFormCorners(Me, 100) ' 25 = corner radius
        'Panel1.BackColor = Color.LightBlue
        'Panel1.BorderStyle = BorderStyle.None
        'RoundPanelCorners(Panel1, 40)  ' 20 = corner radius
        loadcom()
        Label49.Text = Now.Date

        Label44.Text = DateAndTime.Now.ToString("T")
    End Sub

    Private Sub loadcom()
        Try
            ddlComPort.Items.Clear()
            For inti = 0 To My.Computer.Ports.SerialPortNames.Count - 1
                ddlComPort.Items.Add(My.Computer.Ports.SerialPortNames(inti))
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        strComPort = ddlComPort.Text
        ConnectToInstrument()
    End Sub

    Private Sub ConnectToInstrument()
        Try
            'objAccessXWRM10A = New AccessXWRM25.clsAccessXWRM10A

            'ddlComPort.Visible = False
            'lblConnectToInstrument.Visible = False

            Application.DoEvents()

            'lblStatus.ForeColor = colColor

            'picWait.Visible = True

            'lblStatus.Visible = True
            'lblStatus.Text = "Detecting COM Ports."

            'ShowPanel(enmPanelID.PANEL_CONNECTION)

            objAccessXWRM10A.Connect(strComPort)

            'If objAccessXWRM10A.IsRegistered = False Then
            '    objAccessXWRM10A = Nothing
            '    End
            'End If

            If objAccessXWRM10A.IsConnected = False Then
                'lblStatus.ForeColor = Color.Red

                'picWait.Visible = False
                'lblStatus.Text = "Connection Failed, Please Check COM PORT Setting."

                'ddlComPort.Visible = True
                'lblConnectToInstrument.Visible = True
                MsgBox("Connection Failed, Please Check COM PORT Setting.", MsgBoxStyle.Critical)

                Application.DoEvents()
                Exit Sub
            End If

            objAccessXWRM10A.Home()
            System.Threading.Thread.Sleep(1000)

            'objAccessXWRM10A.GetDateTime()
            'System.Threading.Thread.Sleep(1000)

            'lblStatus.Text = "Verifying Communication."
            MsgBox("Verifying Communication.", MsgBoxStyle.Information)
            Application.DoEvents()

            If objAccessXWRM10A.IsCommunicated Then
                'lblStatus.Text = "Instrument Connected Sucessfully."
                MsgBox("Instrument Connected Sucessfully.", MsgBoxStyle.Information)
                Application.DoEvents()
            Else

                objAccessXWRM10A.Disconnect()
                'lblStatus.ForeColor = Color.Red

                'picWait.Visible = False
                'lblStatus.Text = "Instrument Communication Error."

                'ddlComPort.Visible = True
                'lblConnectToInstrument.Visible = True

                MsgBox("Instrument Communication Error.", MsgBoxStyle.Critical)

                Application.DoEvents()

                Exit Sub
            End If


            'pnlMeterPanel.Enabled = True

            'lblStatus.Text = "Verifying Date and Time."
            'MsgBox("Verifying Date and Time.", MsgBoxStyle.Information)
            Application.DoEvents()

            'System.Threading.Thread.Sleep(1000)

            'dtmDateTime = objAccessXWRM10A._ResponseResult.DeviceDateTime

            'lblDisplayDate.Text = Format(dtmDateTime, "ddd, dd MMM, yyyy")
            'lbDisplayTime.Text = Format(dtmDateTime, "hh:mm:ss tt")

            'tmrClock.Enabled = True

            'lblStatus.Text = "Initialized Sucessfully, Starting Application..."
            'Application.DoEvents()

            'System.Threading.Thread.Sleep(500)

            'strComPort = strComPort

            'Gotohome()

            'SavePort()

            'If System.IO.File.Exists(Application.StartupPath + "\srno.txt") Then
            '    lblInstrumentSerialNumberInfo.Text = System.IO.File.ReadAllText(Application.StartupPath + "\srno.txt")
            'End If

            'If dtmDateTime > System.DateTime.Now.AddMinutes(1) Or dtmDateTime < System.DateTime.Now.AddMinutes(-1) Then
            '    Dim frmMessageBox As New frmMessageBox("Instrument date is not synchronized with the system date," & vbCrLf & " are you sure u want to synchronize it.", False, XWRM10A.frmMessageBox.MessageStyle.INFO_MESSAGE)
            '    frmMessageBox.ShowDialog()

            '    If frmMessageBox.IsYes Then
            '        SyncDate()
            '    End If
            'End If
            'cmbTestMode.SelectedIndex = 0
        Catch ex As Exception
            'lblStatus.ForeColor = Color.Red

            'picWait.Visible = False
            'lblStatus.Text = "Connection Failed, Please Check COM PORT Setting."
            MsgBox("Connection Failed, Please Check COM PORT Setting.", MsgBoxStyle.Critical)
            'pnlBackground.Visible = True
            'ddlComPort.Visible = True
            'lblConnectToInstrument.Visible = True

            Application.DoEvents()
        End Try

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        loadcom()
    End Sub

    Private Sub setParameter()

    End Sub


    Private Sub SetVisibility(hr As Boolean, totalTap As Boolean, range As Boolean)
        Label9.Visible = hr
        txtHRTime.Visible = hr
        Label8.Visible = totalTap
        txtTotalTap.Visible = totalTap
        Label2.Visible = range
        Label5.Visible = range
        cmbHVRange.Visible = range
        cmbLVRange.Visible = range
    End Sub

    Private Sub cmbTestMode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTestMode.SelectedIndexChanged

        cmbHVCurrent.Enabled = True
        cmbLVCurrent.Enabled = True
        cmbHVCurrent.Text = ""
        cmbLVCurrent.Text = ""
        txtHRTime.Text = ""
        txtTotalTap.Text = ""
        cmbHVRange.Text = ""
        cmbLVRange.Text = ""

        Select Case cmbTestMode.SelectedIndex
            Case 0, 3, 4
                SetVisibility(False, False, False)

                UN.Visible = True
                VN.Visible = True
                WN.Visible = True

                cmbWindingSelection.Items.Clear()
                cmbWindingSelection.Items.Add("ALL WINDING = 1")
                cmbWindingSelection.Items.Add("DUAL WINDING =2")
                cmbWindingSelection.Items.Add("SINGLE = 3")

                txtTemp1Loc.Visible = False
                Label27.Visible = False
                txtTemp2Loc.Visible = False
                Label25.Visible = False

            Case 1
                SetVisibility(True, False, True)

                UN.Visible = True
                VN.Visible = True
                WN.Visible = True

                cmbWindingSelection.Items.Clear()
                cmbWindingSelection.Items.Add("ALL WINDING = 1")
                cmbWindingSelection.Items.Add("DUAL WINDING =2")
                cmbWindingSelection.Items.Add("SINGLE = 3")

                txtTemp1Loc.Visible = True
                Label27.Visible = True
                txtTemp2Loc.Visible = True
                Label25.Visible = True



            Case 2
                SetVisibility(False, True, False)

                UN.Visible = False
                VN.Visible = False
                WN.Visible = False

                cmbWindingSelection.Items.Clear()
                cmbWindingSelection.Items.Add("HV WINDING = 4")
                cmbWindingSelection.Items.Add("LV WINDING = 5")

                txtTemp1Loc.Visible = False
                Label27.Visible = False
                txtTemp2Loc.Visible = False
                Label25.Visible = False
                Panel1.Visible = True
                'Dim f As New frmOLTC
                'f.Show()
        End Select

        'Select Case cmbTestMode.SelectedIndex
        '    Case 0, 3, 4
        '        SetVisibility(False, False, False)

        '    Case 1
        '        SetVisibility(True, False, True)
        '        cmbHVRange.Enabled = True
        '        cmbLVRange.Enabled = True
        '        cmbHVRange.SelectedIndex = -1
        '        cmbLVRange.SelectedIndex = -1

        '    Case 2
        '        SetVisibility(False, True, True)
        '        cmbHVRange.Enabled = True
        '        cmbLVRange.Enabled = True
        '        cmbHVRange.SelectedIndex = -1
        '        cmbLVRange.SelectedIndex = -1
        'End Select

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Dim objTestmode As New clsAccessXWRM10A.enmTestMode
        Dim objHVRange As New clsAccessXWRM10A.enmRange
        Dim objLVRange As New clsAccessXWRM10A.enmRange
        Dim intHrTime As Integer
        Dim intTotalTap As Integer
        Dim vectorGroupYNYN As Integer
        Dim intCurrentTap As Integer = Val(txtCurrentTap.Text)
        Dim vectorGroup As New clsAccessXWRM10A.enmVectorGroup
        Dim windingSelection As New clsAccessXWRM10A.enmWindingSelection
        Dim winding As New clsAccessXWRM10A.enmWinding
        Dim testObject As String = txtTestObject.Text
        Dim DUTSpecification As String = txtSpecification.Text
        Dim DUTSerialNumber As String = txtSerialNo.Text
        Dim DUTMake As String = txtMake.Text
        Dim DUTTestLocation As String = txtLocation.Text
        Dim year As String = txtYear.Text
        Dim Temp1Loc As String = txtTemp1Loc.Text
        Dim Temp2Loc As String = txtTemp2Loc.Text

        Try
            'If CheckConnection() = False Then Exit Sub

            'Dim frmMessagebox As frmMessageBox

            Select Case cmbTestMode.SelectedIndex
                Case 0
                    objTestmode = clsAccessXWRM10A.enmTestMode.NORMAL
                    intTotalTap = 0
                    intHrTime = 0


                Case 1
                    objTestmode = clsAccessXWRM10A.enmTestMode.HR
                    intTotalTap = 0
                    intHrTime = Val(txtHRTime.Text)
                    If Val(txtHRTime.Text) < 1 Then
                        'frmMessagebox = New frmMessageBox("Please Enter HR Interval", True, frmMessagebox.MessageStyle.INFO_MESSAGE)
                        'frmMessagebox.ShowDialog()
                        MsgBox("Please Enter HR Interval", MsgBoxStyle.Critical)
                        txtHRTime.Focus()
                        Exit Sub
                    End If



                Case 2
                    objTestmode = clsAccessXWRM10A.enmTestMode.OLTC
                    objHVRange = clsAccessXWRM10A.enmRange.AUTORANGE
                    objLVRange = clsAccessXWRM10A.enmRange.AUTORANGE
                    intHrTime = 0

                    If Val(txtTotalTap.Text) < 1 Then
                        'frmMessagebox = New frmMessageBox("Please Enter Tap Number", True, frmMessagebox.MessageStyle.ERROR_MESSAGE)
                        'frmMessagebox.ShowDialog()
                        MsgBox("Please Enter Tap Number", MsgBoxStyle.Critical)

                        txtTotalTap.Focus()
                        Exit Sub
                    End If

                    If txtTotalTap.Text > 35 Then 'Dhananjay-01-12-2022
                        'frmMessagebox = New frmMessageBox("Please Enter Tap Number Between 1 to 35", True, frmMessagebox.MessageStyle.ERROR_MESSAGE)
                        'frmMessagebox.ShowDialog()
                        MsgBox("Please Enter Tap Number Between 1 to 35", MsgBoxStyle.Critical)

                        txtTotalTap.Focus()
                        Exit Sub
                    End If

                    intTotalTap = Val(txtTotalTap.Text)

                Case 3
                    objTestmode = clsAccessXWRM10A.enmTestMode.DEMAGNETIZING
                    intTotalTap = 0
                    intHrTime = 0
                    objHVRange = clsAccessXWRM10A.enmRange.AUTORANGE
                    objLVRange = clsAccessXWRM10A.enmRange.AUTORANGE

                Case 4
                    objTestmode = clsAccessXWRM10A.enmTestMode.DYNAMIC_RESISTANCE
                    intTotalTap = 0
                    intHrTime = 0
                    objHVRange = clsAccessXWRM10A.enmRange.AUTORANGE
                    objLVRange = clsAccessXWRM10A.enmRange.AUTORANGE

            End Select


            Select Case cmbVectorGroup.SelectedIndex
                Case 0 'YNYN

                    vectorGroup = clsAccessXWRM10A.enmVectorGroup.YNYN
                    If rdAuto.Checked Then
                        vectorGroupYNYN = 1
                    Else
                        vectorGroupYNYN = 2
                    End If

                Case 1 'YND

                    vectorGroup = clsAccessXWRM10A.enmVectorGroup.YND
                    vectorGroupYNYN = 0

                Case 2 'DYN

                    vectorGroup = clsAccessXWRM10A.enmVectorGroup.DYN
                    vectorGroupYNYN = 0

                Case 3 'DD

                    vectorGroup = clsAccessXWRM10A.enmVectorGroup.DD
                    vectorGroupYNYN = 0

            End Select

            Select Case cmbWindingSelection.SelectedIndex
                Case 0 'ALL WINDING = 1

                    windingSelection = clsAccessXWRM10A.enmWindingSelection.ALL_WINDING
                    winding = clsAccessXWRM10A.enmWinding.NA

                Case 1 'DUAL WINDING =2

                    windingSelection = clsAccessXWRM10A.enmWindingSelection.DUAL_WINDING
                    winding = clsAccessXWRM10A.enmWinding.NA

                Case 2 'SINGLE WINDING = 3

                    windingSelection = clsAccessXWRM10A.enmWindingSelection.SINGLE_WINDING

                    If UN.Checked And VN.Checked = False And WN.Checked = False Then
                        winding = clsAccessXWRM10A.enmWinding.UN
                    End If

                    If UN.Checked = False And VN.Checked And WN.Checked = False Then
                        winding = clsAccessXWRM10A.enmWinding.VN
                    End If

                    If UN.Checked = False And VN.Checked = False And WN.Checked Then
                        winding = clsAccessXWRM10A.enmWinding.WN
                    End If

                    If UN.Checked = False And VN.Checked = False And WN.Checked = False Then
                        winding = clsAccessXWRM10A.enmWinding.NA
                    End If

                Case 3 ' HV_WINDING = 4

                    windingSelection = clsAccessXWRM10A.enmWindingSelection.HV_WINDING
                    winding = clsAccessXWRM10A.enmWinding.NA

                Case 4 ' LV_WINDING = 5

                    windingSelection = clsAccessXWRM10A.enmWindingSelection.LV_WINDING
                    winding = clsAccessXWRM10A.enmWinding.NA

            End Select

            Dim intHVCurrent As Double
            Dim intLVCurrent As Double

            Select Case cmbHVCurrent.Text
                Case "1 mA = 1"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.ONE_mA

                Case "10 mA = 2"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.TEN_mA

                Case "100 mA = 3"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.HUNDREAD_mA

                Case "1 A = 4"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.ONE_A

                Case "5 A = 5"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.FIVE_A

                Case "10 A = 6"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.TEN_A

                Case "25 A = 7"
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.TWENTYFINE_A

                Case ""
                    intHVCurrent = clsAccessXWRM10A.enmHVCurrent.NA

            End Select

            Select Case cmbLVCurrent.Text
                Case "1 mA = 1"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.ONE_mA

                Case "10 mA = 2"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.TEN_mA

                Case "100 mA = 3"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.HUNDREAD_mA

                Case "1 A = 4"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.ONE_A

                Case "5 A = 5"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.FIVE_A

                Case "10 A = 6"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.TEN_A

                Case "25 A = 7"
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.TWENTYFINE_A

                Case ""
                    intLVCurrent = clsAccessXWRM10A.enmLVCurrent.NA

            End Select

            If cmbTestMode.SelectedIndex = 1 Then 'HR Mode

                'HR Mode
                Select Case cmbHVCurrent.SelectedIndex
                    Case 0 '1 mA
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '20kΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_kΩ
                            Case 1 '2kΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_KΩ
                            Case 2 '200 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_Ω
                            Case 3 '20 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                        End Select

                    Case 1 '10 mA
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '2kΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_KΩ
                            Case 1 '200 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_Ω
                            Case 2 '20 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                            Case 3 '2 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                        End Select

                    Case 2 '100 mA
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '200 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_Ω
                            Case 1 '20 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                            Case 2 '2 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 3 '200 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                        End Select

                    Case 3 '1 A
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '20 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                            Case 1 '2 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 2 '200 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                            Case 3 '20 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_mΩ
                        End Select

                    Case 4 '5 A
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '4 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.FOUR_Ω
                            Case 1 '400 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.FOUR_HUNDREAD_mΩ
                            Case 2 '40 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.FOURTY_mΩ
                            Case 3 '4 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.FOUR_mΩ
                        End Select

                    Case 5 '10 A
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '2 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 1 '200 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                            Case 2 '20 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_mΩ
                            Case 3 '2 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_mΩ
                        End Select

                    Case 6 '25 A
                        Select Case cmbHVRange.SelectedIndex
                            Case 0 '2 Ω
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 1 '200 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                            Case 2 '20 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWENTY_mΩ
                            Case 3 '2 mΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_mΩ
                            Case 4 '200 µΩ
                                objHVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_uΩ
                        End Select

                End Select
            Else
                ' Non-HR mode → AutoRange
                objHVRange = clsAccessXWRM10A.enmRange.AUTORANGE
            End If


            If cmbTestMode.SelectedIndex = 1 Then 'HR

                Select Case cmbLVCurrent.SelectedIndex
                    Case 0 '1 mA
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '20kΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_kΩ
                            Case 1 '2kΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_KΩ
                            Case 2 '200 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_Ω
                            Case 3 '20 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                        End Select

                    Case 1 '10 mA
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '2kΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_KΩ
                            Case 1 '200 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_Ω
                            Case 2 '20 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                            Case 3 '2 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                        End Select
                    Case 2 '100 mA
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '200 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_Ω
                            Case 1 '20 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                            Case 2 '2 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 3 '200 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                        End Select
                    Case 3 '1 A
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '20 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_Ω
                            Case 1 '2 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 2 '200 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                            Case 3 '20 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_mΩ
                        End Select

                    Case 4 '5 A
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '4 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.FOUR_Ω
                            Case 1 '400 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.FOUR_HUNDREAD_mΩ
                            Case 2 '40 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.FOURTY_mΩ
                            Case 3 '4 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.FOUR_mΩ
                        End Select
                    Case 5 '10 A
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '2 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 1 '200 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                            Case 2 '20 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_mΩ
                            Case 3 '2 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_mΩ
                        End Select
                    Case 6 '25 A
                        Select Case cmbLVRange.SelectedIndex
                            Case 0 '2 Ω
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_Ω
                            Case 1 '200 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_mΩ
                            Case 2 '20 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWENTY_mΩ
                            Case 3 '2 mΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_mΩ
                            Case 4 '200 µΩ
                                objLVRange = clsAccessXWRM10A.enmRange.TWO_HUNDREAD_uΩ
                        End Select
                End Select
            Else
                ' Non-HR mode → AutoRange
                objLVRange = clsAccessXWRM10A.enmRange.AUTORANGE
            End If

            Dim blnNormalorCorrected As Boolean = False
            objAccessXWRM10A.SetParameters(objTestmode, vectorGroup, vectorGroupYNYN, windingSelection, winding, intHVCurrent, intLVCurrent, intHrTime, intTotalTap, intCurrentTap, objHVRange, objLVRange, testObject, DUTSpecification, DUTSerialNumber, DUTMake, DUTTestLocation, year, Temp1Loc, Temp2Loc)
            System.Threading.Thread.Sleep(1000)
            If ResponseRecieved("Set Parameters") = False Then Exit Sub
            If sender IsNot Nothing Then
                'frmMessagebox = New frmMessageBox("Test Parameters Set Successfully.", True, frmMessagebox.MessageStyle.INFO_MESSAGE)
                'frmMessagebox.ShowDialog()
                MsgBox("Test Parameters Set Successfully.", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            ShowError()
        End Try
    End Sub

    Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles UN.CheckedChanged, VN.CheckedChanged, WN.CheckedChanged
        Dim maxAllowed As Integer = 0

        Select Case cmbWindingSelection.SelectedIndex + 1
            Case 1 ' ALL WINDING
                maxAllowed = 3
            Case 2 ' DUAL WINDING
                maxAllowed = 2
            Case 3 ' SINGLE
                maxAllowed = 1
        End Select

        ' Count how many are currently checked
        Dim countChecked As Integer = 0
        If UN.Checked Then countChecked += 1
        If VN.Checked Then countChecked += 1
        If WN.Checked Then countChecked += 1

        ' If over limit, uncheck the last one (the one user just clicked)
        If countChecked > maxAllowed Then
            Dim chk As CheckBox = CType(sender, CheckBox)
            chk.Checked = False
            MessageBox.Show("You can select only " & maxAllowed & " winding(s).", "Limit Reached", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Private Sub cmbWindingSelection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbWindingSelection.SelectedIndexChanged
        ' Reset all checkboxes when selection changes
        UN.Checked = False
        VN.Checked = False
        WN.Checked = False
    End Sub
    Public Sub ShowError()
        Dim stStackTrace As StackTrace = New StackTrace()
        MessageBox.Show(Err.Description, stStackTrace.GetFrame(1).GetMethod.ReflectedType.Name & ": " & stStackTrace.GetFrame(1).GetMethod.Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Function ResponseRecieved(ByVal Command As String, Optional ByVal ShowMessage As Boolean = True) As Boolean
        If objAccessXWRM10A.ResponseReceived = False Then

            'If ShowMessage Then
            '    Dim frmMessagebox As New frmMessageBox("Response Not recived for" & Command)
            '    frmMessagebox.ShowDialog()
            'End If

            Return False
        End If
        Return True
    End Function

    'Private Sub cmbHVSelection_SelectionChangeCommitted(sender As Object, e As EventArgs)

    '    Select Case cmbHVSelection.SelectedIndex
    '        Case 0

    '            Label3.Visible = True
    '            cmbHVCurrent.Visible = True
    '        Case 1

    '            Label3.Visible = False
    '            cmbHVCurrent.Visible = False
    '    End Select

    'End Sub


    'Private Sub cmbLVSelection_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Select Case cmbLVSelection.SelectedIndex
    '        Case 0

    '            Label4.Visible = True
    '            cmbLVCurrent.Visible = True
    '        Case 1

    '            Label4.Visible = False
    '            cmbLVCurrent.Visible = False
    '    End Select
    'End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        objAccessXWRM10A.TesON()
        tmrTestON.Start()

        'If cmbTestMode.SelectedIndex = 2 Then
        '    Dim f As New frmOLTC
        '    f.ShowDialog()
        'End If
        'dgvHRdata.AllowUserToAddRows = False
        'dgvHRdata.RowHeadersVisible = False
        'dgvHRdata.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'dgvHRdata.Columns.Add("colTime", "Time (Sec)")
        'dgvHRdata.Columns.Add("colUN", "UN")
        'dgvHRdata.Columns.Add("colun", "un")
        'dgvHRdata.Columns.Add("colTemp1", "Temp1 (°C)")
        'dgvHRdata.Columns.Add("colTemp2", "Temp2 (°C)")
        'Me.Hide()
        'Dim f As New frmTestForm
        'f.ShowDialog()
        'Me.Close()

        'dgvOLTC.AllowUserToAddRows = False
        'dgvOLTC.RowHeadersVisible = False
        'dgvOLTC.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'dgvOLTC.Columns.Add("colTap", "Tap No")
        'dgvOLTC.Columns.Add("colUN", "UN")
        'dgvOLTC.Columns.Add("colvn", "VN")
        'dgvOLTC.Columns.Add("colwn", "WN")


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        objAccessXWRM10A.TestOff()
    End Sub

    Private Sub cmbHVCurrent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHVCurrent.SelectedIndexChanged


        Select Case cmbHVCurrent.SelectedIndex
            'Case 0, 1, 2, 3, 5, 6

            '    cmbHVRange.Items.Clear()
            '    cmbHVRange.Items.Add("2KΩ = 1")
            '    cmbHVRange.Items.Add("200 Ω = 2")
            '    cmbHVRange.Items.Add("20 Ω = 3")
            '    cmbHVRange.Items.Add("2 Ω = 4")
            '    cmbHVRange.Items.Add("200 mΩ = 5")
            '    cmbHVRange.Items.Add("20 mΩ = 6")
            '    cmbHVRange.Items.Add("2 mΩ = 7")
            '    cmbHVRange.Items.Add("200 uΩ =8")


            '    cmbHVRange.Items.Clear()
            '    cmbHVRange.Items.Add("4 Ω = 9")
            '    cmbHVRange.Items.Add("400 mΩ = 10")
            '    cmbHVRange.Items.Add("40 mΩ = 11")
            '    cmbHVRange.Items.Add("4 mΩ = 12")
            Case 0
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("20kΩ = 0")
                cmbHVRange.Items.Add("2kΩ = 1")
                cmbHVRange.Items.Add("200Ω = 2")
                cmbHVRange.Items.Add("20Ω = 3")


            Case 1
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("2kΩ = 1")
                cmbHVRange.Items.Add("200Ω = 2")
                cmbHVRange.Items.Add("20Ω = 3")
                cmbHVRange.Items.Add("2Ω = 4")


            Case 2
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("200Ω = 2")
                cmbHVRange.Items.Add("20Ω = 3")
                cmbHVRange.Items.Add("2Ω = 4")
                cmbHVRange.Items.Add("200mΩ = 5")


            Case 3
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("20Ω = 3")
                cmbHVRange.Items.Add("2Ω = 4")
                cmbHVRange.Items.Add("200mΩ = 5")
                cmbHVRange.Items.Add("20mΩ = 6")


            Case 4
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("4Ω = 9")
                cmbHVRange.Items.Add("400mΩ = 10")
                cmbHVRange.Items.Add("40mΩ = 11")
                cmbHVRange.Items.Add("4mΩ = 12")


            Case 5
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("2Ω = 4")
                cmbHVRange.Items.Add("200mΩ = 5")
                cmbHVRange.Items.Add("20mΩ = 6")
                cmbHVRange.Items.Add("2mΩ = 7")


            Case 6
                cmbHVRange.Items.Clear()
                cmbHVRange.Items.Add("2Ω = 4")
                cmbHVRange.Items.Add("200mΩ = 5")
                cmbHVRange.Items.Add("20mΩ = 6")
                cmbHVRange.Items.Add("2mΩ = 7")
                cmbHVRange.Items.Add("200uΩ = 8")


        End Select
    End Sub

    Private Sub cmbLVCurrent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLVCurrent.SelectedIndexChanged
        Select Case cmbLVCurrent.SelectedIndex
            'Case 0, 1, 2, 3, 5, 6

            '    cmbLVRange.Items.Clear()
            '    cmbLVRange.Items.Add("2KΩ = 1")
            '    cmbLVRange.Items.Add("200 Ω = 2")
            '    cmbLVRange.Items.Add("20 Ω = 3")
            '    cmbLVRange.Items.Add("2 Ω = 4")
            '    cmbLVRange.Items.Add("200 mΩ = 5")
            '    cmbLVRange.Items.Add("20 mΩ = 6")
            '    cmbLVRange.Items.Add("2 mΩ = 7")
            '    cmbLVRange.Items.Add("200 uΩ =8")

            'Case 4
            '    cmbLVRange.Items.Clear()
            '    cmbLVRange.Items.Add("4 Ω = 9")
            '    cmbLVRange.Items.Add("400 mΩ = 10")
            '    cmbLVRange.Items.Add("40 mΩ = 11")
            '    cmbLVRange.Items.Add("4 mΩ = 12")

            Case 0
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("20kΩ = 0")
                cmbLVRange.Items.Add("2kΩ = 1")
                cmbLVRange.Items.Add("200Ω = 2")
                cmbLVRange.Items.Add("20Ω = 3")


            Case 1
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("2kΩ = 1")
                cmbLVRange.Items.Add("200Ω = 2")
                cmbLVRange.Items.Add("20Ω = 3")
                cmbLVRange.Items.Add("2Ω = 4")


            Case 2
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("200Ω = 2")
                cmbLVRange.Items.Add("20Ω = 3")
                cmbLVRange.Items.Add("2Ω = 4")
                cmbLVRange.Items.Add("200mΩ = 5")


            Case 3
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("20Ω = 3")
                cmbLVRange.Items.Add("2Ω = 4")
                cmbLVRange.Items.Add("200mΩ = 5")
                cmbLVRange.Items.Add("20mΩ = 6")


            Case 4
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("4Ω = 9")
                cmbLVRange.Items.Add("400mΩ = 10")
                cmbLVRange.Items.Add("40mΩ = 11")
                cmbLVRange.Items.Add("4mΩ = 12")


            Case 5
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("2Ω = 4")
                cmbLVRange.Items.Add("200mΩ = 5")
                cmbLVRange.Items.Add("20mΩ = 6")
                cmbLVRange.Items.Add("2mΩ = 7")


            Case 6
                cmbLVRange.Items.Clear()
                cmbLVRange.Items.Add("2Ω = 4")
                cmbLVRange.Items.Add("200mΩ = 5")
                cmbLVRange.Items.Add("20mΩ = 6")
                cmbLVRange.Items.Add("2mΩ = 7")
                cmbLVRange.Items.Add("200uΩ = 8")

        End Select
    End Sub

    Private Sub cmbHVCurrent_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbHVCurrent.SelectionChangeCommitted

        If cmbTestMode.SelectedIndex = 2 Then 'OLTC
            cmbLVCurrent.Enabled = False
        Else
            cmbLVCurrent.Enabled = True
        End If

    End Sub

    Private Sub cmbLVCurrent_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbLVCurrent.SelectionChangeCommitted

        If cmbTestMode.SelectedIndex = 2 Then 'OLTC
            cmbHVCurrent.Enabled = False
        Else
            cmbHVCurrent.Enabled = True
        End If


    End Sub

    Private Sub cmbVectorGroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbVectorGroup.SelectedIndexChanged

        Select Case cmbVectorGroup.SelectedIndex
            Case 0
                rdAuto.Visible = True
                rdGen.Visible = True

            Case 1, 2, 3
                rdAuto.Visible = False
                rdGen.Visible = False

        End Select

    End Sub

    Private Sub cmbHVRange_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHVRange.SelectedIndexChanged

    End Sub

    Private Sub txtSerialNo_TextChanged(sender As Object, e As EventArgs) Handles txtSerialNo.TextChanged

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label20_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label28_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label29_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label32_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label42_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label41_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label40_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label44_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub InitializeOLTCGrid(tapCount As Integer)
        totalTaps = tapCount
        currentReadingIndex = 0

        With dgvOLTC
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .Columns.Clear()
            .Rows.Clear()

            .Columns.Add("colTap", "Tap No")
            .Columns.Add("colUN", "UN")
            .Columns.Add("colVN", "VN")
            .Columns.Add("colWN", "WN")

            ' Create rows based on total taps
            For i As Integer = 1 To totalTaps
                .Rows.Add(i, "", "", "")
            Next
        End With
    End Sub
    Dim readingValue As Double
    Dim totalTap As Integer
    Dim varcolUN As Boolean
    Dim varcolVN As Boolean
    Dim varcolWN As Boolean

    Private Sub tnrTestON_Tick(sender As Object, e As EventArgs) Handles tmrTestON.Tick
        Label49.Text = Now.Date

        Label44.Text = DateAndTime.Now.ToString("t")
        totalTap = CInt(txtTotalTap.Text)
        objResponse = objAccessXWRM10A.GetResponse

        lblHVCurrent.Text = objResponse.HVTestCurrent
        lblLVCurrent.Text = objResponse.LVTestCurrent

        'NEW CHANGES

        Label31.Text = objResponse.CURRENT
        TextBox4.Text = objResponse.TapPos
        '''''''''''
        ' === Ensure columns exist (run once, e.g. in Form_Load) ===
        With dgvOLTC
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            If .Columns.Count = 0 Then
                .Columns.Add("colTap", "Tap No")
                .Columns.Add("colUN", "UN")
                .Columns.Add("colvn", "VN")
                .Columns.Add("colwn", "WN")
            End If
        End With

        ' === Add or update row ===
        Dim existingRow As DataGridViewRow = Nothing

        For Each row As DataGridViewRow In dgvOLTC.Rows
            If Not row.IsNewRow AndAlso row.Cells("colTap").Value IsNot Nothing AndAlso
       CInt(row.Cells("colTap").Value) = objResponse.TapPos Then
                existingRow = row
                Exit For
            End If
        Next

        ''If existingRow IsNot Nothing Then

        ''    If existingRow.Index <= totalTap Then

        ''        existingRow.Cells("colUN").Value = objResponse.READING
        ''    End If

        ''    If existingRow.Index <= totalTap * 2 Then

        ''        existingRow.Cells("colvn").Value = objResponse.READING
        ''    End If
        ''    If existingRow.Index <= totalTap * 3 Then

        ''        existingRow.Cells("colwn").Value = objResponse.READING
        ''    End If
        ''Else
        ''    ' ✅ Add new row for new tap number
        ''    Dim rowIndex As Integer = dgvOLTC.Rows.Add()
        ''    With dgvOLTC.Rows(rowIndex)
        ''        .Cells("colTap").Value = objResponse.TapNumber
        ''        If rowIndex <= totalTap Then

        ''            .Cells("colUN").Value = objResponse.READING
        ''        End If
        ''        If rowIndex <= totalTap Then

        ''            .Cells("colvn").Value = objResponse.READING
        ''        End If
        ''        If rowIndex <= totalTap Then

        ''            .Cells("colwn").Value = objResponse.READING
        ''        End If
        ''    End With
        ''End If

        'If existingRow IsNot Nothing Then

        '    If existingRow.Index <= objResponse.TapPos Then

        '        existingRow.Cells("colUN").Value = objResponse.READING
        '    ElseIf txtTotalTap.Text = objResponse.TapPos Then

        '        existingRow.Cells("colvn").Value = objResponse.READING
        '    ElseIf existingRow.Index <= objResponse.TapPos * 3 Then

        '        existingRow.Cells("colwn").Value = objResponse.READING
        '    End If
        'Else

        '    Dim rowIndex As Integer = dgvOLTC.Rows.Add()
        '    With dgvOLTC.Rows(rowIndex)
        '        .Cells("colTap").Value = objResponse.TapPos
        '        If rowIndex <= totalTap Then

        '            .Cells("colUN").Value = objResponse.READING
        '        ElseIf txtTotalTap.Text >= objResponse.TapPos Then

        '            .Cells("colvn").Value = objResponse.READING
        '        ElseIf txtTotalTap.Text <= objResponse.TapPos Then

        '            .Cells("colwn").Value = objResponse.READING
        '        End If
        '    End With
        'End If


        If existingRow IsNot Nothing Then

            If varcolUN = False Then

                existingRow.Cells("colUN").Value = objResponse.READING
                If objResponse.TapPos = txtTotalTap.Text Then
                    varcolUN = True
                End If
            End If
            If varcolUN = True Then

                existingRow.Cells("colvn").Value = objResponse.READING
                If objResponse.TapPos = 1 Then
                    varcolVN = True
                End If
            End If
            If varcolVN = True Then

                existingRow.Cells("colwn").Value = objResponse.READING

            End If
        Else

            Dim rowIndex As Integer = dgvOLTC.Rows.Add()
            With dgvOLTC.Rows(rowIndex)
                .Cells("colTap").Value = objResponse.TapPos
                If rowIndex <= totalTap Then

                    .Cells("colUN").Value = objResponse.READING
                ElseIf txtTotalTap.Text >= objResponse.TapPos Then

                    .Cells("colvn").Value = objResponse.READING
                ElseIf txtTotalTap.Text <= objResponse.TapPos Then

                    .Cells("colwn").Value = objResponse.READING
                End If
            End With
        End If
        '''''''''''''''

        'If totalTaps = 0 Then Exit Sub ' make sure initialized

        'currentReadingIndex += 1

        'Dim colName As String = ""
        'Dim rowIndex As Integer = 0

        '' --- Determine which column and row to fill ---
        'If currentReadingIndex <= totalTaps Then
        '    ' 1st pass → UN (forward)
        '    colName = "colUN"
        '    rowIndex = currentReadingIndex - 1
        'ElseIf currentReadingIndex <= totalTaps * 2 Then
        '    ' 2nd pass → VN (reverse)
        '    colName = "colVN"
        '    rowIndex = (totalTaps * 2) - currentReadingIndex
        'ElseIf currentReadingIndex <= totalTaps * 3 Then
        '    ' 3rd pass → WN (forward)
        '    colName = "colWN"
        '    rowIndex = currentReadingIndex - (totalTaps * 2) - 1
        'Else
        '    ' All readings filled
        '    MessageBox.Show("All readings complete!")
        '    Exit Sub
        'End If

        '' --- Fill the correct cell ---
        'dgvOLTC.Rows(rowIndex).Cells(colName).Value = readingValue


        ' === Ensure columns exist (only once, e.g. Form_Load) ===
        'With dgvOLTC
        '    .AllowUserToAddRows = False
        '    .RowHeadersVisible = False
        '    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        '    If .Columns.Count = 0 Then
        '        .Columns.Add("colTap", "Tap No")
        '        .Columns.Add("colUN", "UN")
        '        .Columns.Add("colvn", "VN")
        '        .Columns.Add("colwn", "WN")
        '    End If
        'End With

        '' === Ensure rows exist (1 to totalTaps) ===
        'If dgvOLTC.Rows.Count < totalTaps Then
        '    dgvOLTC.Rows.Clear()
        '    For i As Integer = 1 To totalTaps
        '        dgvOLTC.Rows.Add(i, Nothing, Nothing, Nothing)
        '        dgvOLTC.Rows(i - 1).Cells("colTap").Value = i
        '    Next
        'End If

        '' === Reading control variables ===
        'Static currentReadingIndex As Integer = 0
        'Static currentPass As Integer = 1 ' 1 = UN, 2 = VN, 3 = WN

        'If totalTaps < 0 Then Exit Sub

        '' Increment reading index each time a reading arrives
        'currentReadingIndex += 1

        '' Determine which column to fill
        'Dim colName As String = ""
        'Select Case currentPass
        '    Case 1 : colName = "colUN"  ' UN pass (forward)
        '    Case 2 : colName = "colvn"  ' VN pass (reverse)
        '    Case 3 : colName = "colwn"  ' WN pass (forward)
        '    Case Else
        '        MessageBox.Show("All readings complete!")
        '        Exit Sub
        'End Select

        '' Determine row index for current reading
        'Dim rowIndex As Integer

        'If currentPass = 1 Then
        '    ' UN forward 1 → totalTaps
        '    rowIndex = currentReadingIndex - 1
        'ElseIf currentPass = 2 Then
        '    ' VN reverse totalTaps → 1
        '    rowIndex = totalTaps - currentReadingIndex
        'ElseIf currentPass = 3 Then
        '    ' WN forward 1 → totalTaps
        '    rowIndex = currentReadingIndex - 1
        'End If

        '' Assign the reading value
        'dgvOLTC.Rows(rowIndex).Cells(colName).Value = objResponse.READING

        '' --- When current pass is complete, reset index and advance ---
        'If currentPass = 1 AndAlso currentReadingIndex >= totalTaps Then
        '    currentPass = 2
        '    currentReadingIndex = 0
        'ElseIf currentPass = 2 AndAlso currentReadingIndex >= totalTaps Then
        '    currentPass = 3
        '    currentReadingIndex = 0
        'ElseIf currentPass = 3 AndAlso currentReadingIndex >= totalTaps Then
        '    MessageBox.Show("✅ All readings complete!")
        'End If



        ''''''''''''
        Select Case objResponse.HVTestCurrentUnit
            Case objAccessXWRM10A.enmCurrentsUnit.Ampere
                lblHVCurrent.Text = FormatCurrentReading(Val(lblHVCurrent.Text)) & " A"
            Case objAccessXWRM10A.enmCurrentsUnit.mAmpere
                lblHVCurrent.Text = FormatCurrentReading(Val(lblHVCurrent.Text)) & " mA"
        End Select

        Select Case objResponse.LVTestCurrentUnit
            Case objAccessXWRM10A.enmCurrentsUnit.Ampere
                lblLVCurrent.Text = FormatCurrentReading(Val(lblLVCurrent.Text)) & " A"
            Case objAccessXWRM10A.enmCurrentsUnit.mAmpere
                lblLVCurrent.Text = FormatCurrentReading(Val(lblLVCurrent.Text)) & " mA"
        End Select

        Select Case objResponse.Winding

            Case objAccessXWRM10A.enmWinding.UN

                lblHV1Reading.Text = FormatChannelReading(objResponse.HVReading)

                Select Case objResponse.HVUnit
                    Case objAccessXWRM10A.enmResistanceUnit.KOhm
                        lblHV1Reading.Text = lblHV1Reading.Text + " kOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.mOhm
                        lblHV1Reading.Text = lblHV1Reading.Text + " mOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.Ohm
                        lblHV1Reading.Text = lblHV1Reading.Text + " Ohm"

                    Case objAccessXWRM10A.enmResistanceUnit.uOhm
                        lblHV1Reading.Text = lblHV1Reading.Text + " uOhm"

                End Select
                'End If


                lblLV1Reading.Text = FormatChannelReading(objResponse.LVReading)

                Select Case objResponse.LVUnit
                    Case objAccessXWRM10A.enmResistanceUnit.KOhm
                        lblLV1Reading.Text = lblLV1Reading.Text + " kOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.mOhm
                        lblLV1Reading.Text = lblLV1Reading.Text + " mOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.Ohm
                        lblLV1Reading.Text = lblLV1Reading.Text + " Ohm"

                    Case objAccessXWRM10A.enmResistanceUnit.uOhm
                        lblLV1Reading.Text = lblLV1Reading.Text + " uOhm"

                End Select


            Case objAccessXWRM10A.enmWinding.VN

                lblHV2Reading.Text = FormatChannelReading(objResponse.HVReading)

                Select Case objResponse.HVUnit
                    Case objAccessXWRM10A.enmResistanceUnit.KOhm
                        lblHV2Reading.Text = lblHV2Reading.Text + " kOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.mOhm
                        lblHV2Reading.Text = lblHV2Reading.Text + " mOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.Ohm
                        lblHV2Reading.Text = lblHV2Reading.Text + " Ohm"

                    Case objAccessXWRM10A.enmResistanceUnit.uOhm
                        lblHV2Reading.Text = lblHV2Reading.Text + " uOhm"

                End Select
                'End If


                lblLV2Reading.Text = FormatChannelReading(objResponse.LVReading)

                Select Case objResponse.LVUnit
                    Case objAccessXWRM10A.enmResistanceUnit.KOhm
                        lblLV2Reading.Text = lblLV2Reading.Text + " kOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.mOhm
                        lblLV2Reading.Text = lblLV2Reading.Text + " mOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.Ohm
                        lblLV2Reading.Text = lblLV2Reading.Text + " Ohm"

                    Case objAccessXWRM10A.enmResistanceUnit.uOhm
                        lblLV2Reading.Text = lblLV2Reading.Text + " uOhm"

                End Select


            Case objAccessXWRM10A.enmWinding.WN

                lblHV3Reading.Text = FormatChannelReading(objResponse.HVReading)

                Select Case objResponse.HVUnit
                    Case objAccessXWRM10A.enmResistanceUnit.KOhm
                        lblHV3Reading.Text = lblHV3Reading.Text + " kOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.mOhm
                        lblHV3Reading.Text = lblHV3Reading.Text + " mOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.Ohm
                        lblHV3Reading.Text = lblHV3Reading.Text + " Ohm"

                    Case objAccessXWRM10A.enmResistanceUnit.uOhm
                        lblHV3Reading.Text = lblHV3Reading.Text + " uOhm"

                End Select
                'End If


                lblLV3Reading.Text = FormatChannelReading(objResponse.LVReading)

                Select Case objResponse.LVUnit
                    Case objAccessXWRM10A.enmResistanceUnit.KOhm
                        lblLV3Reading.Text = lblLV3Reading.Text + " kOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.mOhm
                        lblLV3Reading.Text = lblLV3Reading.Text + " mOhm"

                    Case objAccessXWRM10A.enmResistanceUnit.Ohm
                        lblLV3Reading.Text = lblLV3Reading.Text + " Ohm"

                    Case objAccessXWRM10A.enmResistanceUnit.uOhm
                        lblLV3Reading.Text = lblLV3Reading.Text + " uOhm"

                End Select


        End Select

        lblTemp1.Text = objResponse.T1Reading & " °C"

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

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        objAccessXWRM10A.TestOff()
    End Sub


    Private Sub RoundPanelCorners(ByVal pnl As Panel, ByVal radius As Integer)
        Dim path As New GraphicsPath()
        path.StartFigure()

        ' Top Left Corner
        path.AddArc(New Rectangle(0, 0, radius, radius), 180, 90)

        ' Top Edge
        path.AddLine(radius, 0, pnl.Width - radius, 0)

        ' Top Right Corner
        path.AddArc(New Rectangle(pnl.Width - radius, 0, radius, radius), 270, 90)

        ' Right Edge
        path.AddLine(pnl.Width, radius, pnl.Width, pnl.Height - radius)

        ' Bottom Right Corner
        path.AddArc(New Rectangle(pnl.Width - radius, pnl.Height - radius, radius, radius), 0, 90)

        ' Bottom Edge
        path.AddLine(pnl.Width - radius, pnl.Height, radius, pnl.Height)

        ' Bottom Left Corner
        path.AddArc(New Rectangle(0, pnl.Height - radius, radius, radius), 90, 90)

        ' Left Edge
        path.AddLine(0, pnl.Height - radius, 0, radius)

        path.CloseFigure()

        pnl.Region = New Region(path)
    End Sub


    Private Sub RoundFormCorners(ByVal frm As Form, ByVal radius As Integer)
        Dim path As New GraphicsPath()
        path.StartFigure()

        ' Top Left
        path.AddArc(New Rectangle(0, 0, radius, radius), 180, 90)
        path.AddLine(radius, 0, frm.Width - radius, 0)

        ' Top Right
        path.AddArc(New Rectangle(frm.Width - radius, 0, radius, radius), 270, 90)
        path.AddLine(frm.Width, radius, frm.Width, frm.Height - radius)

        ' Bottom Right
        path.AddArc(New Rectangle(frm.Width - radius, frm.Height - radius, radius, radius), 0, 90)
        path.AddLine(frm.Width - radius, frm.Height, radius, frm.Height)

        ' Bottom Left
        path.AddArc(New Rectangle(0, frm.Height - radius, radius, radius), 90, 90)
        path.AddLine(0, frm.Height - radius, 0, radius)

        path.CloseFigure()

        frm.Region = New Region(path)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Dim rowIndex As Integer = dgvHRdata.Rows.Add()
        'With dgvHRdata.Rows(rowIndex)
        '    .Cells("colTap").Value = objResponse.TapNumber
        '    .Cells("colUN").Value = objResponse.READING
        '    'If objResponse.TapNumber > VarTapNO Then
        '    '    Exit Sub
        '    'End If
        '    .Cells("colvn").Value = objResponse.READING
        '    .Cells("colwn").Value = objResponse.READING

        'End With
        ' Assume objResponse.TapNumber and objResponse.READING are available

        ' First, check if the tap number already exists in the grid
        ' Dim existingRow As DataGridViewRow = Nothing

        ' For Each row As DataGridViewRow In dgvOLTC.Rows
        '     If Not row.IsNewRow AndAlso row.Cells("colTap").Value IsNot Nothing AndAlso
        'CInt(row.Cells("colTap").Value) = objResponse.TapNumber Then
        '         existingRow = row
        '         Exit For
        '     End If
        ' Next

        ' If existingRow IsNot Nothing Then
        '     ' ✅ Update existing row
        '     existingRow.Cells("colUN").Value = objResponse.READING
        '     existingRow.Cells("colvn").Value = objResponse.READING
        '     existingRow.Cells("colwn").Value = objResponse.READING
        ' Else
        '     ' ✅ Add new row for new tap number
        '     Dim rowIndex As Integer = dgvOLTC.Rows.Add()
        '     With dgvOLTC.Rows(rowIndex)
        '         .Cells("colTap").Value = objResponse.TapNumber
        '         .Cells("colUN").Value = objResponse.READING
        '         .Cells("colvn").Value = objResponse.READING
        '         .Cells("colwn").Value = objResponse.READING
        '     End With
        ' End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        objAccessXWRM10A.Home()
    End Sub
End Class
