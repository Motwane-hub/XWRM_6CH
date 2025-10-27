'DVM 06 NOV 2014  New Class Created for XWRM10AData Instrument.
Imports System.ComponentModel
Imports System.Reflection
Imports Microsoft.VisualBasic.Conversion
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Win32.SafeHandles
Imports AccessXWRM10A.clsAccessXWRM10A20



Public Class clsAccessXWRM10A20 : Implements IDisposable
    Private strCOM As String = vbNullString
    Private WithEvents objComPort As IO.Ports.SerialPort

    Private blnRegistered As Boolean = False

    Public _ResponseResult As New Response

    Public Event ResponseRecieved(ByVal objSettingResponse As Response)
    Public Event ReportDataRecieved(ByVal dtReportData As DataTable)
    Public Event SendResponseFrame(ByVal strString As String)
    Public Event SendCommandFrame(ByVal strString As String)

    Public Event ShowGraphData()

    Public LastCommand As enmCommands

    Public _IsCommunicated As Boolean = False

    Public Property IsCommunicated() As Boolean
        Get
            Return _IsCommunicated
        End Get
        Set(ByVal value As Boolean)
            _IsCommunicated = value
        End Set
    End Property
    Public ReadOnly Property liveCurrent() As Double
        Get
            Return liveCur
        End Get
    End Property
    Dim liveCur As Double

#Region "Enum"

    Public Property GetResponse() As Response
        Get
            Return _ResponseResult
        End Get
        Set(ByVal value As Response)
            _ResponseResult = value
        End Set
    End Property

    Enum enmCommands
        SETTING = &H1
        TEST_ON = &H2
        TEST_OFF = &H3
        DATE_SYNC = &H4
        GET_SETTING = &H5
        GET_STATUS = &H6
        GET_REPORT_COUNT = &H7
        GET_ALL_REPORT = &H8
        DEMAGNETIZE = &H9
        HOME = &H10
        SAVE = &H11
        PRINT_COMMAND = &H12
        CLEAR = &H13
        READ_DATE_TIME = &H14
        RECALL = &H15
        GET_DEVICE_TEST_HISTORY = &H16
    End Enum

    Enum enmErrorMessage
        Ser_Wrong_CRC = &HA1
        Ser_Wrong_Opcode = &HA2
        Ser_Wrong_Cmndcode = &HA3
        SUCCESS = 0
    End Enum

    Enum enmCurrentsUnit
        mAmpere = &H0
        Ampere = &H1
    End Enum

    Enum enmResistanceUnit
        uOhm = &H0
        mOhm = &H1
        Ohm = &H2
        KOhm = &H3
    End Enum

    Enum enmTempratureUnit
        degreeCelcius = &H0
    End Enum

    Enum enmDecimalPoint
        DP1 = &H0
        DP2 = &H1
        DP3 = &H2
    End Enum

    Enum enmTestStatus
        TEST_ON = 0
        TEST_OFF = 1
    End Enum

    Enum enmChannels
        ALL_CH_OFF = 0
        CH1_ON = 1
        CH2_ON = 2
        CH3_ON = 4
        CH1_CH2_ON = 3
        CH2_CH3_ON = 6
        CH1_CH3_ON = 5
        ALL_CH_ON = 7
    End Enum

    Enum enmRange
        HUNDREAD = 1
        TEN = 2
        ONE = 3
        DIVAND1 = 4
    End Enum

    Enum enmTestMode
        OLTC = 1
        HR = 2
        NORMAL = 4
    End Enum

    Enum enmWindingElement
        NO_ELEMENT = 0
        COPPER = 1
        ALUMINIUM = 2
    End Enum

#End Region

#Region "Properties"

    Public ReadOnly Property IsRegistered() As Boolean
        Get
            Return blnRegistered
        End Get
    End Property

    '--------Connection--------
    Public ReadOnly Property IsConnected() As Boolean 'checking for com connection
        Get
            If objComPort.IsOpen = True Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

#End Region

#Region "Public Function and Subroutines"

    'Public Sub New()
    '    MyBase.New()

    '    blnRegistered = False

    '    Dim frmRegistration As New Registration()
    '    blnRegistered = frmRegistration.IsRegistered()

    '    objComPort = New Ports.SerialPort
    'End Sub

    Dim dtNewTable As New Data.DataTable

    Public Function Connect(ByVal ComPort As String) As Boolean
        Try
            dtNewTable = New Data.DataTable
            dtNewTable.Columns.Add("Current", System.Type.GetType("System.Double"))        '  dtNewTable.Columns.Add("Unit") 
            dtNewTable.Columns.Add("Elapsed_Time", System.Type.GetType("System.Double"))
            dtNewTable.Columns.Add("Count")

            intSecond = 0

            '   If Not CheckRegistration() Then Return False

            If objComPort.IsOpen = False Then
                objComPort = Nothing
                objComPort = New IO.Ports.SerialPort
                objComPort.PortName = ComPort
                objComPort.BaudRate = 230400 '921600 '230400
                objComPort.DataBits = 8
                objComPort.Parity = Ports.Parity.None
                objComPort.StopBits = Ports.StopBits.One

                AddHandler objComPort.DataReceived, AddressOf objComPort_DataReceived

                objComPort.Open()
                objComPort.DiscardInBuffer()
                objComPort.DiscardOutBuffer()
            Else
                Throw New ApplicationException("ER001")
            End If

            Return objComPort.IsOpen
        Catch ex As Exception
            '  MsgBox(Err.Description)
        End Try
    End Function

    Public Sub Disconnect()
        Try
            If objComPort.IsOpen Then
                objComPort.DiscardInBuffer()
                objComPort.Close()
            Else
                Throw New ApplicationException("ER003")
            End If
        Catch ex As Exception
            '   Throw New ApplicationException("ER004")
        End Try
    End Sub

    Public Sub SetParameters(ByVal current As Double, ByVal currentUnit As enmCurrentsUnit, ByVal channel As enmChannels, ByVal ch1Range As enmRange, ByVal ch2Range As enmRange, ByVal ch3Range As enmRange, ByVal testMode As enmTestMode, ByVal blnSaveTest As Boolean, ByVal HrTestTime As Integer, ByVal DUTSerialNumber As String, ByVal DUTType As String, ByVal DUTTestLocation As String, ByVal HighestTap As Integer)
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.SETTING).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            'Current Value
            Dim strCurrent As String = ConvertSingleToHex(current)
            strCurrent = strCurrent.Insert(0, "#")
            strCurrent = strCurrent.Insert(3, "#")
            strCurrent = strCurrent.Insert(6, "#")
            strCurrent = strCurrent.Insert(9, "#")

            'Current Unit
            strCommand = strCommand & strCurrent

            If currentUnit = enmCurrentsUnit.Ampere Then
                strCommand = strCommand & "#01"
            End If

            If currentUnit = enmCurrentsUnit.mAmpere Then
                strCommand = strCommand & "#00"
            End If

            Select Case channel
                Case enmChannels.ALL_CH_OFF
                    strCommand = strCommand & "#00"

                Case enmChannels.CH1_ON
                    strCommand = strCommand & "#01"

                Case enmChannels.CH2_ON
                    strCommand = strCommand & "#02"

                Case enmChannels.CH3_ON
                    strCommand = strCommand & "#04"

                Case enmChannels.CH1_CH2_ON
                    strCommand = strCommand & "#03"

                Case enmChannels.CH2_CH3_ON
                    strCommand = strCommand & "#06"

                Case enmChannels.CH1_CH3_ON
                    strCommand = strCommand & "#05"

                Case enmChannels.ALL_CH_ON
                    strCommand = strCommand & "#07"
            End Select

            Select Case ch1Range
                Case enmRange.HUNDREAD
                    strCommand = strCommand & "#00"

                Case enmRange.TEN
                    strCommand = strCommand & "#01"

                Case enmRange.ONE
                    strCommand = strCommand & "#02"

                Case enmRange.DIVAND1
                    strCommand = strCommand & "#04"
            End Select

            Dim strBinary1 As String = ""

            Select Case ch2Range
                Case enmRange.HUNDREAD
                    strBinary1 = "0000"

                Case enmRange.TEN
                    strBinary1 = "0001"

                Case enmRange.ONE
                    strBinary1 = "0010"

                Case enmRange.DIVAND1
                    strBinary1 = "0100"
            End Select

            Dim strBinary2 As String = ""

            Select Case ch3Range
                Case enmRange.HUNDREAD
                    strBinary2 = "0000"

                Case enmRange.TEN
                    strBinary2 = "0001"

                Case enmRange.ONE
                    strBinary2 = "0010"

                Case enmRange.DIVAND1
                    strBinary2 = "0100"
            End Select

            Dim strBinary As String = strBinary1 & strBinary2
            strBinary = BinaryToHex(strBinary).PadLeft(2, "0")

            strCommand = strCommand & "#" & strBinary

            Select Case testMode
                Case enmTestMode.NORMAL
                    strBinary1 = "0100"

                Case enmTestMode.OLTC
                    strBinary1 = "0001"

                Case enmTestMode.HR
                    strBinary1 = "0010"
            End Select

            If blnSaveTest Then
                strBinary2 = "0001"
            Else
                strBinary2 = "0000"
            End If

            strBinary = strBinary1 & strBinary2
            strBinary = BinaryToHex(strBinary).PadLeft(2, "0")

            strCommand = strCommand & "#" & strBinary

            Dim strHRTestTime As String = "#" & Hex(HrTestTime).PadLeft(2, "0")

            Select Case testMode
                Case enmTestMode.NORMAL
                    strHRTestTime = "#" & Hex(HrTestTime).PadLeft(2, "0")

                Case enmTestMode.OLTC
                    Dim strBinary5 As String = ToBinary(CDec("&h" & Hex(HrTestTime).PadLeft(2, "0")))
                    Dim strNewByte As String = Val(HighestTap).ToString() & strBinary5.Substring(1, 7)

                    strHRTestTime = "#" & BinaryToHex(strNewByte).PadLeft(2, "0")

                Case enmTestMode.HR
                    strHRTestTime = "#" & Hex(HrTestTime).PadLeft(2, "0")

            End Select

            strCommand = strCommand & strHRTestTime

            ''DUT Serial Number
            Dim strDUTSerialNumber As String = vbNullString

            If DUTSerialNumber.Length > 15 Then
                DUTSerialNumber = DUTSerialNumber.Substring(0, 15)
            End If

            For Each strChar As Char In DUTSerialNumber
                strDUTSerialNumber = strDUTSerialNumber & "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            If strDUTSerialNumber = vbNullString Then
                strDUTSerialNumber = "#00"
            End If

            Do While strDUTSerialNumber.Length < 45
                strDUTSerialNumber = strDUTSerialNumber & "#00"
            Loop

            strCommand = strCommand & strDUTSerialNumber

            ''DUT Type
            Dim strDUTType As String = vbNullString

            If DUTType.Length > 15 Then
                DUTType = DUTType.Substring(0, 15)
            End If

            For Each strChar As Char In DUTType
                strDUTType = strDUTType & "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            If strDUTType = vbNullString Then
                strDUTType = "#00"
            End If

            Do While strDUTType.Length < 45
                strDUTType = strDUTType & "#00"
            Loop

            strCommand = strCommand & strDUTType

            ''DUT Test Location
            Dim strDUTTestLocation As String = vbNullString

            If DUTTestLocation.Length > 15 Then
                DUTTestLocation = DUTTestLocation.Substring(0, 15)
            End If

            For Each strChar As Char In DUTTestLocation
                strDUTTestLocation = strDUTTestLocation & "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            If strDUTTestLocation = vbNullString Then
                strDUTTestLocation = "#00"
            End If

            Do While strDUTTestLocation.Length < 45
                strDUTTestLocation = strDUTTestLocation & "#00"
            Loop

            strCommand = strCommand & strDUTTestLocation
            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(68) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("SetParameters " & Err.Description)
        End Try
    End Sub

    Public Sub TesON()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.TEST_ON).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Test ON " & Err.Description)
        End Try
    End Sub

    Public Sub TestOff()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.TEST_OFF).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Test OFF " & Err.Description)
        End Try
    End Sub

    Public Sub ClearBuffer()
        Try
            objComPort.DiscardInBuffer()
            objComPort.DiscardOutBuffer()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SetDateTime(ByVal DateTime As DateTime)
        Try
            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                           'SOF                '1 Byte
            strCommand = strCommand & "#10"                                             'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.DATE_SYNC).PadLeft(2, "0")   'Op Code '1 Byte
            strCommand = strCommand & "#C0"

            Dim strDateTime As String = ConvertDateToHex(DateTime)
            strCommand = strCommand & strDateTime

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("SetDateTime " & Err.Description)
        End Try
    End Sub

    Public Sub GetSetting()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.GET_SETTING).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("GetSetting " & Err.Description)
        End Try
    End Sub

    Public Sub GetStatus()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.GET_STATUS).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("GetStatus " & Err.Description)
        End Try
    End Sub

    Public Sub GetReportCount()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.GET_REPORT_COUNT).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("GetReportCount " & Err.Description)
        End Try
    End Sub

    Public Sub GetAllReports()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.GET_ALL_REPORT).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("GetAllReports " & Err.Description)
        End Try
    End Sub

    Public Sub DeMagentize(ByVal current As Double, ByVal currentUnit As enmCurrentsUnit)
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.DEMAGNETIZE).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            'Current Value
            Dim strCurrent As String = ConvertSingleToHex(current)
            strCurrent = strCurrent.Insert(0, "#")
            strCurrent = strCurrent.Insert(3, "#")
            strCurrent = strCurrent.Insert(6, "#")
            strCurrent = strCurrent.Insert(9, "#")

            'Current Unit
            strCommand = strCommand & strCurrent

            If currentUnit = enmCurrentsUnit.Ampere Then
                strCommand = strCommand & "#01"
            End If

            If currentUnit = enmCurrentsUnit.mAmpere Then
                strCommand = strCommand & "#00"
            End If

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("DeMagentize " & Err.Description)
        End Try
    End Sub

    Public Sub Home()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.HOME).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Home " & Err.Description)
        End Try
    End Sub

    Public Sub Save()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.SAVE).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Save " & Err.Description)
        End Try
    End Sub

    Public Sub Print(ByVal ReportNumber As Integer)
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.PRINT_COMMAND).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"
            'Command '1 Byte 

            strCommand = strCommand & "#" & Hex(ReportNumber).PadLeft(4, "0").Insert(2, "#")
            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Print " & Err.Description)
        End Try
    End Sub

    Public Sub Clear()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.CLEAR).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Clear " & Err.Description)
        End Try
    End Sub

    Public Sub GetDateTime()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.READ_DATE_TIME).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("GetDateTime " & Err.Description)
        End Try
    End Sub

    Public Sub Recall()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.RECALL).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Recall " & Err.Description)
        End Try
    End Sub

    Public Sub GetDeviceTestHistory()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.GET_DEVICE_TEST_HISTORY).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)

            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("GetDeviceTestHistory " & Err.Description)
        End Try
    End Sub

#End Region

#Region "Private Function and Subroutines"

#Region "Send Frame"

    'Dim strResponseString As New Text.StringBuilder
    Dim tempStringBuilder As New Text.StringBuilder

    'Public Property ResetGraphData() As String
    '    Get
    '        Return strResponseString.ToString
    '    End Get
    '    Set(ByVal value As String)
    '        dtNewTable = New Data.DataTable
    '        dtNewTable.Columns.Add("Current", System.Type.GetType("System.Double"))        '  dtNewTable.Columns.Add("Unit") 
    '        dtNewTable.Columns.Add("Elapsed_Time", System.Type.GetType("System.Double"))
    '        dtNewTable.Columns.Add("Count")
    '        intSecond = 0

    '        strResponseString = New Text.StringBuilder
    '    End Set
    'End Property

    Public ReadOnly Property GetGraphTable() As DataTable
        Get
            Return dtNewTable
        End Get
    End Property

    'Dim intCounter As Integer = 0
    'Dim intTempCounter As Integer = 1

    Dim drNewRow As Data.DataRow
    Dim intSecond As Double = 0
    Dim strByte As String = ""

    Dim blnReset As Boolean = False

    Private Sub objComPort_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles objComPort.DataReceived
        Try
            strByte = Hex(objComPort.ReadByte()).PadLeft(2, "0")

            If strByte.StartsWith("3A") Then
                tempStringBuilder.Append(strByte)

                Do While tempStringBuilder.ToString.Length < 24
                    tempStringBuilder.Append(Hex(objComPort.ReadByte()).PadLeft(2, "0"))

                Loop

                Dim myString As String
                myString = tempStringBuilder.ToString
                File.AppendAllText("test.txt", myString + vbCrLf)
                tempStringBuilder = New Text.StringBuilder

                Dim strData As String = myString(2) & myString(3) & myString(4) & myString(5) & myString(6) & myString(7) & myString(8) & myString(9)
                liveCur = Math.Round(ConvertHexToSingle(strData), 4)
                RaiseEvent ShowGraphData()
            End If
            'While objComPort.BytesToRead > 0
            '    strByte = Hex(objComPort.ReadByte()).PadLeft(2, "0")

            '    If strByte.StartsWith("3A") Then
            '        tempStringBuilder.Append(strByte)

            '        Do While tempStringBuilder.ToString.Length < 24
            '            tempStringBuilder.Append(Hex(objComPort.ReadByte()).PadLeft(2, "0"))
            '        Loop

            '        Dim myString As String
            '        myString = tempStringBuilder.ToString
            '        tempStringBuilder = New Text.StringBuilder

            '        Dim strData As String = myString(2) & myString(3) & myString(4) & myString(5) & myString(6) & myString(7) & myString(8) & myString(9)

            '        drNewRow = dtNewTable.NewRow
            '        drNewRow.Item("Current") = Math.Round(ConvertHexToSingle(strData), 4)

            '        If Val(drNewRow.Item("Current")) > 0 Then
            '            strData = myString(12) & myString(13) & myString(14) & myString(15) & myString(16) & myString(17) & myString(18) & myString(19)

            '            drNewRow.Item("Count") = CDec("&h" & strData)

            '            'intSecond = intSecond + 0.006
            '            intSecond = intSecond + 0.0013
            '            drNewRow.Item("Elapsed_Time") = intSecond

            '            dtNewTable.Rows.Add(drNewRow)
            '        Else
            '            strData = ""
            '        End If

            '        If dtNewTable.Rows.Count Mod 2000 = 0 Then

            '            If blnReset Then
            '                RaiseEvent ShowGraphData(dtNewTable)
            '            Else
            '                blnReset = True
            '            End If

            '            dtNewTable.Rows.Clear()

            '        End If
            '    End If

            'End While

            'If objComPort.BytesToRead = 0 Then
            '    RaiseEvent ShowGraphData(dtNewTable)
            'End If            
        Catch ex As Exception
            'MsgBox(Err.Description)
        End Try
    End Sub

    Public Sub ProcessResponseFrame(ByVal btData As Byte())
        Try
            Dim strResponseString As String = vbNullString

            For Each btbyte In btData
                strResponseString = strResponseString & "#" & Hex(btbyte).PadLeft(2, "0")
            Next

            RaiseEvent SendResponseFrame(strResponseString)

            Dim strBinary As String = vbNullString
            Dim strData As String = vbNullString

            _ResponseResult.Command = CDec("&h" & Hex(btData(2)).PadLeft(2, "0"))

            Select Case btData.Length
                Case 49 'Normal

                    If _ResponseResult.Command = enmCommands.READ_DATE_TIME Then
                        strData = Hex(btData(4)).PadLeft(2, "0") & "#" & Hex(btData(5)).PadLeft(2, "0") & "#" & Hex(btData(6)).PadLeft(2, "0") & "#" & Hex(btData(7)).PadLeft(2, "0") & "#" & Hex(btData(8)).PadLeft(2, "0") & "#" & Hex(btData(9)).PadLeft(2, "0") & "#" & Hex(btData(10)).PadLeft(2, "0")
                        _ResponseResult.DeviceDateTime = GetDateTime(strData) 'Holds real time clock (RTC) data	7   Hex(btData(51)).PadLeft(2, "0")
                    End If

                    If _ResponseResult.Command = enmCommands.GET_REPORT_COUNT Then
                        strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0")
                        _ResponseResult.TotalReportCount = CDec("&h" & strData)

                        strData = Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")
                        _ResponseResult.TotalReadingCount = CDec("&h" & strData)
                    End If

                    If _ResponseResult.Command = enmCommands.GET_DEVICE_TEST_HISTORY Then
                        strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0")
                        _ResponseResult.InstrumentReportCount = CDec("&h" & strData)

                        strData = Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")
                        _ResponseResult.OperationalHours = CDec("&h" & strData)
                    End If

                    _ResponseResult.Message = CDec("&h" & Hex(btData(46)).PadLeft(2, "0"))

                Case 70 'Setting
                    strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0") & Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")

                    _ResponseResult.CurrentValue = ConvertHexToSingle(strData)

                    strBinary = ToBinary(CDec("&h" & Hex(btData(8)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CurrentUnit = enmCurrentsUnit.mAmpere
                            _ResponseResult.CurrentValue = _ResponseResult.CurrentValue * 1000

                            _ResponseResult.CurrentValue = Math.Round(_ResponseResult.CurrentValue)

                        Case "0001"
                            _ResponseResult.CurrentUnit = enmCurrentsUnit.Ampere
                    End Select

                    _ResponseResult.ChannelSelected = CDec("&h" & Hex(btData(9)).PadLeft(2, "0"))

                    strBinary = ToBinary(CDec("&h" & Hex(btData(10)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CH1Range = enmRange.HUNDREAD
                        Case "0001"
                            _ResponseResult.CH1Range = enmRange.TEN
                        Case "0010"
                            _ResponseResult.CH1Range = enmRange.ONE
                        Case "0100"
                            _ResponseResult.CH1Range = enmRange.DIVAND1
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(11)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            _ResponseResult.CH2Range = enmRange.HUNDREAD
                        Case "0001"
                            _ResponseResult.CH2Range = enmRange.TEN
                        Case "0010"
                            _ResponseResult.CH2Range = enmRange.ONE
                        Case "0100"
                            _ResponseResult.CH2Range = enmRange.DIVAND1
                    End Select

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CH3Range = enmRange.HUNDREAD
                        Case "0001"
                            _ResponseResult.CH3Range = enmRange.TEN
                        Case "0010"
                            _ResponseResult.CH3Range = enmRange.ONE
                        Case "0100"
                            _ResponseResult.CH3Range = enmRange.DIVAND1
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(12)).PadLeft(2, "0")))

                    _ResponseResult.IsHighest = False

                    Select Case strBinary.Substring(0, 4)
                        Case "0100"
                            _ResponseResult.TestModeSeeting = enmTestMode.NORMAL
                            _ResponseResult.HrTestIime = CDec("&h" & Hex(btData(13)).PadLeft(2, "0"))

                        Case "0001"
                            _ResponseResult.TestModeSeeting = enmTestMode.OLTC
                            Dim strBinary20 As String = ToBinary(CDec("&h" & Hex(btData(13)).PadLeft(2, "0")))

                            _ResponseResult.HrTestIime = BinaryToDecimal("0" & strBinary20.Substring(1, 7))

                            If strBinary20.Substring(0, 1) = 1 Then
                                _ResponseResult.IsHighest = True
                            Else
                                _ResponseResult.IsHighest = False
                            End If

                        Case "0010"
                            _ResponseResult.TestModeSeeting = enmTestMode.HR
                            _ResponseResult.HrTestIime = CDec("&h" & Hex(btData(13)).PadLeft(2, "0"))
                    End Select

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.SaveResult = False

                        Case "0001"
                            _ResponseResult.SaveResult = True
                    End Select


                    _ResponseResult.DUTSerialNumber = vbNullString

                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(14)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(15)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(16)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(17)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(18)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(19)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(20)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(21)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(22)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(23)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(24)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(25)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(26)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(27)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(28)).PadLeft(2, "0"))

                    _ResponseResult.DUTType = vbNullString

                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(29)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(30)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(31)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(32)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(33)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(34)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(35)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(36)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(37)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(38)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(39)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(40)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(41)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(42)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(43)).PadLeft(2, "0"))

                    _ResponseResult.DUTLocation = vbNullString

                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(44)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(45)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(46)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(47)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(48)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(49)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(50)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(51)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(52)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(53)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(54)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(55)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(56)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(57)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(58)).PadLeft(2, "0"))

                    _ResponseResult.Message = CDec("&h" & Hex(btData(67)).PadLeft(2, "0"))

                Case 92 'Status
                    _ResponseResult.TestModeStatus = CDec("&h" & Hex(btData(4)).PadLeft(2, "0"))
                    strBinary = ToBinary(CDec("&h" & Hex(btData(4)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(0, 4)
                        Case "0100"
                            _ResponseResult.TestModeStatus = enmTestMode.NORMAL

                        Case "0001"
                            _ResponseResult.TestModeStatus = enmTestMode.OLTC

                        Case "0010"
                            _ResponseResult.TestModeStatus = enmTestMode.HR
                    End Select

                    _ResponseResult.TapNumber = CDec("&h" & Hex(btData(5)).PadLeft(2, "0"))

                    strData = Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0") & Hex(btData(8)).PadLeft(2, "0") & Hex(btData(9)).PadLeft(2, "0")
                    _ResponseResult.HVTestCurrent = ConvertHexToSingle(strData)

                    strBinary = ToBinary(CDec("&h" & Hex(btData(10)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CurrentUnit = enmCurrentsUnit.mAmpere
                            _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.mAmpere
                            _ResponseResult.HVTestCurrent = _ResponseResult.HVTestCurrent * 1000
                        Case "0001"
                            _ResponseResult.CurrentUnit = enmCurrentsUnit.Ampere
                            _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.Ampere
                    End Select

                    Select Case _ResponseResult.HVTestCurrent
                        Case Is < 1
                            _ResponseResult.HVTestCurrent = Math.Round(_ResponseResult.HVTestCurrent, 4)
                        Case Is < 10
                            _ResponseResult.HVTestCurrent = Math.Round(_ResponseResult.HVTestCurrent, 3)
                        Case Is < 100
                            _ResponseResult.HVTestCurrent = Math.Round(_ResponseResult.HVTestCurrent, 2)
                        Case Is < 1000
                            _ResponseResult.HVTestCurrent = Math.Round(_ResponseResult.HVTestCurrent, 1)
                        Case Is < 10000
                            _ResponseResult.HVTestCurrent = Math.Round(_ResponseResult.HVTestCurrent, 0)
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            _ResponseResult.TestCurrentDP = enmDecimalPoint.DP1
                        Case "0001"
                            _ResponseResult.TestCurrentDP = enmDecimalPoint.DP2
                        Case "0010"
                            _ResponseResult.TestCurrentDP = enmDecimalPoint.DP3
                    End Select

                    'Channel 1 reading and unit
                    strData = Hex(btData(11)).PadLeft(2, "0") & Hex(btData(12)).PadLeft(2, "0") & Hex(btData(13)).PadLeft(2, "0") & Hex(btData(14)).PadLeft(2, "0")
                    _ResponseResult.CH1Reading = ConvertHexToSingle(strData)

                    Select Case _ResponseResult.CH1Reading
                        Case Is < 1
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 4)
                        Case Is < 10
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 3)
                        Case Is < 100
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 2)
                        Case Is < 1000
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 1)
                        Case Is < 10000
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(15)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CH1Unit = enmResistanceUnit.uOhm

                        Case "0100"
                            _ResponseResult.CH1Unit = enmResistanceUnit.mOhm

                        Case "1000"
                            _ResponseResult.CH1Unit = enmResistanceUnit.Ohm

                        Case "1100"
                            _ResponseResult.CH1Unit = enmResistanceUnit.KOhm
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            _ResponseResult.CH1DP = enmDecimalPoint.DP1
                        Case "0001"
                            _ResponseResult.CH1DP = enmDecimalPoint.DP2
                        Case "0010"
                            _ResponseResult.CH1DP = enmDecimalPoint.DP3
                    End Select

                    'Channel 2 Reading and Unit
                    strData = Hex(btData(16)).PadLeft(2, "0") & Hex(btData(17)).PadLeft(2, "0") & Hex(btData(18)).PadLeft(2, "0") & Hex(btData(19)).PadLeft(2, "0")
                    _ResponseResult.CH2Reading = ConvertHexToSingle(strData)

                    Select Case _ResponseResult.CH2Reading
                        Case Is < 1
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 4)
                        Case Is < 10
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 3)
                        Case Is < 100
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 2)
                        Case Is < 1000
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 1)
                        Case Is < 10000
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(20)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CH2Unit = enmResistanceUnit.uOhm

                        Case "0100"
                            _ResponseResult.CH2Unit = enmResistanceUnit.mOhm

                        Case "1000"
                            _ResponseResult.CH2Unit = enmResistanceUnit.Ohm

                        Case "1100"
                            _ResponseResult.CH2Unit = enmResistanceUnit.KOhm
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            _ResponseResult.CH2DP = enmDecimalPoint.DP1
                        Case "0001"
                            _ResponseResult.CH2DP = enmDecimalPoint.DP2
                        Case "0010"
                            _ResponseResult.CH2DP = enmDecimalPoint.DP3
                    End Select

                    'Channel 3 reading and unit
                    strData = Hex(btData(21)).PadLeft(2, "0") & Hex(btData(22)).PadLeft(2, "0") & Hex(btData(23)).PadLeft(2, "0") & Hex(btData(24)).PadLeft(2, "0")
                    _ResponseResult.CH3Reading = ConvertHexToSingle(strData)

                    Select Case _ResponseResult.CH3Reading
                        Case Is < 1
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 4)
                        Case Is < 10
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 3)
                        Case Is < 100
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 2)
                        Case Is < 1000
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 1)
                        Case Is < 10000
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(25)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CH3Unit = enmResistanceUnit.uOhm

                        Case "0100"
                            _ResponseResult.CH3Unit = enmResistanceUnit.mOhm

                        Case "1000"
                            _ResponseResult.CH3Unit = enmResistanceUnit.Ohm

                        Case "1100"
                            _ResponseResult.CH3Unit = enmResistanceUnit.KOhm
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            _ResponseResult.CH3DP = enmDecimalPoint.DP1
                        Case "0001"
                            _ResponseResult.CH3DP = enmDecimalPoint.DP2
                        Case "0010"
                            _ResponseResult.CH3DP = enmDecimalPoint.DP3
                    End Select

                    'T1 reading and unit
                    strData = Hex(btData(26)).PadLeft(2, "0") & Hex(btData(27)).PadLeft(2, "0") & Hex(btData(28)).PadLeft(2, "0") & Hex(btData(29)).PadLeft(2, "0")
                    _ResponseResult.T1Reading = ConvertHexToSingle(strData)

                    Select Case _ResponseResult.T1Reading
                        Case Is < 1
                            _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 4)
                        Case Is < 10
                            _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 3)
                        Case Is < 100
                            _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 2)
                        Case Is < 1000
                            _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 1)
                        Case Is < 10000
                            _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(30)).PadLeft(2, "0")))

                    _ResponseResult.T1Unit = enmTempratureUnit.degreeCelcius

                    Select Case strBinary.Substring(0, 4)
                        Case "1000"
                            _ResponseResult.T1DP = enmDecimalPoint.DP1
                        Case "1001"
                            _ResponseResult.T1DP = enmDecimalPoint.DP2
                        Case "1010"
                            _ResponseResult.T1DP = enmDecimalPoint.DP3
                    End Select

                    'T2 reading and unit
                    strData = Hex(btData(31)).PadLeft(2, "0") & Hex(btData(32)).PadLeft(2, "0") & Hex(btData(33)).PadLeft(2, "0") & Hex(btData(34)).PadLeft(2, "0")
                    _ResponseResult.T2Reading = ConvertHexToSingle(strData)

                    Select Case _ResponseResult.T2Reading
                        Case Is < 1
                            _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 4)
                        Case Is < 10
                            _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 3)
                        Case Is < 100
                            _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 2)
                        Case Is < 1000
                            _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 1)
                        Case Is < 10000
                            _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btData(35)).PadLeft(2, "0")))

                    _ResponseResult.T2Unit = enmTempratureUnit.degreeCelcius

                    Select Case strBinary.Substring(0, 4)
                        Case "1000"
                            _ResponseResult.T2DP = enmDecimalPoint.DP1

                        Case "1001"
                            _ResponseResult.T2DP = enmDecimalPoint.DP2

                        Case "1010"
                            _ResponseResult.T2DP = enmDecimalPoint.DP3
                    End Select

                    _ResponseResult.DUTSerialNumber = vbNullString

                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(36)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(37)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(38)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(39)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(40)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(41)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(42)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(43)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(44)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(45)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(46)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(47)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(48)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(49)).PadLeft(2, "0"))
                    _ResponseResult.DUTSerialNumber = _ResponseResult.DUTSerialNumber & ConvertHexToASCII(Hex(btData(50)).PadLeft(2, "0"))

                    _ResponseResult.DUTType = vbNullString

                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(51)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(52)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(53)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(54)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(55)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(56)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(57)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(58)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(59)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(60)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(61)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(62)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(63)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(64)).PadLeft(2, "0"))
                    _ResponseResult.DUTType = _ResponseResult.DUTType & ConvertHexToASCII(Hex(btData(65)).PadLeft(2, "0"))

                    _ResponseResult.DUTLocation = vbNullString

                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(66)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(67)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(68)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(69)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(70)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(71)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(72)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(73)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(74)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(75)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(76)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(77)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(78)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(79)).PadLeft(2, "0"))
                    _ResponseResult.DUTLocation = _ResponseResult.DUTLocation & ConvertHexToASCII(Hex(btData(80)).PadLeft(2, "0"))

                    '81 for Instrumnet Status
                    strBinary = ToBinary(CDec("&h" & Hex(btData(81)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(7, 1)
                        Case "0"
                            _ResponseResult.TestON = False
                        Case "1"
                            _ResponseResult.TestON = True
                    End Select

                    Select Case strBinary.Substring(6, 1)
                        Case "0"
                            _ResponseResult.CommunicationBreak = False
                        Case "1"
                            _ResponseResult.CommunicationBreak = True
                    End Select

                    Select Case strBinary.Substring(5, 1)
                        Case "0"
                            _ResponseResult.MemoryFull = False
                        Case "1"
                            _ResponseResult.MemoryFull = True
                    End Select

                    strData = Hex(btData(82)).PadLeft(2, "0") & Hex(btData(83)).PadLeft(2, "0")
                    _ResponseResult.TotalReportCount = CDec("&h" & strData)

                    strData = Hex(btData(84)).PadLeft(2, "0") & Hex(btData(85)).PadLeft(2, "0")
                    _ResponseResult.TotalReadingCount = CDec("&h" & strData)

                    _ResponseResult.Message = CDec("&h" & Hex(btData(89)).PadLeft(2, "0"))
            End Select

            RaiseEvent ResponseRecieved(_ResponseResult)
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub ProcessReportData(ByVal colCollection As Collection)
        Try
            Dim strResponseString As String = vbNullString

            For iSubCount As Integer = 1 To (colCollection.Count)
                strResponseString = strResponseString & "#" & Hex(colCollection.Item(iSubCount)).PadLeft(2, "0")
            Next

            RaiseEvent SendResponseFrame(strResponseString)

            Dim dtDataTable As New Data.DataTable

            dtDataTable.Columns.Add("TestReportNo", GetType(System.Double))
            dtDataTable.Columns.Add("ReportDateTime", GetType(System.DateTime))
            dtDataTable.Columns.Add("DUTSrNo", GetType(System.String))
            dtDataTable.Columns.Add("DUTType", GetType(System.String))
            dtDataTable.Columns.Add("DUTLocation", GetType(System.String))
            dtDataTable.Columns.Add("TestMode", GetType(System.String))
            dtDataTable.Columns.Add("ReadingCount", GetType(System.Double))

            dtDataTable.Columns.Add("TapNumber", GetType(System.Double))

            dtDataTable.Columns.Add("TestCurrentReading", GetType(System.Double))
            dtDataTable.Columns.Add("TestCurrentUnit", GetType(System.String))
            dtDataTable.Columns.Add("TestCurrentDP", GetType(System.Double))

            dtDataTable.Columns.Add("CH1Reading", GetType(System.Double))
            dtDataTable.Columns.Add("CH1Unit", GetType(System.String))
            dtDataTable.Columns.Add("CH1ReadingDP", GetType(System.Double))

            dtDataTable.Columns.Add("CH2Reading", GetType(System.Double))
            dtDataTable.Columns.Add("CH2Unit", GetType(System.String))
            dtDataTable.Columns.Add("CH2ReadingDP", GetType(System.Double))

            dtDataTable.Columns.Add("CH3Reading", GetType(System.Double))
            dtDataTable.Columns.Add("CH3Unit", GetType(System.String))
            dtDataTable.Columns.Add("CH3ReadingDP", GetType(System.Double))

            dtDataTable.Columns.Add("Temp1", GetType(System.Double))
            dtDataTable.Columns.Add("Temp1Unit", GetType(System.String))
            dtDataTable.Columns.Add("Temp1DP", GetType(System.Double))

            dtDataTable.Columns.Add("Temp2", GetType(System.Double))
            dtDataTable.Columns.Add("Temp2Unit", GetType(System.String))
            dtDataTable.Columns.Add("Temp2DP", GetType(System.Double))

            dtDataTable.Columns.Add("TestResultNo", GetType(System.Double))
            dtDataTable.Columns.Add("TestTimeInterval", GetType(System.Double))

            Dim blnFirstFarme As Boolean = True
            Dim intSerialNumber As Integer = 0
            Dim intCounter As Integer = 1

            Dim nextReportCount As Integer = 0

            Dim _TestReportNo As Double = 0
            Dim _DateTime As DateTime = System.DateTime.Now
            Dim _DUTSrNo As String = vbNullString
            Dim _DUTType As String = vbNullString
            Dim _DUTLocation As String = vbNullString
            Dim _TestMode As String = vbNullString
            Dim _ReadingCount As Double = 0

            Dim drNewRow As DataRow = dtDataTable.NewRow
            Dim strData As String = vbNullString

            Do While intCounter < colCollection.Count

                If nextReportCount = 0 Then
                    blnFirstFarme = True
                End If

                If blnFirstFarme Then
                    Dim btsubDataNew(69) As Byte

                    For iSubCount As Integer = 0 To 69
                        btsubDataNew(iSubCount) = colCollection.Item(intCounter)
                        intCounter = intCounter + 1
                    Next

                    strData = Hex(btsubDataNew(57)).PadLeft(2, "0") & Hex(btsubDataNew(58)).PadLeft(2, "0")
                    nextReportCount = CDec("&h" & strData)

                    _ReadingCount = nextReportCount - 1

                    strData = Hex(btsubDataNew(4)).PadLeft(2, "0") & Hex(btsubDataNew(5)).PadLeft(2, "0")
                    _TestReportNo = CDec("&h" & strData)

                    Dim strBinary As String = ToBinary(CDec("&h" & Hex(btsubDataNew(6)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(0, 4)
                        Case "0100"
                            _TestMode = "NORMAL"

                        Case "0001"
                            _TestMode = "OLTC"

                        Case "0010"
                            _TestMode = "HR"
                    End Select

                    Dim intDate As Integer = CDec("&H" & Hex(btsubDataNew(52)).PadLeft(2, "0"))
                    Dim intMonth As Integer = CDec("&H" & Hex(btsubDataNew(53)).PadLeft(2, "0"))
                    Dim intYear As Integer = CDec("&H" & Hex(btsubDataNew(54)).PadLeft(2, "0"))
                    Dim intHours As Integer = CDec("&H" & Hex(btsubDataNew(55)).PadLeft(2, "0"))
                    Dim intMinutes As Integer = CDec("&H" & Hex(btsubDataNew(56)).PadLeft(2, "0"))

                    Try
                        _DateTime = intDate.ToString() & " " & GetMonthName(intMonth) & " " & intYear & " " & intHours & ":" & intMinutes & ":" & 0
                    Catch ex As Exception
                        MsgBox("Date not in valid format for Report Number " & _TestReportNo)
                    End Try

                    _DUTSrNo = vbNullString

                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(7)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(8)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(9)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(10)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(11)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(12)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(13)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(14)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(15)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(16)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(17)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(18)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(19)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(20)).PadLeft(2, "0"))
                    _DUTSrNo = _DUTSrNo & ConvertHexToASCII(Hex(btsubDataNew(21)).PadLeft(2, "0"))

                    _DUTType = vbNullString

                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(22)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(23)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(24)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(25)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(26)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(27)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(28)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(29)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(30)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(31)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(32)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(33)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(34)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(35)).PadLeft(2, "0"))
                    _DUTType = _DUTType & ConvertHexToASCII(Hex(btsubDataNew(36)).PadLeft(2, "0"))

                    _DUTLocation = vbNullString

                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(37)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(38)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(39)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(40)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(41)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(42)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(43)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(44)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(45)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(46)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(47)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(48)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(49)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(50)).PadLeft(2, "0"))
                    _DUTLocation = _DUTLocation & ConvertHexToASCII(Hex(btsubDataNew(51)).PadLeft(2, "0"))

                    blnFirstFarme = False

                Else
                    Dim btsubDataNew(48) As Byte

                    For iSubCount As Integer = 0 To 48
                        btsubDataNew(iSubCount) = colCollection.Item(intCounter)
                        intCounter = intCounter + 1
                    Next

                    'Reading Row
                    drNewRow = dtDataTable.NewRow

                    drNewRow.Item("TestReportNo") = _TestReportNo
                    drNewRow.Item("ReportDateTime") = _DateTime
                    drNewRow.Item("DUTSrNo") = _DUTSrNo
                    drNewRow.Item("DUTType") = _DUTType
                    drNewRow.Item("DUTLocation") = _DUTLocation
                    drNewRow.Item("TestMode") = _TestMode
                    drNewRow.Item("ReadingCount") = _ReadingCount

                    drNewRow.Item("TapNumber") = CDec("&h" & Hex(btsubDataNew(5)).PadLeft(2, "0"))

                    strData = Hex(btsubDataNew(6)).PadLeft(2, "0") & Hex(btsubDataNew(7)).PadLeft(2, "0") & Hex(btsubDataNew(8)).PadLeft(2, "0") & Hex(btsubDataNew(9)).PadLeft(2, "0")

                    drNewRow.Item("TestCurrentReading") = ConvertHexToSingle(strData)

                    Dim strBinary As String = ToBinary(CDec("&h" & Hex(btsubDataNew(10)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            drNewRow.Item("TestCurrentUnit") = "mAmpere"
                            drNewRow.Item("TestCurrentReading") = Val(drNewRow.Item("TestCurrentReading")) * 1000
                        Case "0001"
                            drNewRow.Item("TestCurrentUnit") = "Ampere"
                    End Select

                    Select Case Val(drNewRow.Item("TestCurrentReading"))
                        Case Is < 1
                            drNewRow.Item("TestCurrentReading") = Math.Round(drNewRow.Item("TestCurrentReading"), 4)
                        Case Is < 10
                            drNewRow.Item("TestCurrentReading") = Math.Round(drNewRow.Item("TestCurrentReading"), 3)
                        Case Is < 100
                            drNewRow.Item("TestCurrentReading") = Math.Round(drNewRow.Item("TestCurrentReading"), 2)
                        Case Is < 1000
                            drNewRow.Item("TestCurrentReading") = Math.Round(drNewRow.Item("TestCurrentReading"), 1)
                        Case Is < 10000
                            drNewRow.Item("TestCurrentReading") = Math.Round(drNewRow.Item("TestCurrentReading"), 0)
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            drNewRow.Item("TestCurrentDP") = 1
                        Case "0001"
                            drNewRow.Item("TestCurrentDP") = 2
                        Case "0010"
                            drNewRow.Item("TestCurrentDP") = 3
                    End Select

                    'ch1 reading unit dp
                    strData = Hex(btsubDataNew(11)).PadLeft(2, "0") & Hex(btsubDataNew(12)).PadLeft(2, "0") & Hex(btsubDataNew(13)).PadLeft(2, "0") & Hex(btsubDataNew(14)).PadLeft(2, "0")
                    drNewRow.Item("CH1Reading") = ConvertHexToSingle(strData)

                    Select Case Val(drNewRow.Item("CH1Reading"))
                        Case Is < 1
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 4)
                        Case Is < 10
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 3)
                        Case Is < 100
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 2)
                        Case Is < 1000
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 1)
                        Case Is < 10000
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btsubDataNew(15)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            drNewRow.Item("CH1Unit") = "uOhm"

                        Case "0100"
                            drNewRow.Item("CH1Unit") = "mOhm"

                        Case "1000"
                            drNewRow.Item("CH1Unit") = "Ohm"

                        Case "1100"
                            drNewRow.Item("CH1Unit") = "K Ohm"
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            drNewRow.Item("CH1ReadingDP") = 1
                        Case "0001"
                            drNewRow.Item("CH1ReadingDP") = 2
                        Case "0010"
                            drNewRow.Item("CH1ReadingDP") = 3
                    End Select

                    'ch2 reading,unit,dp
                    strData = Hex(btsubDataNew(16)).PadLeft(2, "0") & Hex(btsubDataNew(17)).PadLeft(2, "0") & Hex(btsubDataNew(18)).PadLeft(2, "0") & Hex(btsubDataNew(19)).PadLeft(2, "0")
                    drNewRow.Item("CH2Reading") = ConvertHexToSingle(strData)

                    Select Case Val(drNewRow.Item("CH2Reading"))
                        Case Is < 1
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 4)
                        Case Is < 10
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 3)
                        Case Is < 100
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 2)
                        Case Is < 1000
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 1)
                        Case Is < 10000
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btsubDataNew(20)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            drNewRow.Item("CH2Unit") = "uOhm"

                        Case "0100"
                            drNewRow.Item("CH2Unit") = "mOhm"

                        Case "1000"
                            drNewRow.Item("CH2Unit") = "Ohm"

                        Case "1100"
                            drNewRow.Item("CH1Unit") = "K Ohm"
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            drNewRow.Item("CH2ReadingDP") = 1
                        Case "0001"
                            drNewRow.Item("CH2ReadingDP") = 2
                        Case "0010"
                            drNewRow.Item("CH2ReadingDP") = 3
                    End Select

                    'ch3 reading,unit,dp
                    strData = Hex(btsubDataNew(21)).PadLeft(2, "0") & Hex(btsubDataNew(22)).PadLeft(2, "0") & Hex(btsubDataNew(23)).PadLeft(2, "0") & Hex(btsubDataNew(24)).PadLeft(2, "0")
                    drNewRow.Item("CH3Reading") = ConvertHexToSingle(strData)

                    Select Case Val(drNewRow.Item("CH3Reading"))
                        Case Is < 1
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 4)
                        Case Is < 10
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 3)
                        Case Is < 100
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 2)
                        Case Is < 1000
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 1)
                        Case Is < 10000
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 0)
                    End Select

                    strBinary = ToBinary(CDec("&h" & Hex(btsubDataNew(25)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            drNewRow.Item("CH3Unit") = "uOhm"

                        Case "0100"
                            drNewRow.Item("CH3Unit") = "mOhm"

                        Case "1000"
                            drNewRow.Item("CH3Unit") = "Ohm"

                        Case "1100"
                            drNewRow.Item("CH3Unit") = "K Ohm"
                    End Select

                    Select Case strBinary.Substring(0, 4)
                        Case "0000"
                            drNewRow.Item("CH3ReadingDP") = 1
                        Case "0001"
                            drNewRow.Item("CH3ReadingDP") = 2
                        Case "0010"
                            drNewRow.Item("CH3ReadingDP") = 3
                    End Select

                    'temp1 reading,unit,dp
                    drNewRow.Item("Temp1") = 0
                    drNewRow.Item("Temp1Unit") = ""

                    'T1 reading and unit
                    strData = Hex(btsubDataNew(26)).PadLeft(2, "0") & Hex(btsubDataNew(27)).PadLeft(2, "0") & Hex(btsubDataNew(28)).PadLeft(2, "0") & Hex(btsubDataNew(29)).PadLeft(2, "0")
                    drNewRow.Item("Temp1") = ConvertHexToSingle(strData)

                    strBinary = ToBinary(CDec("&h" & Hex(btsubDataNew(30)).PadLeft(2, "0")))

                    drNewRow.Item("Temp1Unit") = "°C"

                    Select Case strBinary.Substring(0, 4)
                        Case "1000"
                            drNewRow.Item("Temp1DP") = 1
                        Case "1001"
                            drNewRow.Item("Temp1DP") = 2
                        Case "1010"
                            drNewRow.Item("Temp1DP") = 3
                    End Select

                    'T2 reading and unit
                    strData = Hex(btsubDataNew(31)).PadLeft(2, "0") & Hex(btsubDataNew(32)).PadLeft(2, "0") & Hex(btsubDataNew(33)).PadLeft(2, "0") & Hex(btsubDataNew(34)).PadLeft(2, "0")
                    drNewRow.Item("Temp2") = ConvertHexToSingle(strData)

                    strBinary = ToBinary(CDec("&h" & Hex(btsubDataNew(35)).PadLeft(2, "0")))

                    drNewRow.Item("Temp2Unit") = enmTempratureUnit.degreeCelcius

                    Select Case strBinary.Substring(0, 4)
                        Case "1000"
                            drNewRow.Item("Temp2DP") = 1

                        Case "1001"
                            drNewRow.Item("Temp2DP") = 2

                        Case "1010"
                            drNewRow.Item("Temp2DP") = 3
                    End Select

                    strData = Hex(btsubDataNew(36)).PadLeft(2, "0") & Hex(btsubDataNew(37)).PadLeft(2, "0")
                    drNewRow.Item("TestResultNo") = CDec("&h" & strData) - 1

                    strData = Hex(btsubDataNew(38)).PadLeft(2, "0") & Hex(btsubDataNew(39)).PadLeft(2, "0")
                    drNewRow.Item("TestTimeInterval") = CDec("&h" & strData)

                    strData = Hex(btsubDataNew(40)).PadLeft(2, "0") & Hex(btsubDataNew(41)).PadLeft(2, "0")
                    Dim intTeetReportNumber As Double = CDec("&h" & strData)

                    dtDataTable.Rows.Add(drNewRow)
                End If

                nextReportCount = nextReportCount - 1
            Loop

            RaiseEvent ReportDataRecieved(dtDataTable)

        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Function ToBinary(ByVal dec As Integer) As String
        If dec = 0 Then Return "00000000"

        Dim bin As Integer
        Dim output As String = ""
        While dec <> 0
            If dec Mod 2 = 0 Then
                bin = 0
            Else
                bin = 1
            End If
            dec = dec \ 2
            output = Convert.ToString(bin) & output
        End While
        If output Is Nothing Then
            Return "0"
        Else
            output = output.PadLeft(8, "0")
            Return output
        End If
    End Function

    Private Sub objComPort_ErrorReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialErrorReceivedEventArgs) Handles objComPort.ErrorReceived
        WriteLog("Error : objComPort_ErrorReceived()", Err.Description)
    End Sub

    Private Sub SendCommand()

    End Sub

    'Calculate CRC (16th Byte) - XOR of all bytes prior to this field.
    Private Function CalculateCRC(ByVal strCmd() As String) As String
        Dim btCRC As Integer
        Dim inti As Integer

        For inti = 0 To UBound(strCmd) - 2
            Dim intByte As Integer
            intByte = Integer.Parse(CDec("&H" & strCmd(inti)))
            btCRC = btCRC Xor intByte
        Next

        Return Hex(btCRC)
    End Function

    'Sends the Frame to Meter
    Private Sub SendFrame(ByVal strCommandFrame() As String)
        Dim inti, intReadingCount As Integer

        Try
            LastCommand = CDec("&h" & strCommandFrame(2))

            If objComPort.IsOpen = True Then
                Dim aByte(strCommandFrame.Length - 1) As Byte
                Dim newByte As Byte

                For inti = 0 To UBound(strCommandFrame)
                    intReadingCount = CDec("&H" & strCommandFrame(inti))

                    newByte = Byte.Parse(strCommandFrame(inti), System.Globalization.NumberStyles.HexNumber)
                    aByte(inti) = newByte
                Next

                objComPort.Write(aByte, 0, strCommandFrame.Length)
                objComPort.DiscardOutBuffer()
            Else
                Throw New ApplicationException("ER005")
            End If
        Catch ex As Exception
            objComPort.Close()
            MsgBox("SendFrame " & ex.Message)
            Throw New ApplicationException("ER006")
        End Try
    End Sub

    'Private Function CheckRegistration() As Boolean

    '    Try
    '        If blnRegistered Then Return True

    '        blnRegistered = False

    '        Dim frmRegistration As New Registration()
    '        frmRegistration.ShowDialog()

    '        If frmRegistration.IsRegistered = False Then
    '            MsgBox("Please Register Application", MsgBoxStyle.Critical)
    '            blnRegistered = False
    '            Return False
    '        Else
    '            blnRegistered = True
    '            Return True
    '        End If

    '        Return False
    '    Catch ex As Exception
    '        Return False
    '    End Try

    'End Function

#End Region

    Private Function ConvertCharToHex(ByVal _Char As String) As String
        Dim strHex As String = vbNullString

        strHex = Hex(AscW(_Char)).PadLeft(2, "0")

        Return strHex
    End Function

    Private Function ConvertHexToASCII(ByVal HexString As String) As String
        If HexString = "00" Then Return vbNullString

        Dim strResult As String = vbNullString

        strResult = ChrW(Convert.ToInt32(HexString, 16))

        Return strResult
    End Function

    Public Function ConvertHexToSingle(ByVal hexValue As String) As Single
        Dim bArray(3) As Byte

        Try
            '  If hexValue = "3DCCCCCD" Then Return 0.1

            Dim intInputIndex As Integer = 0
            Dim intOutputIndex As Integer = 0

            For intInputIndex = 0 To hexValue.Length - 1 Step 2
                bArray(intOutputIndex) = Byte.Parse(hexValue.Chars(intInputIndex) & hexValue.Chars(intInputIndex + 1), Globalization.NumberStyles.HexNumber)
                intOutputIndex += 1
            Next

            Array.Reverse(bArray)

        Catch ex As Exception
            WriteLog("ConvertHexToSingle", "The supplied hex value is either empty or in an incorrect format. Use the following format: 00000000")
        End Try

        Return Convert.ToDouble(BitConverter.ToSingle(bArray, 0))
    End Function

    Private Function ConvertSingleToHex(ByVal SngValue As Single) As String
        Dim tmpHex As String = ""

        Try
            Dim tmpBytes() As Byte

            tmpBytes = BitConverter.GetBytes(SngValue)

            For b As Integer = tmpBytes.GetUpperBound(0) To 0 Step -1
                If Hex(tmpBytes(b)).Length = 1 Then tmpHex += "0" '0..F -> 00..0F 
                tmpHex += Hex(tmpBytes(b))
            Next

        Catch ex As Exception
            WriteLog("ConvertSingleToHex", Err.Description)
        End Try

        Return tmpHex
    End Function

    Public Function BinaryToDecimal(ByVal Bin As String) 'function to convert a binary number to decimal
        Dim dec As Double = Nothing

        Try
            Dim length As Integer = Len(Bin)

            Dim temp As Integer = Nothing
            Dim x As Integer = Nothing

            For x = 1 To length
                temp = Val(Mid(Bin, length, 1))
                length = length - 1
                If temp <> "0" Then
                    dec += (2 ^ (x - 1))
                End If
            Next

        Catch ex As Exception
            WriteLog("BinaryToDecimal", Err.Description)
        End Try

        Return dec
    End Function

    Public Function BinaryToHex(ByVal bin As String) As String
        Return Hex(BinaryToDecimal(bin))
    End Function

    Private Function BCDtoDecimal(ByVal sBinary As String) As Integer
        Try
            Dim iNumber As Integer = Val(BinaryToDecimal(sBinary.Substring(0, 4).PadLeft(8, "0")) & BinaryToDecimal(sBinary.Substring(4, 4).PadLeft(8, "0")))
            Return iNumber
        Catch ex As Exception
            WriteLog("BCDtoDecimal", Err.Description)
        End Try
    End Function

    'Functions Realted to RTC (Real Time Clock)
    Private Function GetDateTime(ByVal HexString As String) As DateTime
        Try
            Dim dtDateTime As DateTime
            Dim strSplitArray As String() = HexString.Split("#")

            'Seconds, Minutes, Hours, Day, Date, Month and Year
            Dim intDate As Integer = CDec("&H" & strSplitArray(0))
            Dim intMonth As Integer = CDec("&H" & strSplitArray(1))
            Dim intYear As Integer = CDec("&H" & strSplitArray(2))
            Dim intDay As Integer = CDec("&H" & strSplitArray(3))
            Dim intHours As Integer = CDec("&H" & strSplitArray(4))
            Dim intMinutes As Integer = CDec("&H" & strSplitArray(5))
            Dim intSeconds As Integer = CDec("&H" & strSplitArray(6))

            dtDateTime = intDate.ToString() & " " & GetMonthName(intMonth) & " " & intYear & " " & intHours & ":" & intMinutes & ":" & intSeconds

            Return dtDateTime
        Catch ex As Exception
            Return "1 Jan 2015"
        End Try
    End Function

    Private Function ConvertDateToHex(ByVal dtpDateTime As DateTime) As String
        Dim strHexTime As String = vbNullString

        Try
            'Seconds, Minutes, Hours, Day, Date, Month and Year
            strHexTime = "#" & Hex(dtpDateTime.Day).ToString.PadLeft(2, "0")
            strHexTime = strHexTime & "#" & Hex(dtpDateTime.Month).ToString.PadLeft(2, "0")
            strHexTime = strHexTime & "#" & Hex(Val(dtpDateTime.Year.ToString.Substring(2, 2))).ToString.PadLeft(2, "0")
            strHexTime = strHexTime & "#" & Hex(GetDayofWeek(dtpDateTime.DayOfWeek)).ToString.PadLeft(2, "0")
            strHexTime = strHexTime & "#" & Hex(dtpDateTime.Hour).ToString.PadLeft(2, "0")
            strHexTime = strHexTime & "#" & Hex(dtpDateTime.Minute.ToString).PadLeft(2, "0")
            strHexTime = strHexTime & "#" & Hex(dtpDateTime.Second).ToString.PadLeft(2, "0")

        Catch ex As Exception
            WriteLog("ConvertDateToHex", Err.Description)
        End Try

        Return strHexTime
    End Function

    Private Function GetMonthName(ByVal MonthID As Integer) As String
        Try
            Select Case MonthID
                Case 1
                    Return "JAN"

                Case 2
                    Return "FEB"

                Case 3
                    Return "MAR"

                Case 4
                    Return "APR"

                Case 5
                    Return "MAY"

                Case 6
                    Return "JUN"

                Case 7
                    Return "JUL"

                Case 8
                    Return "AUG"

                Case 9
                    Return "SEP"

                Case 10
                    Return "OCT"

                Case 11
                    Return "NOV"

                Case 12
                    Return "DEC"

                Case Else
                    Return "JAN"
            End Select
        Catch ex As Exception
            WriteLog("GetMonthName", Err.Description)
        End Try

        Return "JAN"
    End Function

    Private Function GetDayofWeek(ByVal DayofWeek As DayOfWeek) As Integer
        Try

            Select Case DayofWeek
                Case System.DayOfWeek.Sunday
                    Return 1

                Case System.DayOfWeek.Monday
                    Return 2

                Case System.DayOfWeek.Tuesday
                    Return 3

                Case System.DayOfWeek.Wednesday
                    Return 4

                Case System.DayOfWeek.Thursday
                    Return 5

                Case System.DayOfWeek.Friday
                    Return 6

                Case System.DayOfWeek.Saturday
                    Return 7
            End Select

        Catch ex As Exception
            WriteLog("GetDayofWeek", Err.Description)
        End Try
    End Function

    Private Function FormatValue(ByVal MyValue As Double) As Double
        Try
            'Select Case MyValue
            '    Case Is < 10
            '        MyValue = Math.Round(MyValue, 4)

            '    Case Is < 100
            '        MyValue = Math.Round(MyValue, 2)

            '    Case Is < 1000
            '        MyValue = Math.Round(MyValue, 1)

            '    Case Is < 10000
            '        MyValue = Math.Round(MyValue, 0)

            '    Case Is < 100000
            '        MyValue = Math.Round(MyValue, 0)
            'End Select

            Return MyValue
        Catch ex As Exception
            WriteLog("FormatValue", Err.Description)
        End Try
    End Function

    Private Function FormatValue20(ByVal MyValue As Double) As Double
        Try
            Select Case MyValue
                Case Is < 100
                    MyValue = Math.Round(MyValue, 2)

                Case Is < 1000
                    MyValue = Math.Round(MyValue, 0)

            End Select

            Return MyValue
        Catch ex As Exception
            WriteLog("FormatValue", Err.Description)
        End Try
    End Function

    Private Sub WriteLog(ByVal strErrorSource As String, ByVal strErrorMsg As String)
        IO.File.AppendAllText("XWRM10A_ERROR.Log", strErrorSource & " : " & strErrorMsg)
    End Sub

#End Region

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Try
            objComPort.Close()
            objComPort.Dispose()

            GC.SuppressFinalize(Me)
        Catch ex As UnauthorizedAccessException

        End Try
    End Sub
End Class
