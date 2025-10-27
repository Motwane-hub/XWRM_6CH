'DVM 06 NOV 2014  New Class Created for XWRM10AData Instrument.
Imports System.ComponentModel
Imports System.Reflection
Imports Microsoft.VisualBasic.Conversion
Imports System.IO

Imports AccessXWRM25.clsAccessXWRM10A



Public Class clsAccessXWRM10A : Implements IDisposable
    Private strCOM As String = vbNullString
    Private WithEvents objComPort As IO.Ports.SerialPort

    Private blnRegistered As Boolean = False

    Public _ResponseResult As New Response

    Public Event ResponseRecieved(ByVal objSettingResponse As Response)
    Public Event ReportDataRecieved(ByVal dtReportData As DataTable)
    Public Event SendResponseFrame(ByVal strString As String)
    Public Event SendCommandFrame(ByVal strString As String)

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

    Public _comport As String = vbNullString

    Public Property Comport() As String
        Get
            Return _comport
        End Get
        Set(ByVal value As String)
            _comport = value
        End Set
    End Property

    Public _ResponseReceived As Boolean = False

    Public Property ResponseReceived() As Boolean
        Get
            Return _ResponseReceived
        End Get
        Set(ByVal value As Boolean)
            _ResponseReceived = value
        End Set
    End Property

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
        TEST_OFF = &H4
        DATE_SYNC = &H4
        GET_SETTING = &H5
        GET_STATUS = &H6
        GET_REPORT_COUNT = &H7
        GET_ALL_REPORT = &H8
        DEMAGNETIZE = &H9
        HOME = &H9
        SAVE = &H11
        PRINT_COMMAND = &H12
        CLEAR = &H13
        READ_DATE_TIME = &H14
        RECALL = &H15
        GET_DEVICE_TEST_HISTORY = &H16
        HR_ELAPSEDTIME = &H17
        NORMAL_CORRECTED_RESULT = &H18
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
        'HUNDREAD = 1
        'TEN = 2
        'ONE = 3
        'DIVAND1 = 4
        'AUTORANGE = 5

        TWO_KΩ = 1
        TWO_HUNDREAD_Ω = 2
        TWENTY_Ω = 3
        TWO_Ω = 4
        TWO_HUNDREAD_mΩ = 5
        TWENTY_mΩ = 6
        TWO_mΩ = 7
        TWO_HUNDREAD_uΩ = 8
        FOUR_Ω = 9
        FOUR_HUNDREAD_mΩ = 10
        FOURTY_mΩ = 11
        FOUR_mΩ = 12
        TWENTY_kΩ = 13
        AUTORANGE = 0


    End Enum

    Enum enmWindingSelection
        ALL_WINDING = 1
        DUAL_WINDING = 2
        SINGLE_WINDING = 3
        HV_WINDING = 4
        LV_WINDING = 5

    End Enum

    Enum enmWinding
        UN = 1
        VN = 2
        WN = 3
        NA = 0
    End Enum

    Enum enmTestMode
        NORMAL = 1
        HR = 2
        OLTC = 3
        DEMAGNETIZING = 4
        DYNAMIC_RESISTANCE = 5

    End Enum

    Enum enmWindingElement
        NO_ELEMENT = 0
        COPPER = 1
        ALUMINIUM = 2
    End Enum

    'Enum enmHVSelection
    '    SELECTED = 1
    '    Not_SELECTED = 0

    'End Enum

    'Enum enmLVSelection
    '    SELECTED = 1
    '    Not_SELECTED = 0

    'End Enum

    Enum enmHVCurrent
        ONE_mA = 1
        TEN_mA = 2
        HUNDREAD_mA = 3
        ONE_A = 4
        FIVE_A = 5
        TEN_A = 6
        TWENTYFINE_A = 7
        NA = 0  'NOT ANSWER

    End Enum

    Enum enmLVCurrent
        ONE_mA = 1
        TEN_mA = 2
        HUNDREAD_mA = 3
        ONE_A = 4
        FIVE_A = 5
        TEN_A = 6
        TWENTYFINE_A = 7
        NA = 0  'NOT ANSWER

    End Enum

    Enum enmVectorGroup
        YNYN = 1
        YND = 2
        DYN = 3
        DD = 4

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

    Public Sub New()
        MyBase.New()

        blnRegistered = True

        'Dim frmRegistration As New Registration()
        'blnRegistered = frmRegistration.IsRegistered()

        objComPort = New Ports.SerialPort
    End Sub

    Public Function Connect(ByVal ComPort As String) As Boolean
        Try
            'If Not CheckRegistration() Then Return False
            _comport = ComPort

            If objComPort.IsOpen = False Then
                objComPort = Nothing
                objComPort = New IO.Ports.SerialPort
                objComPort.PortName = ComPort
                objComPort.BaudRate = 115200
                objComPort.DataBits = 8
                objComPort.Parity = Ports.Parity.None
                objComPort.StopBits = Ports.StopBits.One
                objComPort.ReadBufferSize = 5000000
                objComPort.WriteBufferSize = 500000
                objComPort.ReadTimeout = 0
                objComPort.DtrEnable = True
                objComPort.RtsEnable = True
                objComPort.Handshake = Ports.Handshake.None

                AddHandler objComPort.DataReceived, AddressOf objComPort_DataReceived

                objComPort.Open()
                objComPort.DiscardInBuffer()
            Else
                Throw New ApplicationException("ER001")
            End If

            Return objComPort.IsOpen
        Catch ex As Exception
            MsgBox(Err.Description)
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
            ' Throw New ApplicationException("ER004")
        End Try
    End Sub


    Public Sub SetParameters(ByVal objTestmode As enmTestMode, ByVal vectorGroup As String, ByVal vectorGroupYNYN As Integer, ByVal windingSelection As Double, ByVal winding As Double, ByVal intHVCurrent As Double, ByVal intLVCurrent As Double, ByVal intHrTime As Integer, ByVal intTotalTap As Integer, ByVal intCurrentTap As Integer, ByVal objHVRange As enmRange, ByVal objLVRange As enmRange, ByVal testObject As String, ByVal DUTSpecification As String, ByVal DUTSerialNumber As String, ByVal DUTMake As String, ByVal DUTTestLocation As String, ByVal year As String, ByVal Temp1Loc As String, ByVal Temp2Loc As String)
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.SETTING).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            'Test Mode
            Select Case objTestmode
                Case enmTestMode.NORMAL
                    strCommand = strCommand & "#01"

                Case enmTestMode.HR
                    strCommand = strCommand & "#02"

                Case enmTestMode.OLTC
                    strCommand = strCommand & "#03"

                Case enmTestMode.DEMAGNETIZING
                    strCommand = strCommand & "#04"

                Case enmTestMode.DYNAMIC_RESISTANCE
                    strCommand = strCommand & "#05"
            End Select


            'Vector Group
            Select Case vectorGroup
                Case enmVectorGroup.YNYN
                    strCommand = strCommand & "#01"

                    If vectorGroupYNYN = 1 Then
                        strCommand = strCommand & "#01"
                    ElseIf vectorGroupYNYN = 2 Then
                        strCommand = strCommand & "#02"
                    End If

                Case enmVectorGroup.YND
                    strCommand = strCommand & "#02"
                    strCommand = strCommand & "#00"

                Case enmVectorGroup.DYN
                    strCommand = strCommand & "#03"
                    strCommand = strCommand & "#00"

                Case enmVectorGroup.DD
                    strCommand = strCommand & "#04"
                    strCommand = strCommand & "#00"

            End Select


            Select Case windingSelection
                Case enmWindingSelection.ALL_WINDING
                    strCommand = strCommand & "#01"
                    strCommand = strCommand & "#00"

                Case enmWindingSelection.DUAL_WINDING
                    strCommand = strCommand & "#02"
                    strCommand = strCommand & "#00"

                Case enmWindingSelection.SINGLE_WINDING
                    strCommand = strCommand & "#03"

                    Select Case winding
                        Case enmWinding.UN
                            strCommand = strCommand & "#01"

                        Case enmWinding.VN
                            strCommand = strCommand & "#02"

                        Case enmWinding.WN
                            strCommand = strCommand & "#03"

                    End Select

                Case enmWindingSelection.HV_WINDING

                    strCommand = strCommand & "#04"
                    strCommand = strCommand & "#00"

                Case enmWindingSelection.LV_WINDING

                    strCommand = strCommand & "#05"
                    strCommand = strCommand & "#00"

            End Select


            Select Case intHVCurrent
                Case enmHVCurrent.ONE_mA
                    strCommand = strCommand & "#01"
                Case enmHVCurrent.TEN_mA
                    strCommand = strCommand & "#02"
                Case enmHVCurrent.HUNDREAD_mA
                    strCommand = strCommand & "#03"
                Case enmHVCurrent.ONE_A
                    strCommand = strCommand & "#04"
                Case enmHVCurrent.FIVE_A
                    strCommand = strCommand & "#05"
                Case enmHVCurrent.TEN_A
                    strCommand = strCommand & "#06"
                Case enmHVCurrent.TWENTYFINE_A
                    strCommand = strCommand & "#07"
                Case enmHVCurrent.NA
                    strCommand = strCommand & "#00"
            End Select

            Select Case intLVCurrent
                Case enmLVCurrent.ONE_mA
                    strCommand = strCommand & "#01"
                Case enmLVCurrent.TEN_mA
                    strCommand = strCommand & "#02"
                Case enmLVCurrent.HUNDREAD_mA
                    strCommand = strCommand & "#03"
                Case enmLVCurrent.ONE_A
                    strCommand = strCommand & "#04"
                Case enmLVCurrent.FIVE_A
                    strCommand = strCommand & "#05"
                Case enmLVCurrent.TEN_A
                    strCommand = strCommand & "#06"
                Case enmLVCurrent.TWENTYFINE_A
                    strCommand = strCommand & "#07"
                Case enmLVCurrent.TWENTYFINE_A
                    strCommand = strCommand & "#00"
            End Select


            strCommand = strCommand & "#" & Hex(intHrTime).PadLeft(2, "0")      'HR Time 1 Byte

            strCommand = strCommand & "#" & Hex(intTotalTap).PadLeft(2, "0")

            strCommand = strCommand & "#" & Hex(intCurrentTap).PadLeft(2, "0")

            'MANUAL RES RANGE (IF HR TETS SELECTED)
            Select Case objHVRange

                Case enmRange.TWO_KΩ
                    strCommand = strCommand & "#01"

                Case enmRange.TWO_HUNDREAD_Ω
                    strCommand = strCommand & "#02"

                Case enmRange.TWENTY_Ω
                    strCommand = strCommand & "#03"

                Case enmRange.TWO_Ω
                    strCommand = strCommand & "#04"

                Case enmRange.TWO_HUNDREAD_mΩ
                    strCommand = strCommand & "#05"

                Case enmRange.TWENTY_mΩ
                    strCommand = strCommand & "#06"

                Case enmRange.TWO_mΩ
                    strCommand = strCommand & "#07"

                Case enmRange.TWO_HUNDREAD_uΩ
                    strCommand = strCommand & "#08"

                Case enmRange.FOUR_Ω
                    strCommand = strCommand & "#09"

                Case enmRange.FOUR_HUNDREAD_mΩ
                    strCommand = strCommand & "#10"

                Case enmRange.FOURTY_mΩ
                    strCommand = strCommand & "#11"

                Case enmRange.FOUR_mΩ
                    strCommand = strCommand & "#12"

                Case enmRange.TWENTY_kΩ
                    strCommand = strCommand & "#13"

                Case enmRange.AUTORANGE
                    strCommand = strCommand & "#00"


            End Select

            Select Case objLVRange

                Case enmRange.TWO_KΩ
                    strCommand = strCommand & "#01"

                Case enmRange.TWO_HUNDREAD_Ω
                    strCommand = strCommand & "#02"

                Case enmRange.TWENTY_Ω
                    strCommand = strCommand & "#03"

                Case enmRange.TWO_Ω
                    strCommand = strCommand & "#04"

                Case enmRange.TWO_HUNDREAD_mΩ
                    strCommand = strCommand & "#05"

                Case enmRange.TWENTY_mΩ
                    strCommand = strCommand & "#06"

                Case enmRange.TWO_mΩ
                    strCommand = strCommand & "#07"

                Case enmRange.TWO_HUNDREAD_uΩ
                    strCommand = strCommand & "#08"

                Case enmRange.FOUR_Ω
                    strCommand = strCommand & "#09"

                Case enmRange.FOUR_HUNDREAD_mΩ
                    strCommand = strCommand & "#10"

                Case enmRange.FOURTY_mΩ
                    strCommand = strCommand & "#11"

                Case enmRange.FOUR_mΩ
                    strCommand = strCommand & "#12"

                Case enmRange.TWENTY_kΩ
                    strCommand = strCommand & "#13"

                Case enmRange.AUTORANGE
                    strCommand = strCommand & "#00"

            End Select

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim remainingBytes As Integer

            ' DUT testObject
            Dim strDUTTestObject As String = vbNullString

            If testObject.Length > 16 Then
                testObject = testObject.Substring(0, 16)
            End If

            For Each strChar As Char In testObject
                strDUTTestObject &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - testObject.Length
            For i As Integer = 1 To remainingBytes
                strDUTTestObject &= "#00"
            Next

            If strDUTTestObject = vbNullString Then
                strDUTTestObject = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTTestObject
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' DUT Specification
            Dim strDUTSpecification As String = vbNullString

            If DUTSpecification.Length > 16 Then
                DUTSpecification = DUTSpecification.Substring(0, 16)
            End If

            For Each strChar As Char In DUTSpecification
                strDUTSpecification &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - DUTSpecification.Length
            For i As Integer = 1 To remainingBytes
                strDUTSpecification &= "#00"
            Next

            If strDUTSpecification = vbNullString Then
                strDUTSpecification = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTSpecification

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' DUT Serial Number
            Dim strDUTSerialNumber As String = vbNullString

            If DUTSerialNumber.Length > 16 Then
                DUTSerialNumber = DUTSerialNumber.Substring(0, 16)
            End If

            For Each strChar As Char In DUTSerialNumber
                strDUTSerialNumber &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - DUTSerialNumber.Length
            For i As Integer = 1 To remainingBytes
                strDUTSerialNumber &= "#00"
            Next

            If strDUTSerialNumber = vbNullString Then
                strDUTSerialNumber = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTSerialNumber

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' DUT Make
            Dim strDUTMake As String = vbNullString

            If DUTMake.Length > 16 Then
                DUTMake = DUTMake.Substring(0, 16)
            End If

            For Each strChar As Char In DUTMake
                strDUTMake &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - DUTMake.Length
            For i As Integer = 1 To remainingBytes
                strDUTMake &= "#00"
            Next

            If strDUTMake = vbNullString Then
                strDUTMake = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTMake

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' DUT Test Location
            Dim strDUTTestLocation As String = vbNullString

            If DUTTestLocation.Length > 16 Then
                DUTTestLocation = DUTTestLocation.Substring(0, 16)
            End If

            For Each strChar As Char In DUTTestLocation
                strDUTTestLocation &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - DUTTestLocation.Length
            For i As Integer = 1 To remainingBytes
                strDUTTestLocation &= "#00"
            Next

            If strDUTTestLocation = vbNullString Then
                strDUTTestLocation = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTTestLocation

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' DUT YEAR
            Dim strYear As String = vbNullString

            If year.Length > 4 Then
                year = year.Substring(0, 4)
            End If

            For Each strChar As Char In year
                strYear &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 4 - year.Length
            For i As Integer = 1 To remainingBytes
                strYear &= "#00"
            Next

            If strYear = vbNullString Then
                strYear = String.Join("", Enumerable.Repeat("#00", 4))
            End If

            strCommand &= strYear

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''Temp1Loc
            Dim strDUTTemp1Loc As String = vbNullString

            If Temp1Loc.Length > 16 Then
                Temp1Loc = Temp1Loc.Substring(0, 16)
            End If

            For Each strChar As Char In Temp1Loc
                strDUTTemp1Loc &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - Temp1Loc.Length
            For i As Integer = 1 To remainingBytes
                strDUTTemp1Loc &= "#00"
            Next

            If strDUTTemp1Loc = vbNullString Then
                strDUTTemp1Loc = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTTemp1Loc

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''Temp2Loc
            Dim strDUTTemp2Loc As String = vbNullString

            If Temp2Loc.Length > 16 Then
                Temp2Loc = Temp2Loc.Substring(0, 16)
            End If

            For Each strChar As Char In Temp2Loc
                strDUTTemp2Loc &= "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
            Next

            remainingBytes = 16 - Temp2Loc.Length
            For i As Integer = 1 To remainingBytes
                strDUTTemp2Loc &= "#00"
            Next

            If strDUTTemp2Loc = vbNullString Then
                strDUTTemp2Loc = String.Join("", Enumerable.Repeat("#00", 16))
            End If

            strCommand &= strDUTTemp2Loc

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            strCommand = strCommand & "#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(133) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("SetParameters " & Err.Description)
        End Try
    End Sub


    'Public Sub SetParameters(ByVal current As Double, ByVal currentUnit As enmCurrentsUnit, ByVal channel As enmChannels, ByVal ch1Range As enmRange, ByVal ch2Range As enmRange, ByVal ch3Range As enmRange, ByVal testMode As enmTestMode, ByVal blnSaveTest As Boolean, ByVal HrTestTime As Integer, ByVal DUTSerialNumber As String, ByVal DUTType As String, ByVal DUTTestLocation As String, ByVal HighestTap As Integer, ByVal WindingElement As enmWindingElement, ByVal ProbeIsConnected As Boolean, ByVal CorrectedTempratue As Boolean, ByVal ambientTemprature As Double)
    '    Try
    '        If objComPort.IsOpen = False Then Exit Sub

    '        Dim strCommand As String = vbNullString

    '        strCommand = "3A"                                                   'SOF                '1 Byte
    '        strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
    '        strCommand = strCommand & "#" & Hex(enmCommands.SETTING).PadLeft(2, "0")  'Op Code '1 Byte
    '        strCommand = strCommand & "#C0"                                     'Command '1 Byte 

    '        'Current Value
    '        Dim strCurrent As String = ConvertSingleToHex(current)
    '        strCurrent = strCurrent.Insert(0, "#")
    '        strCurrent = strCurrent.Insert(3, "#")
    '        strCurrent = strCurrent.Insert(6, "#")
    '        strCurrent = strCurrent.Insert(9, "#")

    '        'Current Unit
    '        strCommand = strCommand & strCurrent

    '        If currentUnit = enmCurrentsUnit.Ampere Then
    '            strCommand = strCommand & "#01"
    '        End If

    '        If currentUnit = enmCurrentsUnit.mAmpere Then
    '            strCommand = strCommand & "#00"
    '        End If

    '        Select Case channel
    '            Case enmChannels.ALL_CH_OFF
    '                strCommand = strCommand & "#00"

    '            Case enmChannels.CH1_ON
    '                strCommand = strCommand & "#01"

    '            Case enmChannels.CH2_ON
    '                strCommand = strCommand & "#02"

    '            Case enmChannels.CH3_ON
    '                strCommand = strCommand & "#04"

    '            Case enmChannels.CH1_CH2_ON
    '                strCommand = strCommand & "#03"

    '            Case enmChannels.CH2_CH3_ON
    '                strCommand = strCommand & "#06"

    '            Case enmChannels.CH1_CH3_ON
    '                strCommand = strCommand & "#05"

    '            Case enmChannels.ALL_CH_ON
    '                strCommand = strCommand & "#07"
    '        End Select

    '        Select Case ch1Range
    '            Case enmRange.HUNDREAD
    '                strCommand = strCommand & "#00"

    '            Case enmRange.TEN
    '                strCommand = strCommand & "#02"

    '            Case enmRange.ONE
    '                strCommand = strCommand & "#01"

    '            Case enmRange.DIVAND1
    '                strCommand = strCommand & "#04"

    '            Case enmRange.AUTORANGE
    '                strCommand = strCommand & "#08"
    '        End Select

    '        Dim strBinary1 As String = ""

    '        Select Case ch2Range
    '            Case enmRange.HUNDREAD
    '                strBinary1 = "0000"

    '            Case enmRange.TEN
    '                strBinary1 = "0010"

    '            Case enmRange.ONE
    '                strBinary1 = "0001"

    '            Case enmRange.DIVAND1
    '                strBinary1 = "0100"

    '            Case enmRange.AUTORANGE
    '                strBinary1 = "1000"
    '        End Select

    '        Dim strBinary2 As String = ""

    '        Select Case ch3Range
    '            Case enmRange.HUNDREAD
    '                strBinary2 = "0000"

    '            Case enmRange.TEN
    '                strBinary2 = "0010"

    '            Case enmRange.ONE
    '                strBinary2 = "0001"

    '            Case enmRange.DIVAND1
    '                strBinary2 = "0100"

    '            Case enmRange.AUTORANGE
    '                strBinary2 = "1000"
    '        End Select

    '        Dim strBinary As String = strBinary1 & strBinary2
    '        strBinary = BinaryToHex(strBinary).PadLeft(2, "0")

    '        strCommand = strCommand & "#" & strBinary

    '        Select Case testMode
    '            Case enmTestMode.NORMAL
    '                strBinary1 = "0100"

    '            Case enmTestMode.OLTC
    '                strBinary1 = "0001"

    '            Case enmTestMode.HR
    '                strBinary1 = "0010"

    '            Case enmTestMode.DYNAMIC_RESISTANCE
    '                strBinary1 = "1000"
    '        End Select

    '        If blnSaveTest Then
    '            strBinary2 = "0001"
    '        Else
    '            strBinary2 = "0000"
    '        End If

    '        strBinary = strBinary1 & strBinary2
    '        strBinary = BinaryToHex(strBinary).PadLeft(2, "0")

    '        strCommand = strCommand & "#" & strBinary

    '        Dim strHRTestTime As String = "#" & Hex(HrTestTime).PadLeft(2, "0")

    '        Select Case testMode
    '            Case enmTestMode.NORMAL
    '                strHRTestTime = "#" & Hex(HrTestTime).PadLeft(2, "0")

    '            Case enmTestMode.OLTC
    '                Dim strBinary5 As String = ToBinary(CDec("&h" & Hex(HrTestTime).PadLeft(2, "0")))
    '                Dim strNewByte As String = Val(HighestTap).ToString() & strBinary5.Substring(1, 7)

    '                strHRTestTime = "#" & BinaryToHex(strNewByte).PadLeft(2, "0")

    '            Case enmTestMode.HR
    '                strHRTestTime = "#" & Hex(HrTestTime).PadLeft(2, "0")
    '        End Select

    '        strCommand = strCommand & strHRTestTime

    '        ''DUT Serial Number
    '        Dim strDUTSerialNumber As String = vbNullString

    '        If DUTSerialNumber.Length > 15 Then
    '            DUTSerialNumber = DUTSerialNumber.Substring(0, 15)
    '        End If

    '        For Each strChar As Char In DUTSerialNumber
    '            strDUTSerialNumber = strDUTSerialNumber & "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
    '        Next

    '        If strDUTSerialNumber = vbNullString Then
    '            strDUTSerialNumber = "#00"
    '        End If

    '        Do While strDUTSerialNumber.Length < 45
    '            strDUTSerialNumber = strDUTSerialNumber & "#00"
    '        Loop

    '        strCommand = strCommand & strDUTSerialNumber

    '        ''DUT Type
    '        Dim strDUTType As String = vbNullString

    '        If DUTType.Length > 15 Then
    '            DUTType = DUTType.Substring(0, 15)
    '        End If

    '        For Each strChar As Char In DUTType
    '            strDUTType = strDUTType & "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
    '        Next

    '        If strDUTType = vbNullString Then
    '            strDUTType = "#00"
    '        End If

    '        Do While strDUTType.Length < 45
    '            strDUTType = strDUTType & "#00"
    '        Loop

    '        strCommand = strCommand & strDUTType

    '        ''DUT Test Location
    '        Dim strDUTTestLocation As String = vbNullString

    '        If DUTTestLocation.Length > 15 Then
    '            DUTTestLocation = DUTTestLocation.Substring(0, 15)
    '        End If

    '        For Each strChar As Char In DUTTestLocation
    '            strDUTTestLocation = strDUTTestLocation & "#" & ConvertCharToHex(strChar).PadLeft(2, "0")
    '        Next

    '        If strDUTTestLocation = vbNullString Then
    '            strDUTTestLocation = "#00"
    '        End If

    '        Do While strDUTTestLocation.Length < 45
    '            strDUTTestLocation = strDUTTestLocation & "#00"
    '        Loop

    '        strCommand = strCommand & strDUTTestLocation

    '        'Start here other location
    '        'ByVal WindingElement As Integer, 
    '        'ByVal ProbeIsConnected As Boolean,
    '        'ByVal CorrectedTempratue As Boolean, 
    '        Dim strBinarytemp As String = ""

    '        Select Case WindingElement
    '            Case enmWindingElement.NO_ELEMENT
    '                strBinarytemp = "0000"
    '            Case enmWindingElement.COPPER
    '                strBinarytemp = "0001"
    '            Case enmWindingElement.ALUMINIUM
    '                strBinarytemp = "0010"
    '        End Select

    '        Dim strbit4 As String = "0"

    '        If ProbeIsConnected Then
    '            strbit4 = "0"
    '        Else
    '            strbit4 = "1"
    '        End If

    '        Dim strbit7 As String = "0"

    '        If CorrectedTempratue Then
    '            strbit7 = "1"
    '        Else
    '            strbit7 = "0"
    '        End If

    '        strBinarytemp = strbit7 & "00" & strbit4 & strBinarytemp

    '        strBinary = BinaryToHex(strBinarytemp).PadLeft(2, "0")
    '        strCommand = strCommand & "#" & strBinary

    '        'Current Value
    '        Dim strAmbientTemprature As String = ConvertSingleToHex(ambientTemprature)
    '        strAmbientTemprature = strAmbientTemprature.Insert(0, "#")
    '        strAmbientTemprature = strAmbientTemprature.Insert(3, "#")
    '        strAmbientTemprature = strAmbientTemprature.Insert(6, "#")
    '        strAmbientTemprature = strAmbientTemprature.Insert(9, "#")

    '        strCommand = strCommand & strAmbientTemprature
    '        strCommand = strCommand & "#01"

    '        strCommand = strCommand & "#00#00#00"

    '        strCommand = strCommand & "#00"                                     'CRC
    '        strCommand = strCommand & "#25"                                     'EOF

    '        Dim strCommandFrame() As String
    '        strCommandFrame = strCommand.Split("#".ToCharArray)

    '        strCommandFrame(68) = CalculateCRC(strCommandFrame)
    '        RaiseEvent SendCommandFrame(strCommand)
    '        SendFrame(strCommandFrame)
    '    Catch ex As Exception
    '        MsgBox("SetParameters " & Err.Description)
    '    End Try
    'End Sub

    Public Sub TesON()
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.TEST_ON).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(48) = CalculateCRC(strCommandFrame)
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

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(48) = CalculateCRC(strCommandFrame)
            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("Test OFF " & Err.Description)
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
            LastCommand = CDec("&h" & strCommandFrame(2))
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

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(48) = CalculateCRC(strCommandFrame)
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

    Public Sub SetHRElapsedTime(ByVal dblHRTime As Double)
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.HR_ELAPSEDTIME).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            Dim strHex As String = Hex(dblHRTime).PadLeft(4, "0")
            strHex = "#" & strHex.Insert(2, "#")

            strCommand = strCommand & strHex

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)

            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("SetHRElapsedTime " & Err.Description)
        End Try
    End Sub

    Public Sub SetNormalHR(ByVal NormalCorrected As Double)
        Try
            If objComPort.IsOpen = False Then Exit Sub

            Dim strCommand As String = vbNullString

            strCommand = "3A"                                                   'SOF                '1 Byte
            strCommand = strCommand & "#10"                                     'Poduct ID          '1 Byte
            strCommand = strCommand & "#" & Hex(enmCommands.NORMAL_CORRECTED_RESULT).PadLeft(2, "0")  'Op Code '1 Byte
            strCommand = strCommand & "#C0"                                     'Command '1 Byte 

            Dim strHex As String = Hex(NormalCorrected).PadLeft(4, "0")
            strHex = "#" & strHex.Insert(2, "#")

            strCommand = strCommand & strHex

            strCommand = strCommand & "#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00#00"

            strCommand = strCommand & "#00"                                     'CRC
            strCommand = strCommand & "#25"                                     'EOF

            Dim strCommandFrame() As String
            strCommandFrame = strCommand.Split("#".ToCharArray)

            strCommandFrame(47) = CalculateCRC(strCommandFrame)

            RaiseEvent SendCommandFrame(strCommand)
            SendFrame(strCommandFrame)
        Catch ex As Exception
            MsgBox("SetNormalHR " & Err.Description)
        End Try
    End Sub

    Public Function GetCorrectedReading(ByVal Reading As Double, ByVal CorrectedTemprature As Double, ByVal MeasuredTemprature As Double, ByVal Ref_Element As enmWindingElement) As Double
        Try
            Dim Tk_matrl_temp As Double = 234.5
            Dim temp_cnstnt As Double = 1.0

            Select Case Ref_Element
                Case enmWindingElement.COPPER
                    Tk_matrl_temp = 234.5

                Case enmWindingElement.ALUMINIUM
                    Tk_matrl_temp = 225.0

                Case enmWindingElement.NO_ELEMENT
                    Tk_matrl_temp = 0
            End Select

            temp_cnstnt = (CorrectedTemprature + Tk_matrl_temp) / (MeasuredTemprature + Tk_matrl_temp)
            Return Reading * (temp_cnstnt)
        Catch ex As Exception
            Return 0
        End Try
    End Function

#End Region

#Region "Private Function and Subroutines"

#Region "Send Frame"

    Dim strResponseString As New Text.StringBuilder
    Dim tempStringBuilder As New Text.StringBuilder

    Public Property GetReportData() As String
        Get
            Return strResponseString.ToString
        End Get
        Set(ByVal value As String)
            If value = "" Then
                strResponseString = New Text.StringBuilder
                tempStringBuilder = New Text.StringBuilder
            Else
                strResponseString.Append(value)
                tempStringBuilder = New Text.StringBuilder
            End If

        End Set
    End Property

    Private Sub objComPort_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles objComPort.DataReceived
        Try
            System.Threading.Thread.Sleep(200)

            If objComPort.BytesToRead = 0 Then
                Exit Sub
            End If

            _IsCommunicated = True
            _ResponseReceived = True

            Select Case LastCommand
                Case enmCommands.TEST_OFF

                Case enmCommands.GET_ALL_REPORT
                    Dim objCollection As New Collection

                    While objComPort.BytesToRead > 0
                        objCollection.Add(objComPort.ReadByte())

                        If objComPort.BytesToRead = 0 Then
                            System.Threading.Thread.Sleep(100)
                        End If
                    End While

                    ProcessReportData(objCollection)

                Case enmCommands.GET_SETTING

                    Dim btDataNew(69) As Byte

                    For iBytesCount As Integer = 0 To 69
                        btDataNew(iBytesCount) = objComPort.ReadByte()
                    Next

                    ProcessResponseFrame(btDataNew)

                Case enmCommands.GET_STATUS, enmCommands.TEST_ON
                    'Dim btDataNew(91) As Byte

                    'For iBytesCount As Integer = 0 To 91
                    '    btDataNew(iBytesCount) = objComPort.ReadByte()
                    'Next

                    Dim btDataNew(49) As Byte

                    For iBytesCount As Integer = 0 To 49
                        btDataNew(iBytesCount) = objComPort.ReadByte()
                    Next

                    ProcessResponseFrame(btDataNew)

                Case enmCommands.DEMAGNETIZE


                Case Else

                    Dim btDataNew(49) As Byte

                    For iBytesCount As Integer = 0 To 49
                        btDataNew(iBytesCount) = objComPort.ReadByte()
                    Next

                    ProcessResponseFrame(btDataNew)
            End Select

            'objComPort.DiscardOutBuffer()
        Catch ex As Exception
            'MsgBox(Err.Description)
            '  WriteLog("Error : objComPort_DataReceived()", Err.Description)
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
                Case 50 'Normal

                    Select Case btData(3)

                        'Winding Position

                        Case 209 'Normal Response

                            Select Case btData(28)
                                Case 1
                                    _ResponseResult.Winding = enmWinding.UN
                                Case 2
                                    _ResponseResult.Winding = enmWinding.VN
                                Case 3
                                    _ResponseResult.Winding = enmWinding.WN
                            End Select

                            'HV Current Reading
                            strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0") & Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")
                            _ResponseResult.HVTestCurrent = ConvertHexToSingle(strData)

                            'HV Current Unit

                            If btData(8) = 1 Or btData(8) = 2 Or btData(8) = 3 Then
                                _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.mAmpere
                            ElseIf btData(8) = 4 Or btData(8) = 5 Or btData(8) = 6 Or btData(8) = 7 Then
                                _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.Ampere
                            End If

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

                            'LV Current Reading
                            strData = Hex(btData(9)).PadLeft(2, "0") & Hex(btData(10)).PadLeft(2, "0") & Hex(btData(11)).PadLeft(2, "0") & Hex(btData(12)).PadLeft(2, "0")
                            _ResponseResult.LVTestCurrent = ConvertHexToSingle(strData)

                            'LV Current Unit

                            If btData(13) = 1 Or btData(13) = 2 Or btData(13) = 3 Then
                                _ResponseResult.LVTestCurrentUnit = enmCurrentsUnit.mAmpere
                            ElseIf btData(13) = 4 Or btData(13) = 5 Or btData(13) = 6 Or btData(13) = 7 Then
                                _ResponseResult.LVTestCurrentUnit = enmCurrentsUnit.Ampere
                            End If

                            'HV Winding Reading
                            strData = Hex(btData(14)).PadLeft(2, "0") & Hex(btData(15)).PadLeft(2, "0") & Hex(btData(16)).PadLeft(2, "0") & Hex(btData(17)).PadLeft(2, "0")
                            _ResponseResult.HVReading = ConvertHexToSingle(strData)

                            Select Case _ResponseResult.HVReading
                                Case Is < 1
                                    _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 5)
                                Case Is < 10
                                    _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 4)
                                Case Is < 100
                                    _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 3)
                                Case Is < 1000
                                    _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 2)
                                Case Is < 10000
                                    _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 1)
                            End Select

                            'HV Winding Unit

                            Select Case btData(18)
                                Case 1
                                    _ResponseResult.HVUnit = enmResistanceUnit.uOhm

                                Case 2
                                    _ResponseResult.HVUnit = enmResistanceUnit.mOhm

                                Case 3
                                    _ResponseResult.HVUnit = enmResistanceUnit.Ohm

                                Case 4
                                    _ResponseResult.HVUnit = enmResistanceUnit.KOhm
                            End Select

                            'LV Winding Reading
                            strData = Hex(btData(19)).PadLeft(2, "0") & Hex(btData(20)).PadLeft(2, "0") & Hex(btData(21)).PadLeft(2, "0") & Hex(btData(22)).PadLeft(2, "0")
                            _ResponseResult.LVReading = ConvertHexToSingle(strData)

                            Select Case _ResponseResult.LVReading
                                Case Is < 1
                                    _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 5)
                                Case Is < 10
                                    _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 4)
                                Case Is < 100
                                    _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 3)
                                Case Is < 1000
                                    _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 2)
                                Case Is < 10000
                                    _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 1)
                            End Select

                            'LV Winding Unit

                            Select Case btData(23)
                                Case 1
                                    _ResponseResult.LVUnit = enmResistanceUnit.uOhm

                                Case 2
                                    _ResponseResult.LVUnit = enmResistanceUnit.mOhm

                                Case 3
                                    _ResponseResult.LVUnit = enmResistanceUnit.Ohm

                                Case 4
                                    _ResponseResult.LVUnit = enmResistanceUnit.KOhm
                            End Select

                            'T1 reading and unit
                            strData = Hex(btData(24)).PadLeft(2, "0") & Hex(btData(25)).PadLeft(2, "0") & Hex(btData(26)).PadLeft(2, "0") & Hex(btData(27)).PadLeft(2, "0")
                            _ResponseResult.T1Reading = Math.Round(ConvertHexToSingle(strData), 1)


                        Case 210 'OLTC Response


                            Select Case btData(4)  'Current
                                Case 1
                                    strData = Hex(btData(5)).PadLeft(2, "0") & Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0") & Hex(btData(8)).PadLeft(2, "0")
                                    _ResponseResult.CURRENT = ConvertHexToSingle(strData)
                                Case 2
                                    strData = Hex(btData(5)).PadLeft(2, "0") & Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0") & Hex(btData(8)).PadLeft(2, "0")
                                    _ResponseResult.CURRENT = ConvertHexToSingle(strData)

                            End Select

                            'CURRENT UNIT
                            If btData(9) = 1 Or btData(9) = 2 Or btData(9) = 3 Then
                                _ResponseResult.CURRENT_UNIT = enmCurrentsUnit.mAmpere
                            ElseIf btData(9) = 4 Or btData(9) = 5 Or btData(9) = 6 Or btData(9) = 7 Then
                                _ResponseResult.CURRENT_UNIT = enmCurrentsUnit.Ampere
                            End If

                            ' Reading
                            strData = Hex(btData(10)).PadLeft(2, "0") & Hex(btData(11)).PadLeft(2, "0") & Hex(btData(12)).PadLeft(2, "0") & Hex(btData(13)).PadLeft(2, "0")
                            _ResponseResult.READING = ConvertHexToSingle(strData)

                            Select Case _ResponseResult.READING
                                Case Is < 1
                                    _ResponseResult.READING = Math.Round(_ResponseResult.READING, 5)
                                Case Is < 10
                                    _ResponseResult.READING = Math.Round(_ResponseResult.READING, 4)
                                Case Is < 100
                                    _ResponseResult.READING = Math.Round(_ResponseResult.READING, 3)
                                Case Is < 1000
                                    _ResponseResult.READING = Math.Round(_ResponseResult.READING, 2)
                                Case Is < 10000
                                    _ResponseResult.READING = Math.Round(_ResponseResult.READING, 1)
                            End Select

                            'reading unit
                            Select Case btData(14)
                                Case 1
                                    _ResponseResult.READING_UNIT = enmResistanceUnit.uOhm

                                Case 2
                                    _ResponseResult.READING_UNIT = enmResistanceUnit.mOhm

                                Case 3
                                    _ResponseResult.READING_UNIT = enmResistanceUnit.Ohm

                                Case 4
                                    _ResponseResult.READING_UNIT = enmResistanceUnit.KOhm
                            End Select

                            'T1 reading and unit
                            strData = Hex(btData(15)).PadLeft(2, "0") & Hex(btData(16)).PadLeft(2, "0") & Hex(btData(17)).PadLeft(2, "0") & Hex(btData(18)).PadLeft(2, "0")
                            _ResponseResult.T1Reading = Math.Round(ConvertHexToSingle(strData), 1)

                            Select Case btData(19) 'winding Pos
                                Case 1
                                    _ResponseResult.Winding = enmWinding.UN
                                Case 2
                                    _ResponseResult.Winding = enmWinding.VN
                                Case 3
                                    _ResponseResult.Winding = enmWinding.WN
                            End Select

                            _ResponseResult.TapNumber = btData(20) 'tap Numbers

                        Case 211 'HR Response

                                    'Winding Position

                                    Select Case btData(32)
                                        Case 1
                                            _ResponseResult.Winding = enmWinding.UN
                                        Case 2
                                            _ResponseResult.Winding = enmWinding.VN
                                        Case 3
                                            _ResponseResult.Winding = enmWinding.WN
                                    End Select

                                    'HV Current Reading
                                    strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0") & Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")
                                    _ResponseResult.HVTestCurrent = ConvertHexToSingle(strData)

                                    'HV Current Unit

                                    If btData(8) = 1 Or btData(8) = 2 Or btData(8) = 3 Then
                                        _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.mAmpere
                                    ElseIf btData(8) = 4 Or btData(8) = 5 Or btData(8) = 6 Or btData(8) = 7 Then
                                        _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.Ampere
                                    End If

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

                                    'LV Current Reading
                                    strData = Hex(btData(9)).PadLeft(2, "0") & Hex(btData(10)).PadLeft(2, "0") & Hex(btData(11)).PadLeft(2, "0") & Hex(btData(12)).PadLeft(2, "0")
                                    _ResponseResult.LVTestCurrent = ConvertHexToSingle(strData)


                                    'LV Current Unit
                                    If btData(13) = 1 Or btData(13) = 2 Or btData(13) = 3 Then
                                        _ResponseResult.LVTestCurrentUnit = enmCurrentsUnit.mAmpere
                                    ElseIf btData(13) = 4 Or btData(13) = 5 Or btData(13) = 6 Or btData(13) = 7 Then
                                        _ResponseResult.LVTestCurrentUnit = enmCurrentsUnit.Ampere
                                    End If

                                    'HV Winding Reading
                                    strData = Hex(btData(14)).PadLeft(2, "0") & Hex(btData(15)).PadLeft(2, "0") & Hex(btData(16)).PadLeft(2, "0") & Hex(btData(17)).PadLeft(2, "0")
                                    _ResponseResult.HVReading = ConvertHexToSingle(strData)

                                    Select Case _ResponseResult.HVReading
                                        Case Is < 1
                                            _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 5)
                                        Case Is < 10
                                            _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 4)
                                        Case Is < 100
                                            _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 3)
                                        Case Is < 1000
                                            _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 2)
                                        Case Is < 10000
                                            _ResponseResult.HVReading = Math.Round(_ResponseResult.HVReading, 1)
                                    End Select

                                    'HV Winding Unit

                                    Select Case btData(18)
                                        Case 1
                                            _ResponseResult.HVUnit = enmResistanceUnit.uOhm

                                        Case 2
                                            _ResponseResult.HVUnit = enmResistanceUnit.mOhm

                                        Case 3
                                            _ResponseResult.HVUnit = enmResistanceUnit.Ohm

                                        Case 4
                                            _ResponseResult.HVUnit = enmResistanceUnit.KOhm
                                    End Select

                                    'LV Winding Reading
                                    strData = Hex(btData(19)).PadLeft(2, "0") & Hex(btData(20)).PadLeft(2, "0") & Hex(btData(21)).PadLeft(2, "0") & Hex(btData(22)).PadLeft(2, "0")
                                    _ResponseResult.LVReading = ConvertHexToSingle(strData)

                                    Select Case _ResponseResult.LVReading
                                        Case Is < 1
                                            _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 5)
                                        Case Is < 10
                                            _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 4)
                                        Case Is < 100
                                            _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 3)
                                        Case Is < 1000
                                            _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 2)
                                        Case Is < 10000
                                            _ResponseResult.LVReading = Math.Round(_ResponseResult.LVReading, 1)
                                    End Select

                                    'LV Winding Unit

                                    Select Case btData(23)
                                        Case 1
                                            _ResponseResult.LVUnit = enmResistanceUnit.uOhm

                                        Case 2
                                            _ResponseResult.LVUnit = enmResistanceUnit.mOhm

                                        Case 3
                                            _ResponseResult.LVUnit = enmResistanceUnit.Ohm

                                        Case 4
                                            _ResponseResult.LVUnit = enmResistanceUnit.KOhm
                                    End Select

                                    'T1 reading and unit
                                    strData = Hex(btData(24)).PadLeft(2, "0") & Hex(btData(25)).PadLeft(2, "0") & Hex(btData(26)).PadLeft(2, "0") & Hex(btData(27)).PadLeft(2, "0")
                                    _ResponseResult.T1Reading = Math.Round(ConvertHexToSingle(strData), 1)

                                    'T2 reading and unit
                                    strData = Hex(btData(28)).PadLeft(2, "0") & Hex(btData(29)).PadLeft(2, "0") & Hex(btData(30)).PadLeft(2, "0") & Hex(btData(31)).PadLeft(2, "0")
                                    _ResponseResult.T2Reading = Math.Round(ConvertHexToSingle(strData), 1)

                            End Select


                    'If _ResponseResult.Command = enmCommands.READ_DATE_TIME Then
                    '    strData = Hex(btData(4)).PadLeft(2, "0") & "#" & Hex(btData(5)).PadLeft(2, "0") & "#" & Hex(btData(6)).PadLeft(2, "0") & "#" & Hex(btData(7)).PadLeft(2, "0") & "#" & Hex(btData(8)).PadLeft(2, "0") & "#" & Hex(btData(9)).PadLeft(2, "0") & "#" & Hex(btData(10)).PadLeft(2, "0")
                    '    _ResponseResult.DeviceDateTime = GetDateTime(strData) 'Holds real time clock (RTC) data	7   Hex(btData(51)).PadLeft(2, "0")
                    'End If

                    'If _ResponseResult.Command = enmCommands.GET_REPORT_COUNT Then
                    '    strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0")
                    '    _ResponseResult.TotalReportCount = CDec("&h" & strData)

                    '    strData = Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")
                    '    _ResponseResult.TotalReadingCount = CDec("&h" & strData)
                    'End If

                    'If _ResponseResult.Command = enmCommands.GET_DEVICE_TEST_HISTORY Then
                    '    strData = Hex(btData(4)).PadLeft(2, "0") & Hex(btData(5)).PadLeft(2, "0")
                    '    _ResponseResult.InstrumentReportCount = CDec("&h" & strData)

                    '    strData = Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0")
                    '    _ResponseResult.OperationalHours = CDec("&h" & strData)
                    'End If

                    '_ResponseResult.Message = CDec("&h" & Hex(btData(46)).PadLeft(2, "0"))

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

                    'Select Case strBinary.Substring(4, 4)
                    '    Case "0000"
                    '        _ResponseResult.CH1Range = enmRange.HUNDREAD
                    '    Case "0001"
                    '        _ResponseResult.CH1Range = enmRange.TEN
                    '    Case "0010"
                    '        _ResponseResult.CH1Range = enmRange.ONE
                    '    Case "0100"
                    '        _ResponseResult.CH1Range = enmRange.DIVAND1
                    '    Case "1000"
                    '        _ResponseResult.CH1Range = enmRange.AUTORANGE
                    'End Select

                    'strBinary = ToBinary(CDec("&h" & Hex(btData(11)).PadLeft(2, "0")))

                    'Select Case strBinary.Substring(0, 4)
                    '    Case "0000"
                    '        _ResponseResult.CH2Range = enmRange.HUNDREAD
                    '    Case "0001"
                    '        _ResponseResult.CH2Range = enmRange.TEN
                    '    Case "0010"
                    '        _ResponseResult.CH2Range = enmRange.ONE
                    '    Case "0100"
                    '        _ResponseResult.CH2Range = enmRange.DIVAND1
                    '    Case "1000"
                    '        _ResponseResult.CH2Range = enmRange.AUTORANGE
                    'End Select

                    'Select Case strBinary.Substring(4, 4)
                    '    Case "0000"
                    '        _ResponseResult.CH3Range = enmRange.HUNDREAD
                    '    Case "0001"
                    '        _ResponseResult.CH3Range = enmRange.TEN
                    '    Case "0010"
                    '        _ResponseResult.CH3Range = enmRange.ONE
                    '    Case "0100"
                    '        _ResponseResult.CH3Range = enmRange.DIVAND1
                    '    Case "1000"
                    '        _ResponseResult.CH3Range = enmRange.AUTORANGE
                    'End Select

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

                        Case "1000"
                            _ResponseResult.TestModeSeeting = enmTestMode.DYNAMIC_RESISTANCE
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

                    strBinary = ToBinary(CDec("&h" & Hex(btData(59)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.WindingElment = clsAccessXWRM10A20.enmWindingElement.NO_ELEMENT
                        Case "0001"
                            _ResponseResult.WindingElment = clsAccessXWRM10A20.enmWindingElement.COPPER
                        Case "0010"
                            _ResponseResult.WindingElment = clsAccessXWRM10A20.enmWindingElement.ALUMINIUM
                    End Select

                    If strBinary.Substring(3, 1) = "1" Then
                        _ResponseResult.ProbeConnected = False
                    Else
                        _ResponseResult.ProbeConnected = True
                    End If

                    If strBinary.Substring(0, 1) = "1" Then
                        _ResponseResult.IsTempratureCorrection = True
                    Else
                        _ResponseResult.IsTempratureCorrection = False
                    End If

                    strData = Hex(btData(60)).PadLeft(2, "0") & Hex(btData(61)).PadLeft(2, "0") & Hex(btData(62)).PadLeft(2, "0") & Hex(btData(63)).PadLeft(2, "0")
                    _ResponseResult.AmbientTempratue = ConvertHexToSingle(strData)

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

                        Case "1000"
                            _ResponseResult.TestModeStatus = enmTestMode.DYNAMIC_RESISTANCE
                    End Select

                    _ResponseResult.TapNumber = CDec("&h" & Hex(btData(5)).PadLeft(2, "0"))

                    strData = Hex(btData(6)).PadLeft(2, "0") & Hex(btData(7)).PadLeft(2, "0") & Hex(btData(8)).PadLeft(2, "0") & Hex(btData(9)).PadLeft(2, "0")
                    _ResponseResult.HVTestCurrent = ConvertHexToSingle(strData)

                    strBinary = ToBinary(CDec("&h" & Hex(btData(10)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            _ResponseResult.CurrentUnit = enmCurrentsUnit.mAmpere
                            _ResponseResult.HVTestCurrentUnit = enmCurrentsUnit.mAmpere
                            _ResponseResult.HVTestCurrent = _ResponseResult.HVTestCurrent '* 1000
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
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 5)
                        Case Is < 10
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 4)
                        Case Is < 100
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 3)
                        Case Is < 1000
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 2)
                        Case Is < 10000
                            _ResponseResult.CH1Reading = Math.Round(_ResponseResult.CH1Reading, 1)
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
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 5)
                        Case Is < 10
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 4)
                        Case Is < 100
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 3)
                        Case Is < 1000
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 2)
                        Case Is < 10000
                            _ResponseResult.CH2Reading = Math.Round(_ResponseResult.CH2Reading, 1)
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
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 5)
                        Case Is < 10
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 4)
                        Case Is < 100
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 3)
                        Case Is < 1000
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 2)
                        Case Is < 10000
                            _ResponseResult.CH3Reading = Math.Round(_ResponseResult.CH3Reading, 1)
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
                    _ResponseResult.T1Reading = Math.Round(ConvertHexToSingle(strData), 1)

                    'Select Case _ResponseResult.T1Reading
                    '    Case Is < 1
                    '        _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 4)
                    '    Case Is < 10
                    '        _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 3)
                    '    Case Is < 100
                    '        _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 2)
                    '    Case Is < 1000
                    '        _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 1)
                    '    Case Is < 10000
                    '        _ResponseResult.T1Reading = Math.Round(_ResponseResult.T1Reading, 0)
                    'End Select

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
                    _ResponseResult.T2Reading = Math.Round(ConvertHexToSingle(strData), 1)

                    'Select Case _ResponseResult.T2Reading
                    '    Case Is < 1
                    '        _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 4)
                    '    Case Is < 10
                    '        _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 3)
                    '    Case Is < 100
                    '        _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 2)
                    '    Case Is < 1000
                    '        _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 1)
                    '    Case Is < 10000
                    '        _ResponseResult.T2Reading = Math.Round(_ResponseResult.T2Reading, 0)
                    'End Select

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

                    _ResponseResult.Corrected_CH1_Reading = GetCorrectedReading(_ResponseResult.CH1Reading, 75, _ResponseResult.T1Reading, _ResponseResult.WindingElment)

                    Select Case _ResponseResult.Corrected_CH1_Reading
                        Case Is < 1
                            _ResponseResult.Corrected_CH1_Reading = Math.Round(_ResponseResult.Corrected_CH1_Reading, 5)
                        Case Is < 10
                            _ResponseResult.Corrected_CH1_Reading = Math.Round(_ResponseResult.Corrected_CH1_Reading, 4)
                        Case Is < 100
                            _ResponseResult.Corrected_CH1_Reading = Math.Round(_ResponseResult.Corrected_CH1_Reading, 3)
                        Case Is < 1000
                            _ResponseResult.Corrected_CH1_Reading = Math.Round(_ResponseResult.Corrected_CH1_Reading, 2)
                        Case Is < 10000
                            _ResponseResult.Corrected_CH1_Reading = Math.Round(_ResponseResult.Corrected_CH1_Reading, 1)
                    End Select

                    _ResponseResult.Corrected_CH2_Reading = GetCorrectedReading(_ResponseResult.CH2Reading, 75, _ResponseResult.T1Reading, _ResponseResult.WindingElment)

                    Select Case _ResponseResult.Corrected_CH2_Reading
                        Case Is < 1
                            _ResponseResult.Corrected_CH2_Reading = Math.Round(_ResponseResult.Corrected_CH2_Reading, 5)
                        Case Is < 10
                            _ResponseResult.Corrected_CH2_Reading = Math.Round(_ResponseResult.Corrected_CH2_Reading, 4)
                        Case Is < 100
                            _ResponseResult.Corrected_CH2_Reading = Math.Round(_ResponseResult.Corrected_CH2_Reading, 3)
                        Case Is < 1000
                            _ResponseResult.Corrected_CH2_Reading = Math.Round(_ResponseResult.Corrected_CH2_Reading, 2)
                        Case Is < 10000
                            _ResponseResult.Corrected_CH2_Reading = Math.Round(_ResponseResult.Corrected_CH2_Reading, 1)
                    End Select

                    _ResponseResult.Corrected_CH3_Reading = GetCorrectedReading(_ResponseResult.CH3Reading, 75, _ResponseResult.T1Reading, _ResponseResult.WindingElment)

                    Select Case _ResponseResult.Corrected_CH3_Reading
                        Case Is < 1
                            _ResponseResult.Corrected_CH3_Reading = Math.Round(_ResponseResult.Corrected_CH3_Reading, 5)
                        Case Is < 10
                            _ResponseResult.Corrected_CH3_Reading = Math.Round(_ResponseResult.Corrected_CH3_Reading, 4)
                        Case Is < 100
                            _ResponseResult.Corrected_CH3_Reading = Math.Round(_ResponseResult.Corrected_CH3_Reading, 3)
                        Case Is < 1000
                            _ResponseResult.Corrected_CH3_Reading = Math.Round(_ResponseResult.Corrected_CH3_Reading, 2)
                        Case Is < 10000
                            _ResponseResult.Corrected_CH3_Reading = Math.Round(_ResponseResult.Corrected_CH3_Reading, 1)
                    End Select
            End Select

            RaiseEvent ResponseRecieved(_ResponseResult)
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub ProcessReportData(ByVal colCollection As Collection)
        Dim dtDataTable As New Data.DataTable

        Try
            'Dim strResponseString As String = vbNullString

            'For iSubCount As Integer = 1 To (colCollection.Count)
            '    strResponseString = strResponseString & "#" & Hex(colCollection.Item(iSubCount)).PadLeft(2, "0")
            'Next

            'RaiseEvent SendResponseFrame(strResponseString)

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
            dtDataTable.Columns.Add("CH1Readings", GetType(System.String))
            dtDataTable.Columns.Add("CH1Unit", GetType(System.String))
            dtDataTable.Columns.Add("CH1ReadingDP", GetType(System.Double))

            dtDataTable.Columns.Add("CH2Reading", GetType(System.Double))
            dtDataTable.Columns.Add("CH2Readings", GetType(System.String))
            dtDataTable.Columns.Add("CH2Unit", GetType(System.String))
            dtDataTable.Columns.Add("CH2ReadingDP", GetType(System.Double))

            dtDataTable.Columns.Add("CH3Reading", GetType(System.Double))
            dtDataTable.Columns.Add("CH3Readings", GetType(System.String))
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

            dtDataTable.Columns.Add("ElapsedTime", GetType(System.Double))

            dtDataTable.Columns.Add("WindingElement", GetType(System.Double))

            dtDataTable.Columns.Add("CH1ReadingCorrected", GetType(System.Double))
            dtDataTable.Columns.Add("CH2ReadingCorrected", GetType(System.Double))
            dtDataTable.Columns.Add("CH3ReadingCorrected", GetType(System.Double))

            dtDataTable.Columns.Add("CH1ReadingsCorrected", GetType(System.String))
            dtDataTable.Columns.Add("CH2ReadingsCorrected", GetType(System.String))
            dtDataTable.Columns.Add("CH3ReadingsCorrected", GetType(System.String))

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

            Dim PreviousHRElapseValue As Integer = 0

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

                        Case "1000"
                            _TestMode = "DYNAMIC RESISTANCE"
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

                    If _TestMode = "NORMAL" Then
                        drNewRow.Item("TapNumber") = 0
                    End If

                    strData = Hex(btsubDataNew(6)).PadLeft(2, "0") & Hex(btsubDataNew(7)).PadLeft(2, "0") & Hex(btsubDataNew(8)).PadLeft(2, "0") & Hex(btsubDataNew(9)).PadLeft(2, "0")

                    drNewRow.Item("TestCurrentReading") = ConvertHexToSingle(strData)

                    Dim strBinary As String = ToBinary(CDec("&h" & Hex(btsubDataNew(10)).PadLeft(2, "0")))

                    Select Case strBinary.Substring(4, 4)
                        Case "0000"
                            drNewRow.Item("TestCurrentUnit") = "mAmpere"
                            drNewRow.Item("TestCurrentReading") = Val(drNewRow.Item("TestCurrentReading"))
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
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 5)
                        Case Is < 10
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 4)
                        Case Is < 100
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 3)
                        Case Is < 1000
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 2)
                        Case Is < 10000
                            drNewRow.Item("CH1Reading") = Math.Round(drNewRow.Item("CH1Reading"), 1)
                    End Select

                    Select Case drNewRow.Item("CH1Reading")
                        Case 333333344
                            drNewRow.Item("CH1Readings") = "Rev"

                        Case 555555584.0
                            drNewRow.Item("CH1Readings") = "???"

                        Case 111111112.0
                            drNewRow.Item("CH1Readings") = "U L"

                        Case 1000000000
                            drNewRow.Item("CH1Readings") = "O L"

                        Case Else
                            drNewRow.Item("CH1Readings") = drNewRow.Item("CH1Reading")
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
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 5)
                        Case Is < 10
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 4)
                        Case Is < 100
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 3)
                        Case Is < 1000
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 2)
                        Case Is < 10000
                            drNewRow.Item("CH2Reading") = Math.Round(drNewRow.Item("CH2Reading"), 1)
                    End Select

                    Select Case drNewRow.Item("CH2Reading")
                        Case 333333344
                            drNewRow.Item("CH2Readings") = "Rev"

                        Case 555555584.0
                            drNewRow.Item("CH2Readings") = "???"

                        Case 111111112.0
                            drNewRow.Item("CH2Readings") = "U L"

                        Case 1000000000
                            drNewRow.Item("CH2Readings") = "O L"

                        Case Else
                            drNewRow.Item("CH2Readings") = drNewRow.Item("CH2Reading")
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
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 5)
                        Case Is < 10
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 4)
                        Case Is < 100
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 3)
                        Case Is < 1000
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 2)
                        Case Is < 10000
                            drNewRow.Item("CH3Reading") = Math.Round(drNewRow.Item("CH3Reading"), 1)
                    End Select

                    Select Case drNewRow.Item("CH3Reading")
                        Case 333333344
                            drNewRow.Item("CH3Readings") = "Rev"

                        Case 555555584.0
                            drNewRow.Item("CH3Readings") = "???"

                        Case 111111112.0
                            drNewRow.Item("CH3Readings") = "U L"

                        Case 1000000000
                            drNewRow.Item("CH3Readings") = "O L"

                        Case Else
                            drNewRow.Item("CH3Readings") = drNewRow.Item("CH3Reading")
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
                    drNewRow.Item("Temp1") = Math.Round(ConvertHexToSingle(strData), 1)

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
                    drNewRow.Item("Temp2") = Math.Round(ConvertHexToSingle(strData), 1)

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

                    If drNewRow.Item("TestMode") = "HR" Then
                        If Val(drNewRow.Item("TestResultNo")) = 1 Then
                            strData = Hex(btsubDataNew(43)).PadLeft(2, "0") & Hex(btsubDataNew(44)).PadLeft(2, "0")
                            drNewRow.Item("ElapsedTime") = CDec("&h" & strData)
                            PreviousHRElapseValue = CDec("&h" & strData)
                        Else
                            drNewRow.Item("ElapsedTime") = PreviousHRElapseValue
                        End If
                    Else
                        drNewRow.Item("ElapsedTime") = 0
                    End If

                    'WindingElement
                    strData = Hex(btsubDataNew(42)).PadLeft(2, "0")

                    drNewRow.Item("WindingElement") = CDec("&h" & strData)

                    drNewRow.Item("CH1ReadingCorrected") = GetCorrectedReading(drNewRow.Item("CH1Reading"), 75, drNewRow.Item("Temp1"), Val(drNewRow.Item("WindingElement").ToString))

                    Select Case Val(drNewRow.Item("CH1ReadingCorrected"))
                        Case Is < 1
                            drNewRow.Item("CH1ReadingCorrected") = Math.Round(drNewRow.Item("CH1ReadingCorrected"), 5)
                        Case Is < 10
                            drNewRow.Item("CH1ReadingCorrected") = Math.Round(drNewRow.Item("CH1ReadingCorrected"), 4)
                        Case Is < 100
                            drNewRow.Item("CH1ReadingCorrected") = Math.Round(drNewRow.Item("CH1ReadingCorrected"), 3)
                        Case Is < 1000
                            drNewRow.Item("CH1ReadingCorrected") = Math.Round(drNewRow.Item("CH1ReadingCorrected"), 2)
                        Case Is < 10000
                            drNewRow.Item("CH1ReadingCorrected") = Math.Round(drNewRow.Item("CH1ReadingCorrected"), 1)
                    End Select

                    Select Case drNewRow.Item("CH1ReadingCorrected")

                    End Select

                    drNewRow.Item("CH2ReadingCorrected") = GetCorrectedReading(drNewRow.Item("CH2Reading"), 75, drNewRow.Item("Temp1"), Val(drNewRow.Item("WindingElement").ToString))

                    Select Case Val(drNewRow.Item("CH2ReadingCorrected"))
                        Case Is < 1
                            drNewRow.Item("CH2ReadingCorrected") = Math.Round(drNewRow.Item("CH2ReadingCorrected"), 5)
                        Case Is < 10
                            drNewRow.Item("CH2ReadingCorrected") = Math.Round(drNewRow.Item("CH2ReadingCorrected"), 4)
                        Case Is < 100
                            drNewRow.Item("CH2ReadingCorrected") = Math.Round(drNewRow.Item("CH2ReadingCorrected"), 3)
                        Case Is < 1000
                            drNewRow.Item("CH2ReadingCorrected") = Math.Round(drNewRow.Item("CH2ReadingCorrected"), 2)
                        Case Is < 10000
                            drNewRow.Item("CH2ReadingCorrected") = Math.Round(drNewRow.Item("CH2ReadingCorrected"), 1)
                    End Select

                    drNewRow.Item("CH3ReadingCorrected") = GetCorrectedReading(drNewRow.Item("CH3Reading"), 75, drNewRow.Item("Temp1"), Val(drNewRow.Item("WindingElement").ToString))

                    Select Case Val(drNewRow.Item("CH3ReadingCorrected"))
                        Case Is < 1
                            drNewRow.Item("CH3ReadingCorrected") = Math.Round(drNewRow.Item("CH3ReadingCorrected"), 5)
                        Case Is < 10
                            drNewRow.Item("CH3ReadingCorrected") = Math.Round(drNewRow.Item("CH3ReadingCorrected"), 4)
                        Case Is < 100
                            drNewRow.Item("CH3ReadingCorrected") = Math.Round(drNewRow.Item("CH3ReadingCorrected"), 3)
                        Case Is < 1000
                            drNewRow.Item("CH3ReadingCorrected") = Math.Round(drNewRow.Item("CH3ReadingCorrected"), 2)
                        Case Is < 10000
                            drNewRow.Item("CH3ReadingCorrected") = Math.Round(drNewRow.Item("CH3ReadingCorrected"), 1)
                    End Select

                    Select Case drNewRow.Item("CH1Readings")
                        Case "Rev"
                            drNewRow.Item("CH1ReadingsCorrected") = "Rev"

                        Case "???"
                            drNewRow.Item("CH1ReadingsCorrected") = "???"

                        Case "U L"
                            drNewRow.Item("CH1ReadingsCorrected") = "U L"

                        Case "O L"
                            drNewRow.Item("CH1ReadingsCorrected") = "O L"

                        Case Else
                            drNewRow.Item("CH1ReadingsCorrected") = drNewRow.Item("CH1ReadingCorrected")

                    End Select

                    Select Case drNewRow.Item("CH2Readings")

                        Case "Rev"
                            drNewRow.Item("CH2ReadingsCorrected") = "Rev"

                        Case "???"
                            drNewRow.Item("CH2ReadingsCorrected") = "???"

                        Case "U L"
                            drNewRow.Item("CH2ReadingsCorrected") = "U L"

                        Case "O L"
                            drNewRow.Item("CH2ReadingsCorrected") = "O L"

                        Case Else
                            drNewRow.Item("CH2ReadingsCorrected") = drNewRow.Item("CH2ReadingCorrected")
                    End Select

                    Select Case drNewRow.Item("CH3Readings")
                        Case "Rev"
                            drNewRow.Item("CH3ReadingsCorrected") = "Rev"

                        Case "???"
                            drNewRow.Item("CH3ReadingsCorrected") = "???"

                        Case "U L"
                            drNewRow.Item("CH3ReadingsCorrected") = "U L"

                        Case "O L"
                            drNewRow.Item("CH3ReadingsCorrected") = "O L"

                        Case Else
                            drNewRow.Item("CH3ReadingsCorrected") = drNewRow.Item("CH3ReadingCorrected")
                    End Select

                    dtDataTable.Rows.Add(drNewRow)
                End If

                nextReportCount = nextReportCount - 1
            Loop

            RaiseEvent ReportDataRecieved(dtDataTable)

        Catch ex As Exception
            RaiseEvent ReportDataRecieved(dtDataTable)
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

        _ResponseReceived = False

        Try

            'Dim strResponseString As String = vbNullString

            'For iSubCount As Integer = 0 To (strCommandFrame.Length - 1)
            '    strResponseString = strResponseString & "#" & strCommandFrame(iSubCount)
            'Next

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

Public Class Response
    Public CurrentValue As Double

    Public CurrentUnit As clsAccessXWRM10A20.enmCurrentsUnit
    Public ChannelSelected As clsAccessXWRM10A20.enmChannels
    Public CH1Range As clsAccessXWRM10A20.enmRange
    Public CH2Range As clsAccessXWRM10A20.enmRange
    Public CH3Range As clsAccessXWRM10A20.enmRange
    Public TestModeSeeting As clsAccessXWRM10A20.enmTestMode
    Public SaveResult As Boolean
    Public HrTestIime As Double
    Public DUTSerialNumber As String
    Public DUTType As String
    Public DUTLocation As String
    Public DeviceDateTime As DateTime

    Public TestModeStatus As clsAccessXWRM10A20.enmTestMode
    Public TapNumber As Integer = 0

    Public HVTestCurrent As Double = 0.0
    Public LVTestCurrent As Double = 0.0

    Public HVTestCurrentUnit As enmCurrentsUnit
    Public LVTestCurrentUnit As enmCurrentsUnit

    Public TestCurrentDP As Integer = 0
    Public CH1Reading As Double

    Public HVReading As Double
    Public LVReading As Double

    Public CH1Unit As enmResistanceUnit

    Public HVUnit As enmResistanceUnit
    Public LVUnit As enmResistanceUnit

    Public Winding As enmWindingElement
    Public WindingSelection As Integer

    Public CH1DP As Integer = 0
    Public CH2Reading As Double
    Public CH2Unit As enmResistanceUnit
    Public CH2DP As Integer = 0
    Public CH3Reading As Double
    Public CH3Unit As enmResistanceUnit
    Public CH3DP As Integer = 0
    Public T1Reading As Double
    Public T1Unit As enmTempratureUnit
    Public T1DP As Integer = 0
    Public T2Reading As Double
    Public T2Unit As enmTempratureUnit
    Public T2DP As Integer = 0

    Public Command As clsAccessXWRM10A20.enmCommands
    Public Message As enmErrorMessage = enmErrorMessage.SUCCESS

    Public TotalReportCount As Integer = 0
    Public TotalReadingCount As Integer = 0

    'Optional one
    Public InstrumentReportCount As Double = 0
    Public OperationalHours As Double = 0

    Public TestON As Boolean = False
    Public MemoryFull As Boolean = False
    Public CommunicationBreak = False

    Public IsHighest As Boolean = False

    Public WindingElment As enmWindingElement
    Public ProbeConnected As Boolean = False
    Public IsTempratureCorrection As Boolean = False
    Public AmbientTempratue As Double = 0

    Public Corrected_CH1_Reading As Double = 0
    Public Corrected_CH2_Reading As Double = 0
    Public Corrected_CH3_Reading As Double = 0

    Public HV_LV_SELECTION As Integer 'oltc
    Public CURRENT As Double
    Public CURRENT_UNIT As Integer
    Public READING As Double
    Public READING_UNIT As Integer

End Class
