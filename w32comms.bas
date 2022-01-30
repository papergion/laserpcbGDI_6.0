Attribute VB_Name = "W32Comms"
' COMM stuff from WIN32API.TXT

'  Serial provider type.
Const SP_SERIALCOMM = &H1&

'  Provider SubTypes
Const PST_UNSPECIFIED = &H0&
Const PST_RS232 = &H1&
Const PST_PARALLELPORT = &H2&
Const PST_RS422 = &H3&
Const PST_RS423 = &H4&
Const PST_RS449 = &H5&
Const PST_FAX = &H21&
Const PST_SCANNER = &H22&
Const PST_NETWORK_BRIDGE = &H100&
Const PST_LAT = &H101&
Const PST_TCPIP_TELNET = &H102&
Const PST_X25 = &H103&

'  Provider capabilities flags.
Const PCF_DTRDSR = &H1&
Const PCF_RTSCTS = &H2&
Const PCF_RLSD = &H4&
Const PCF_PARITY_CHECK = &H8&
Const PCF_XONXOFF = &H10&
Const PCF_SETXCHAR = &H20&
Const PCF_TOTALTIMEOUTS = &H40&
Const PCF_INTTIMEOUTS = &H80&
Const PCF_SPECIALCHARS = &H100&
Const PCF_16BITMODE = &H200&

'  Comm provider settable parameters.
Const SP_PARITY = &H1&
Const SP_BAUD = &H2&
Const SP_DATABITS = &H4&
Const SP_STOPBITS = &H8&
Const SP_HANDSHAKING = &H10&
Const SP_PARITY_CHECK = &H20&
Const SP_RLSD = &H40&

'  Settable baud rates in the provider.
Const BAUD_075 = &H1&
Const BAUD_110 = &H2&
Const BAUD_134_5 = &H4&
Const BAUD_150 = &H8&
Const BAUD_300 = &H10&
Const BAUD_600 = &H20&
Const BAUD_1200 = &H40&
Const BAUD_1800 = &H80&
Const BAUD_2400 = &H100&
Const BAUD_4800 = &H200&
Const BAUD_7200 = &H400&
Const BAUD_9600 = &H800&
Const BAUD_14400 = &H1000&
Const BAUD_19200 = &H2000&
Const BAUD_38400 = &H4000&
Const BAUD_56K = &H8000&
Const BAUD_128K = &H10000
Const BAUD_115200 = &H20000
Const BAUD_57600 = &H40000
Const BAUD_USER = &H10000000

'  Settable Data Bits
Const DATABITS_5 = &H1&
Const DATABITS_6 = &H2&
Const DATABITS_7 = &H4&
Const DATABITS_8 = &H8&
Const DATABITS_16 = &H10&
Const DATABITS_16X = &H20&

'  Settable Stop and Parity bits.
Const STOPBITS_10 = &H1&
Const STOPBITS_15 = &H2&
Const STOPBITS_20 = &H4&
Const PARITY_NONE = &H100&
Const PARITY_ODD = &H200&
Const PARITY_EVEN = &H400&
Const PARITY_MARK = &H800&
Const PARITY_SPACE = &H1000&

Type COMMPROP
        wPacketLength As Integer
        wPacketVersion As Integer
        dwServiceMask As Long
        dwReserved1 As Long
        dwMaxTxQueue As Long
        dwMaxRxQueue As Long
        dwMaxBaud As Long
        dwProvSubType As Long
        dwProvCapabilities As Long
        dwSettableParams As Long
        dwSettableBaud As Long
        wSettableData As Integer
        wSettableStopParity As Integer
        dwCurrentTxQueue As Long
        dwCurrentRxQueue As Long
        dwProvSpec1 As Long
        dwProvSpec2 As Long
        wcProvChar(1) As Integer
End Type

'Type COMSTAT
'        fCtsHold As Long
'        fDsrHold As Long
'        fRlsdHold As Long
'        fXoffHold As Long
'        fXoffSent As Long
'        fEof As Long
'        fTxim As Long
'        fReserved As Long
'        cbInQue As Long
'        cbOutQue As Long
'End Type

Type COMSTAT
        fBitFields As Long 'See Comment in Win32API.Txt
        cbInQue As Long
        cbOutQue As Long
End Type
' The eight actual COMSTAT bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName     Bit #     Description
' ---------     -----     ---------------------------
' fCtsHold        1       Tx waiting for CTS signal
' fDsrHold        2       Tx waiting for DSR signal
' fRlsdHold       3       Tx waiting for RLSD signal
' fXoffHold       4       Tx waiting, XOFF char rec'd
' fXoffSent       5       Tx waiting, XOFF char sent
' fEof            6       EOF character sent
' fTxim           7       character waiting for Tx
' fReserved       8       reserved (25 bits)

'  DTR Control Flow Values.
Const DTR_CONTROL_DISABLE = &H0
Const DTR_CONTROL_ENABLE = &H1
Const DTR_CONTROL_HANDSHAKE = &H2

'  RTS Control Flow Values
Const RTS_CONTROL_DISABLE = &H0
Const RTS_CONTROL_ENABLE = &H1
Const RTS_CONTROL_HANDSHAKE = &H2
Const RTS_CONTROL_TOGGLE = &H3

'Type DCB
'        DCBlength As Long
'        BaudRate As Long
'        fBinary As Long
'        fParity As Long
'        fOutxCtsFlow As Long
'        fOutxDsrFlow As Long
'        fDtrControl As Long
'        fDsrSensitivity As Long
'        fTXContinueOnXoff As Long
'        fOutX As Long
'        fInX As Long
'        fErrorChar As Long
'        fNull As Long
'        fRtsControl As Long
'        fAbortOnError As Long
'        fDummy2 As Long
'        wReserved As Integer
'        XonLim As Integer
'        XoffLim As Integer
'        ByteSize As Byte
'        Parity As Byte
'        StopBits As Byte
'        XonChar As Byte
'        XoffChar As Byte
'        ErrorChar As Byte
'        EofChar As Byte
'        EvtChar As Byte
'End Type

Type dcb
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long 'See Comments in Win32API.Txt
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer 'Reserved; Do Not Use
End Type
' The fourteen actual DCB bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName             Bit #     Description
' -----------------     -----     ------------------------------
' fBinary                 1       binary mode, no EOF check
' fParity                 2       enable parity checking
' fOutxCtsFlow            3       CTS output flow control
' fOutxDsrFlow            4       DSR output flow control
' fDtrControl             5       DTR flow control type (2 bits)
' fDsrSensitivity         7       DSR sensitivity
' fTXContinueOnXoff       8       XOFF continues Tx
' fOutX                   9       XON/XOFF out flow control
' fInX                   10       XON/XOFF in flow control
' fErrorChar             11       enable error replacement
' fNull                  12       enable null stripping
' fRtsControl            13       RTS flow control (2 bits)
' fAbortOnError          15       abort reads/writes on error
' fDummy2                16       reserved

Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type


' COMM declarations
Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As dcb) As Long
Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As dcb) As Long
Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As dcb) As Long
Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As dcb, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function TransmitCommChar Lib "kernel32" (ByVal nCid As Long, ByVal cChar As Byte) As Long
Declare Function SetCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Declare Function ClearCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Declare Function EscapeCommFunction Lib "kernel32" (ByVal nCid As Long, ByVal nFunc As Long) As Long
Declare Function GetCommMask Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long) As Long
Declare Function GetCommProperties Lib "kernel32" (ByVal hFile As Long, lpCommProp As COMMPROP) As Long
Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, lpModemStat As Long) As Long
'Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long

