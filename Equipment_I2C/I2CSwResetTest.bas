Attribute VB_Name = "I2CSwResetTest"
Sub I2CSW_RESET()
    'Open immediate window for debugging output - Ctrl + G
    
    'Write a 1 to all registers, write 1 to sw reset
    'Readback all registers
    'Compare against POR value
    
    'POR Default as string stored in column L (paste in from regmap.com)
    'Find POR Value (DEC)
    'Compare against read back value column N
    'Report status in column O
    
    Call ReadConvertPORString
    Call WriteOnesAndReset
    Call readbackRegisters
    
End Sub

Sub EnterTestMode()
    Dim DEVADDRI2C As Integer: DEVADDRI2C = &H74
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    Dim RevIDReg As Integer: RevIDReg = &H1FF
    
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, RevIDReg, &H54)
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, RevIDReg, &H4D)
    
    Call I2CBridgeDemos.I2CBridgeConnected
    
End Sub

Sub readbackRegisters()
    'readback all registers and dump to column N
    Dim LIMIT As Integer: LIMIT = 1000 'Stop loop if this count is exceeded
    Dim done As Boolean: done = False
    Dim count As Integer: count = 0
    Dim row As Integer: row = count + 2
    Dim hexAddressCol As Integer: hexAddressCol = 2 ' Col B for hex addresses
    Dim CellReadValue As String
    Dim regHexAddr As Integer
    Dim SWResetRegAddr As Integer: SWResetRegAddr = &H100
    Dim ColWrite As Integer: ColWrite = 14 ' 14 = N
    Dim CellWriteValue As String
    Dim Readback As Integer
    
    Dim DEVADDRI2C As Integer: DEVADDRI2C = &H74
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Call EnterTestMode
    Call I2CBridgeDemos.I2CBridgeConnected
    
    Do While (done <> True) Or (count > LIMIT)
        DoEvents
        If IsEmpty(Cells(row, hexAddressCol)) Then
            done = True
            Exit Do
        End If
        CellReadValue = Cells(row, hexAddressCol).value
        regHexAddr = Val("&H" + Mid(CellReadValue, 3, 4)) ' grab 4 chars after the "0x"
        'Convert the hex string to an integer to write to the device
        
        Call i2c.I2CReadByte16bit(DEVADDRI2C, regHexAddr, Readback) '
        Cells(row, ColWrite).value = Readback
        count = count + 1
        row = count + 2
    Loop
    
    Call i2c.I2CReadByte16bit(DEVADDRI2C, testReadbackReg, testReadbackData)
    If testReadbackData = 1 Then
        Debug.Print "Written successfully"
    Else
        Debug.Print "Fail to write"
    End If
    
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, SWResetRegAddr, &H1)
    
End Sub

Sub WriteOnesAndReset()
    'Write ones to all the registers except the SW reset ( 0x0100)
    Dim LIMIT As Integer: LIMIT = 1000 'Stop loop if this count is exceeded
    Dim done As Boolean: done = False
    Dim count As Integer: count = 0
    Dim row As Integer: row = count + 2
    Dim hexAddressCol As Integer: hexAddressCol = 2 ' Col B for hex addresses
    Dim CellReadValue As String
    Dim regHexAddr As Integer
    Dim SWResetRegAddr As Integer: SWResetRegAddr = &H100
    Dim testReadbackReg As Integer: testReadbackReg = &HC
    Dim testReadbackData As Integer
    
    Dim DEVADDRI2C As Integer: DEVADDRI2C = &H74
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    'Call EnterTestMode
    Call I2CBridgeDemos.I2CBridgeConnected
    
    Do While (done <> True) Or (count > LIMIT)
        DoEvents
        If IsEmpty(Cells(row, hexAddressCol)) Then
            done = True
            Exit Do
        End If
        CellReadValue = Cells(row, hexAddressCol).value
        regHexAddr = Val("&H" + Mid(CellReadValue, 3, 4)) ' grab 4 chars after the "0x"
        'Convert the hex string to an integer to write to the device
        
        If regHexAddr = SWResetRegAddr Then
            'Don't do the write
            LIMIT = LIMIT ' just dummy breakpoint
        Else
            Call i2c.I2CWriteByte16bit(DEVADDRI2C, regHexAddr, &H1) ' Write ones to every register address
        End If
        count = count + 1
        row = count + 2
    Loop
    
    Call i2c.I2CReadByte16bit(DEVADDRI2C, testReadbackReg, testReadbackData)
    If testReadbackData = 1 Then
        Debug.Print "Written successfully"
    Else
        Debug.Print "Fail to write"
    End If
    
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, SWResetRegAddr, &H1)
    
End Sub

Sub ReadConvertPORString()
    'Read Column L and convert to Column M
    Dim LIMIT As Integer: LIMIT = 1000 'Stop loop if this count is exceeded
    Dim done As Boolean: done = False
    Dim count As Integer: count = 0
    Dim row As Integer: row = count + 2
    Dim ColRead As Integer: ColRead = 12 ' 12 = L
    Dim ColWrite As Integer: ColWrite = 13 ' 13 = M
    Dim CellReadValue As String
    Dim CellWriteValue As String
    
    Do While (done <> True) Or (count > LIMIT)
        DoEvents
        If IsEmpty(Cells(row, ColRead)) Then
            done = True
            Exit Do
        End If
        CellReadValue = Cells(row, ColRead).value
        CellWriteValue = ConvertPORStringToInt(CellReadValue)
        Cells(row, ColWrite).value = CellWriteValue
        
        row = row + 1
        count = count + 1
    Loop
    
End Sub

Function ConvertPORStringToInt(PORString As String) As Integer
    'Convert "0000_0000" to Int, return Int
    Dim MSBs As String
    Dim LSBs As String
    Dim MSB As Integer
    Dim LSB As Integer
    
    MSBs = Mid(PORString, 1, 4)
    MSB = DataManip.Bin2Dec(MSBs)
    LSBs = Mid(PORString, 6, 4)
    LSB = DataManip.Bin2Dec(LSBs)
    ConvertPORStringToInt = MSB * 16 + LSB
    
End Function
