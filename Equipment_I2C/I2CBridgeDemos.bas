Attribute VB_Name = "I2CBridgeDemos"
Sub I2CBridgeConnected()    'Check if I2C Interface Board is connected
Attribute I2CBridgeConnected.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    If i2c.Connect() = True Then
        'ActiveCell.value = "Connected"
        Debug.Print "Connected"
    Else
        'ActiveCell.value = "Disconnected"
        Debug.Print "Disconnected"
    End If
End Sub

Sub SearchForDevices()      'find slave addresses present on I2C bus
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim addressList(0 To 60) As Byte
    Dim numFound As Byte
    
    If (i2c.I2CDeviceSearch(numFound, addressList) = True) Then
        ActiveCell.value = numFound
        For i = 0 To (numFound - 1)
            Cells((ActiveCell.row + i), (ActiveCell.Column + 1)).value = "0x" & Hex(addressList(i))
        Next i
    End If
End Sub

Sub CheckIfDeviceIsPresent()    'check whether or not a device is present on I2C bus
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim slaveAddress As Byte
    
    slaveAddress = &H74 'set slave address we want to find
    
    If i2c.I2CDevicePresent(slaveAddress) Then
        ActiveCell.value = "Device at 0x" & Hex(slaveAddress) & " is present"
    Else
        ActiveCell.value = "Device at 0x" & Hex(slaveAddress) & " is absent"
    End If
End Sub

Sub ReadRevId()     'Write/read using base I2CWriteAndRead function
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim slaveAddress As Byte
    Dim writeBuffer(0 To 2) As Byte
    Dim readBuffer(0 To 2) As Byte
    
    slaveAddress = &H74
    writeBuffer(0) = &H1
    writeBuffer(1) = &HFF
    
    'Prototype for I2CWriteAndRead function:
    'bool I2CWriteAndRead(byte slaveAddress, byte* writeBuffer, byte* readBuffer, byte writeCount, byte readCount);
    
    If i2c.I2CWriteAndRead(slaveAddress, writeBuffer, readBuffer, 2, 1) = True Then
        ActiveCell.value = "0x" & Hex(readBuffer(0))
    Else
        ActiveCell.value = "Error"
    End If
End Sub

Sub WriteReg()      '16-bit register write
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim slaveAddress As Byte
    Dim regAddress As Integer
    Dim regData As Byte
    
    slaveAddress = &H64
    regAddress = &H19
    regData = &HAA
    
    If i2c.I2CWriteByte16bit(slaveAddress, regAddress, regData) = True Then
        ActiveCell.value = "Success"
    Else
        ActiveCell.value = "Error"
    End If
End Sub

Sub ReadReg()       '16-bit register read
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim slaveAddress As Byte
    Dim regAddress As Integer
    Dim regData As Byte
    
    slaveAddress = &H64
    regAddress = &H19
    regData = 0
    
    If i2c.I2CReadByte16bit(slaveAddress, regAddress, regData) = True Then
        ActiveCell.value = "Read: 0x" & Hex(regData)
    Else
        ActiveCell.value = "Error"
    End If
End Sub

Sub SetBaudRate()   'Set I2C baud rate (SCL frequency)
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim desiredRate As Double
    Dim actualRate As Double
    
    desiredRate = 1000000#      'Try to set to 1MHz (1Mbps)
    
    actualRate = i2c.I2CSetBaudRate(desiredRate)
    
    If actualRate < 0 Then
        ActiveCell.value = "Error setting baud rate"
    Else
        ActiveCell.value = actualRate
    End If
End Sub

Sub GetBaudRate()   'Read I2C baud rate (SCL frequency)
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim actualRate As Double
    
    actualRate = i2c.I2CGetBaudRate()
    
    If actualRate < 0 Then
        ActiveCell.value = "Error getting baud rate"
    Else
        ActiveCell.value = actualRate   'actualRate is in bps
    End If
End Sub
































