Attribute VB_Name = "SW_EN_DIS_CHECK"
Sub SW_EN_CHECK()
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Dim slaveAddress As Byte
    Dim regAddress As Integer 'Read in new value for each row
    Dim regData As Byte
    Dim STARTROW: STARTROW = 1
    Dim ENDROW: ENDROW = 245
    Dim REGADDR_COL: REGADDR_COL = 2
    Dim regAddrStr As String
    Dim SWENADDRESS: SWENADDRESS = &H100
    Dim DATAOUTPUT_COL: DATAOUTPUT_COL = 13
    slaveAddress = &H74
    
    regData = &HFF
    
    'Loop through every reg address writing to them
    For row = STARTROW To ENDROW
        regAddrStr = Cells(row, REGADDR_COL).value 'Read in row's reg address
        regAddress = hexStringToInt(regAddrStr) 'Convert to regAddr integer value
        
        If regAddress = SWENADDRESS Then
            'don't write to SW_EN
        Else
            Call i2c.I2CWriteByte16bit(slaveAddress, regAddress, regData) ' write to row's reg address
        End If
    
    Next row
    
    'SW_EN = 0 then 1 -- reset the device
    Dim i: i = 5
    
    'Loop through every reg address reading values
    For row = STARTROW To ENDROW
        regAddrStr = Cells(row, REGADDR_COL).value 'Read in row's reg address
        regAddress = hexStringToInt(regAddrStr) 'Convert to regAddr integer value
        
        Call i2c.I2CReadByte16bit(slaveAddress, regAddress, regData)
        
        
        
    Next row
    
    
End Sub

Function hexStringToInt(in_string As String)
    'Convert hex string to int and return int
    'eg Convert "0x0a" to 10 (int)
    
End Function
