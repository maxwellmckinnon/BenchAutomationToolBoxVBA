Attribute VB_Name = "ADC_PSRR"
Sub ADC_PSRR_HZ03()
    'Test the rejection ripple on the measurement ADC
    'Read the ADCs over I2C many times and make a histogram
    
    iterations = 1000
    
    rowstart = 37
    COL_iter = 1
    COL_pvdd = 2
    COL_adc = 4
    
    devaddr = &H62
    PVDDADC_hi = &H20
    PVDDADC_lo = &H54
    THERMADC_hi = &H20
    THERMADC_lo = &H55
    
    Call ADC_PSRR_HZ03_FormatHeader
    
    Dim readback As Byte
    If Not Equipment_I2C.I2C_Controls_.I2C_bridge_16Bit_Read_Control(devaddr, PVDDADC_hi, PVDDADC_lo, readback) Then MsgBox "Check I2C Connection! Failure!"
    
    r = rowstart
    For i = 0 To iterations
        DoEvents
        If Not Equipment_I2C.I2C_Controls_.I2C_bridge_16Bit_Read_Control(devaddr, PVDDADC_hi, PVDDADC_lo, readback) Then MsgBox "Check I2C Connection! Failure!"
        pvdd_read = readback
        If Not Equipment_I2C.I2C_Controls_.I2C_bridge_16Bit_Read_Control(devaddr, THERMADC_hi, THERMADC_lo, readback) Then MsgBox "Check I2C Connection! Failure!"
        therm_read = readback
        
        Cells(r, COL_iter).value = r - rowstart + 1
        Cells(r, COL_iter + 2).value = r - rowstart + 1
        Cells(r, COL_iter + 4).value = r - rowstart + 1
        Cells(r, COL_iter + 6).value = r - rowstart + 1
        
        Cells(r, COL_pvdd).value = pvdd_read
        Cells(r, COL_adc).value = therm_read
        r = r + 1
    Next i
    
End Sub

Private Sub ADC_PSRR_HZ03_FormatHeader()
    Cells(35, 1).value = "PVDD raw read"
    Cells(35, 3).value = "Therm raw read)"
End Sub
