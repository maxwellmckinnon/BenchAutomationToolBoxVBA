Attribute VB_Name = "BestBoost_Quiescent"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub Quiescent_AX80()
    'Measure the quiescent VBAT current across boost settings
    
    CURR_GPIB = "GPIB::11" ' Fluke 8854A
    
    BOARD = Cells(5, 21).value
    
    DEV = &H74
    
    DT_REG = &HCF ' deadtime register
    DT_Default = &H5  ' Best known from AX80 A0
    SR_REG = &HD2 ' Slew rate register
    SR_Default = &H3
    
    
    OUTPUTDATACELL_x = 37
    OUTPUTDATACELL_Y = 18
    
    Dim inVoltage As Double
    Dim inCurrent As Double
    
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, SR_REG, SR_Default)
    
    Cells(OUTPUTDATACELL_x - 2, OUTPUTDATACELL_Y).value = "Quiescent Current VBAT"
    For cf = 0 To 16
        DoEvents
        X = OUTPUTDATACELL_x + cf
        y = OUTPUTDATACELL_Y
        Cells(X, y).value = "cf = " & CStr(cf)
        Cells(X, y + 1).value = cf
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, DT_REG, cf)
        
        Sleep (1000)
        inCurrent = Fluke_Meter.ReadCurrent_Fluke(CURR_GPIB) ' Not tested very well
        
        Cells(X, y + 2).value = inCurrent
        
    Next cf
    
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, DT_REG, DT_Default)
    
    Dim SlewRateArray(1 To 9) As Integer
    SlewRateArray(1) = &H0
    SlewRateArray(2) = &H3
    SlewRateArray(3) = &HF
    SlewRateArray(4) = &H20
    SlewRateArray(5) = &HE0
    SlewRateArray(6) = &H23
    SlewRateArray(7) = &H2F
    SlewRateArray(8) = &HE3
    SlewRateArray(9) = &HEF
    
    y = OUTPUTDATACELL_Y + 5
    Cells(OUTPUTDATACELL_x - 2, y).value = "Quiescent Current VBAT"
    For SR_i = 0 To 8
        DoEvents
        X = OUTPUTDATACELL_x + SR_i
        Cells(X, y).value = "SR_i = 0x" & Hex(SlewRateArray(SR_i + 1))
        Cells(X, y + 1).value = SlewRateArray(SR_i + 1)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, SR_REG, SlewRateArray(SR_i + 1))
        
        Sleep (1000)
        inCurrent = Fluke_Meter.ReadCurrent_Fluke(CURR_GPIB) ' Not tested very well
       
        Cells(X, y + 2).value = inCurrent
    Next SR_i

End Sub
