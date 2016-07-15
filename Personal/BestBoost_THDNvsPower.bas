Attribute VB_Name = "BestBoost_THDNvsPower"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub THDNvsPower_AX80()
    'Measure the output noise across boost settings
    
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
    
    Cells(OUTPUTDATACELL_x - 2, OUTPUTDATACELL_Y).value = "Output Noise"
    Cells(OUTPUTDATACELL_x - 2, OUTPUTDATACELL_Y + 2).value = "Not Weighted"
    Cells(OUTPUTDATACELL_x - 2, OUTPUTDATACELL_Y + 3).value = "A-Weighted"
    
    AP.Sweep.Append = False
    
    For cf = 0 To 16
        DoEvents
        X = OUTPUTDATACELL_x + cf
        y = OUTPUTDATACELL_Y
        Cells(X, y).value = "cf = " & CStr(cf)
        Cells(X, y + 1).value = cf
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, DT_REG, cf)
        
        AP.Sweep.Start
        AP.Graph.Legend.comment(cf + 1, 1) = "0xcf = " & Hex(cf)
        AP.Sweep.Append = True
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
    Cells(OUTPUTDATACELL_x - 2, y).value = "Output Noise"
    Cells(OUTPUTDATACELL_x - 2, y + 2).value = "Not Weighted"
    Cells(OUTPUTDATACELL_x - 2, y + 3).value = "A-Weighted"
    For SR_i = 0 To 8
        DoEvents
        X = OUTPUTDATACELL_x + SR_i
        Cells(X, y).value = "SR_i = 0x" & Hex(SlewRateArray(SR_i + 1))
        Cells(X, y + 1).value = SlewRateArray(SR_i + 1)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, SR_REG, SlewRateArray(SR_i + 1))
        
        AP.Sweep.Start
        AP.Graph.Legend.comment(SR_i + 1 + 16, 1) = "0xd2 = " & Hex(SlewRateArray(SR_i + 1))
       
    Next SR_i

End Sub


