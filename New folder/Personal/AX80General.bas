Attribute VB_Name = "AX80General"
Sub bestBoostWrite(Optional DEV As Integer = &H74)
    
    'writes the best stable yet efficient boost settings
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H54)  ' T(est mode)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H4D) ' M
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCF, &H5) ' BST DT
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HC9, &H2)  ' DEM CLK/8
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCC, &H10) ' Force Skip
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HD2, &H3)  ' BST_SLEW
End Sub

Sub BW()
    Call bestBoostWrite
End Sub

Sub ATB1_VBATOK()
    DEV = &H74
    VBATOKAY_addrhi = &H0
    VBATOKAY_addrlo = &HA4
    VBATOKAY_en = &H80
    ATB1EN_addrhi = &H0
    ATB1EN_addrlo = &HAA
    ATB1EN_en = &H1
    
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H54)  ' T(est mode)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H4D) ' M
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, VBATOKAY_addrhi, VBATOKAY_addrlo, VBATOKAY_en)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ATB1EN_addrhi, ATB1EN_addrlo, ATB1EN_en)
    
End Sub
