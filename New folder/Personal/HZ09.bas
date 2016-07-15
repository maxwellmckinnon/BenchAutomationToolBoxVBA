Attribute VB_Name = "HZ09"
Sub bestBoostWrite(Optional DEV As Integer = &H64)

    'writes the best stable yet efficient boost settings
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H54)  ' T(est mode)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H4D) ' M
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HC9, &H6)  ' DEM/4
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCB, &H0)  '
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCC, &H10)  '
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCF, &H4)   '
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HD2, &H0)   '
End Sub

Sub bestBoostWriteNoise(Optional DEV As Integer = &H64)

    'writes the best stable yet efficient boost settings
    'Turns off DEM for best noise performance
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H54)  ' T(est mode)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H4D) ' M
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HC9, &H0)  ' DEM Off
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCB, &H0)  '
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCC, &H10)  '
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HCF, &H0)  '
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HD2, &HCF)  '
End Sub

Sub BW()
    Call bestBoostWrite
End Sub

