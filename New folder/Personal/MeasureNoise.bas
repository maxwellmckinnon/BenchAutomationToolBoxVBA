Attribute VB_Name = "MeasureNoise"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub GrabMeasurement()
    ActiveCell.value = AP.Anlr.FuncRdg("V")
End Sub

Sub MeasureNoiseDEM_HZ09()
    ' Loop across SSM 660 / 330 (with correct 2/6 and 5/6), DEM off/8/4/2/1
    
    ' SETUP PARAMETERS
    THDN_FUNCTION_THRESHOLD = -55
    DataDumpX = 15 ' Modify this to change row where datadump starts
    DataDumpY = 39 '
    
    DEV = &H64
    EN_addrhi = &H0  '16 bit addressing, EN 0x00FF
    EN_addrlo = &HFF
    EN_bitOn = &H1
    EN_bitOff = &H0
    
    SSM_addrhi = &H0
    SSM_addrlo = &H3D
    SSM_660 = &H81 ' SSM 660 and 2/6
    SSM_330 = &H8C  ' SSM 330 and 5/6
    
    DEM_addrhi = &H0
    DEM_addrlo = &HC9
    DEM_off = &H0
    DEM_8 = &H2
    DEM_4 = &H6
    DEM_2 = &HA
    DEM_1 = &HD
    
    ' MEASUREMENT LOOP
    ' For each setting, verify that the path is working by measuring a THDN beyond the threshold.
    ' If the threshold is not met, flag the corresponding cell
    ' If the path is working, turn off the signal and measure the THDN Amplitude with 22Hz - 20KHz SPCL and A-Weighting
    
    X = DataDumpX
    y = DataDumpY
    
    Call HZ09.bestBoostWriteNoise   'Call the best test mode boost settings
    
    For SSM = 0 To 1
        DoEvents
        If SSM = 0 Then
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_660)
        ElseIf SSM = 1 Then
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_330)
        End If
        
        For DEM = 0 To 4
            If DEM = 0 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DEM_addrhi, DEM_addrlo, DEM_off)
            ElseIf DEM = 1 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DEM_addrhi, DEM_addrlo, DEM_8)
            ElseIf DEM = 2 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DEM_addrhi, DEM_addrlo, DEM_4)
            ElseIf DEM = 3 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DEM_addrhi, DEM_addrlo, DEM_2)
            ElseIf DEM = 4 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DEM_addrhi, DEM_addrlo, DEM_1)
            End If
                 
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOff) ' Reset part
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOn) ' Reset part

            AP.Anlr.FuncMode = 4 ' Switch to THDN
            AP.DGen.Output = True
            Sleep 1000 ' Give measurement time to settle
            Var = AP.Anlr.FuncRdg("dB") 'Grab THDN value in dB
            If Var < THDN_FUNCTION_THRESHOLD Then
                Cells(X, y + 3).value = "Good Signal Path"
            Else
                Cells(X, y + 3).value = "Bad Signal Path"
            End If
                
                AP.DGen.Output = False ' Turn off DGEN
                AP.Anlr.FuncMode = 3 ' Switch to THDN Amplitude
                AP.Anlr.FuncFilterLP = 5 ' Setup for Aweighting
                AP.Anlr.FuncFilter = 1
                Sleep 2000 ' Give measurement time to settle
                AW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                AP.Anlr.FuncFilterLP = 4
                Sleep 2000
                UW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                Cells(X, y).value = AW
                Cells(X, y + 1).value = UW
                
            X = X + 1
        
            DoEvents
        
        Next DEM
    Next SSM
    
End Sub

Sub MeasureNoise_HZ09()
    ' Loop across SSM on/off, DRE on/off, Gain 18/12/3dB
    ' Uses I2CBridge
    
    ' dump data to (3,17) AW, (3,18) UW
    ' Check if part is functioning correctly with THDN test
    ' SR is controlled at DIG_IF_SR = 0x23
    ' {48, 44.1, 11.025, 8} = {0x08, 0x07, 0x01, 0x00}
    ' SSM is controlled at SSM_ENA = 0x3d
    ' SR controlled by SPK_SR and DIG_IF_SR
    
    ' SETUP PARAMETERS
    THDN_FUNCTION_THRESHOLD = -55
    DataDumpX = 53 ' Modify this to change row where datadump starts
    DataDumpY = 24 ' column 'X' in 'raw' sheet
    
    DEV = &H64
    EN_addrhi = &H0  '16 bit addressing, EN 0x00FF
    EN_addrlo = &HFF
    EN_bitOn = &H1
    EN_bitOff = &H0
    
    SSM_addrhi = &H0
    SSM_addrlo = &H3D
    SSM_bitOn = &H81 ' SSM mode 1 on
    SSM_bitOff = &H1  ' SSM mode 1 off
    
    DRE_addrhi = &H0
    DRE_addrlo = &H39
    DRE_bitOn = &H1
    DRE_bitOff = &H0
    
    Gain_addrhi = &H0
    Gain_addrlo = &H3C
    Gain_18 = &H6
    Gain_12 = &H4
    Gain_3 = &H1
    
    'Three settings to change SR
    SRDIG_addrhi = &H0
    SRDIG_addrlo = &H24
    SRDIG_48 = &H8
    SRDIG_44 = &H7
    SRDIG_11 = &H1
    SRDIG_8 = &H0
    
    SRSPK_addrhi = &H0
    SRSPK_addrlo = &H25
    SRSPK_48 = &H80
    SRSPK_44 = &H70
    SRSPK_11 = &H10
    SRSPK_8 = &H0
    
    AP_MCLK_48 = 256
    AP_MCLK_44 = 256
    AP_MCLK_11 = 1536
    AP_MCLK_8 = 1536
    
    ' MEASUREMENT LOOP
    ' For each setting, verify that the path is working by measuring a THDN beyond the threshold.
    ' If the threshold is not met, flag the corresponding cell
    ' If the path is working, turn off the signal and measure the THDN Amplitude with 22Hz - 20KHz SPCL and A-Weighting
    
    X = DataDumpX
    y = DataDumpY
    
    Call HZ09.bestBoostWriteNoise   'Call the best test mode boost settings
    
    For SR = 0 To 3
        DoEvents
        If SR = 0 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_48)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_48)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 48000#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_48
            ElseIf SR = 1 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_44)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_44)
                
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 44100#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_44
                
            ElseIf SR = 2 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_11)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_11)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 11025#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_11
            ElseIf SR = 3 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_8)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_8)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 8000#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_8
            End If
        For gain = 0 To 2
            If gain = 0 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_18)
            ElseIf gain = 1 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_12)
            ElseIf gain = 2 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_3)
            End If
            
            For DRE = 0 To 1
                If DRE = 0 Then
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DRE_addrhi, DRE_addrlo, DRE_bitOn)
                ElseIf DRE = 1 Then
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DRE_addrhi, DRE_addrlo, DRE_bitOff)
                End If
                
                
                For SSM = 0 To 1
                    If SSM = 0 Then
                        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_bitOn)
                    ElseIf SSM = 1 Then
                        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_bitOff)
                    End If
                    
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOff) ' Reset part
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOn) ' Reset part
        
                    AP.Anlr.FuncMode = 4 ' Switch to THDN
                    AP.DGen.Output = True
                    Sleep 750 ' Give measurement time to settle
                    Var = AP.Anlr.FuncRdg("dB") 'Grab THDN value in dB
                    If Var < THDN_FUNCTION_THRESHOLD Then
                        Cells(X, y + 3).value = "Good Signal Path"
                    Else
                        Cells(X, y + 3).value = "Bad Signal Path"
                    End If
                        
                        AP.DGen.Output = False ' Turn off DGEN
                        AP.Anlr.FuncMode = 3 ' Switch to THDN Amplitude
                        Sleep 750 ' Give measurement time to settle
                        AP.Anlr.FuncFilterLP = 5 ' Setup for Aweighting
                        AP.Anlr.FuncFilter = 1
                        AW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                        Cells(X, y).value = AW
                        AP.Anlr.FuncFilterLP = 4
                        Sleep 750
                        UW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                        Cells(X, y + 1).value = UW
                    X = X + 1
                
                    DoEvents
                Next SSM
            Next DRE
        Next gain
    Next SR
    
End Sub

Sub MeasureNoise_AX80()
    ' Loop across SSM on/off, DRE on/off, Gain 18/12/3dB
    ' Uses I2CBridge
    
    ' dump data to (3,17) AW, (3,18) UW
    ' Check if part is functioning correctly with THDN test
    ' SR is controlled at DIG_IF_SR = 0x23
    ' {48, 44.1, 11.025, 8} = {0x08, 0x07, 0x01, 0x00}
    ' SSM is controlled at SSM_ENA = 0x3d
    ' SR controlled by SPK_SR and DIG_IF_SR
    
    ' SETUP PARAMETERS
    THDN_FUNCTION_THRESHOLD = -55
    DataDumpX = 149 ' Modify this to change row where datadump starts
    DataDumpY = 24 ' column X
    
    DEV = &H74
    EN_addrhi = &H0  '16 bit addressing, EN 0x00FF
    EN_addrlo = &HFF
    EN_bitOn = &H1
    EN_bitOff = &H0
    
    SSM_addrhi = &H0
    SSM_addrlo = &H3D
    SSM_bitOn = &H81 ' SSM mode 1 on
    SSM_bitOff = &H1  ' SSM mode 1 off
    
    DRE_addrhi = &H0
    DRE_addrlo = &H39
    DRE_bitOn = &H1
    DRE_bitOff = &H0
    
    Gain_addrhi = &H0
    Gain_addrlo = &H3C
    Gain_18 = &H6
    Gain_12 = &H4
    Gain_3 = &H1
    
    'Three settings to change SR
    SRDIG_addrhi = &H0
    SRDIG_addrlo = &H23
    SRDIG_48 = &H8
    SRDIG_44 = &H7
    SRDIG_11 = &H1
    SRDIG_8 = &H0
    
    SRSPK_addrhi = &H0
    SRSPK_addrlo = &H24
    SRSPK_48 = &H80
    SRSPK_44 = &H70
    SRSPK_11 = &H10
    SRSPK_8 = &H0
    
    AP_MCLK_48 = 256
    AP_MCLK_44 = 256
    AP_MCLK_11 = 1536
    AP_MCLK_8 = 1536
    
    ' MEASUREMENT LOOP
    ' For each setting, verify that the path is working by measuring a THDN beyond the threshold.
    ' If the threshold is not met, flag the corresponding cell
    ' If the path is working, turn off the signal and measure the THDN Amplitude with 22Hz - 20KHz SPCL and A-Weighting
    
    X = DataDumpX
    y = DataDumpY
    
    Call bestBoostWrite   'Call the best test mode boost settings
    
    For SR = 0 To 3
        DoEvents
        If SR = 0 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_48)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_48)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 48000#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_48
            ElseIf SR = 1 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_44)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_44)
                
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 44100#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_44
                
            ElseIf SR = 2 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_11)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_11)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 11025#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_11
            ElseIf SR = 3 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_8)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_8)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 8000#
                AP.PSIA.Tx.NFsClk.Factor = AP_MCLK_8
            End If
        For gain = 0 To 2
            If gain = 0 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_18)
            ElseIf gain = 1 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_12)
            ElseIf gain = 2 Then
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_3)
            End If
            
            For DRE = 0 To 1
                If DRE = 0 Then
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DRE_addrhi, DRE_addrlo, DRE_bitOn)
                ElseIf DRE = 1 Then
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, DRE_addrhi, DRE_addrlo, DRE_bitOff)
                End If
                
                
                For SSM = 0 To 1
                    If SSM = 0 Then
                        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_bitOn)
                    ElseIf SSM = 1 Then
                        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_bitOff)
                    End If
                    
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOff) ' Reset part
                    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOn) ' Reset part
        
                    AP.Anlr.FuncMode = 4 ' Switch to THDN
                    AP.DGen.Output = True
                    Sleep 750 ' Give measurement time to settle
                    Var = AP.Anlr.FuncRdg("dB") 'Grab THDN value in dB
                    If Var < THDN_FUNCTION_THRESHOLD Then
                        Cells(X, y + 3).value = "Good Signal Path"
                    Else
                        Cells(X, y + 3).value = "Bad Signal Path"
                    End If
                        
                        AP.DGen.Output = False ' Turn off DGEN
                        AP.Anlr.FuncMode = 3 ' Switch to THDN Amplitude
                        Sleep 750 ' Give measurement time to settle
                        AP.Anlr.FuncFilterLP = 5 ' Setup for Aweighting
                        AP.Anlr.FuncFilter = 1
                        AW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                        Cells(X, y).value = AW
                        AP.Anlr.FuncFilterLP = 4
                        Sleep 750
                        UW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                        Cells(X, y + 1).value = UW
                    X = X + 1
                
                    DoEvents
                Next SSM
            Next DRE
        Next gain
    Next SR
    
End Sub

Sub MeasureNoise_AX90()
    ' Loop across SSM on/off, DRE on/off, Gain 18/12/3dB
    
    ' dump data to (3,17) AW, (3,18) UW
    ' Check if part is functioning correctly with THDN test
    ' SR is controlled at DIG_IF_SR = 0x23
    ' {48, 44.1, 11.025, 8} = {0x08, 0x07, 0x01, 0x00}
    ' SSM is controlled at SSM_ENA = 0x3d
    ' SR controlled by SPK_SR and DIG_IF_SR
    
    ' SETUP PARAMETERS
    THDN_FUNCTION_THRESHOLD = -55
    DataDumpX = 53
    DataDumpY = 24 ' X
    
    DEV = &H74
    EN_addrhi = &H0  '16 bit addressing, EN 0x00FF
    EN_addrlo = &HFF
    EN_bitOn = &H1
    EN_bitOff = &H0
    
    SSM_addrhi = &H0
    SSM_addrlo = &H3D
    SSM_bitOn = &H81 ' SSM mode 1 on
    SSM_bitOff = &H1  ' SSM mode 1 off
    
    DRE_addrhi = &H0
    DRE_addrlo = &H39
    DRE_bitOn = &H1
    DRE_bitOff = &H0
    
    Gain_addrhi = &H0
    Gain_addrlo = &H3C
    Gain_18 = &H5
    Gain_12 = &H4
    Gain_3 = &H1
    
    'Two settings to change SR
    SRDIG_addrhi = &H0
    SRDIG_addrlo = &H23
    SRDIG_48 = &H8
    SRDIG_44 = &H7
    SRDIG_11 = &H1
    SRDIG_8 = &H0
    
    SRSPK_addrhi = &H0
    SRSPK_addrlo = &H24
    SRSPK_48 = &H80
    SRSPK_44 = &H70
    SRSPK_11 = &H10
    SRSPK_8 = &H0
    
    ' MEASUREMENT LOOP
    ' For each setting, verify that the path is working by measuring a THDN beyond the threshold.
    ' If the threshold is not met, flag the corresponding cell
    ' If the path is working, turn off the signal and measure the THDN Amplitude with 22Hz - 20KHz SPCL and A-Weighting
    
    X = DataDumpX
    y = DataDumpY
    
    For SR = 0 To 3
        DoEvents
        If SR = 0 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_48)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_48)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 48000#
            ElseIf SR = 1 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_44)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_44)
                
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 44100#
                
            ElseIf SR = 2 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_11)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_11)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 11025#
            ElseIf SR = 3 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRDIG_addrhi, SRDIG_addrlo, SRDIG_8)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SRSPK_addrhi, SRSPK_addrlo, SRSPK_8)
                AP.PSIA.Tx.FrameClk.Rate("Hz") = 8000#
            End If
        For gain = 0 To 2
            If gain = 0 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_18)
            ElseIf gain = 1 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_12)
            ElseIf gain = 2 Then
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, Gain_addrhi, Gain_addrlo, Gain_3)
            End If
            
            For DRE = 0 To 1
                If DRE = 0 Then
                    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, DRE_addrhi, DRE_addrlo, DRE_bitOn)
                ElseIf DRE = 1 Then
                    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, DRE_addrhi, DRE_addrlo, DRE_bitOff)
                End If
                
                
                For SSM = 0 To 1
                    If SSM = 0 Then
                        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_bitOn)
                    ElseIf SSM = 1 Then
                        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, SSM_addrhi, SSM_addrlo, SSM_bitOff)
                    End If
                    
                    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOff) ' Reset part
                    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, EN_addrhi, EN_addrlo, EN_bitOn) ' Reset part
        
                    AP.Anlr.FuncMode = 4 ' Switch to THDN
                    AP.DGen.Output = True
                    Sleep 750 ' Give measurement time to settle
                    Var = AP.Anlr.FuncRdg("dB") 'Grab THDN value in dB
                    If Var < THDN_FUNCTION_THRESHOLD Then
                        Cells(X, y).value = "Good Signal Path"
                        AP.DGen.Output = False ' Turn off DGEN
                        AP.Anlr.FuncMode = 3 ' Switch to THDN Amplitude
                        Sleep 750 ' Give measurement time to settle
                        AP.Anlr.FuncFilterLP = 5 ' Setup for Aweighting
                        AP.Anlr.FuncFilter = 1
                        AW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                        Cells(X, y).value = AW
                        AP.Anlr.FuncFilterLP = 4
                        Sleep 750
                        UW = AP.Anlr.FuncRdg("V") ' Grab value in V (uV)
                        Cells(X, y + 1).value = UW
                    Else
                        Cells(X, y).value = "Bad Signal Path"
                    End If
                    X = X + 1
                
                Next SSM
            Next DRE
        Next gain
    Next SR
    
I2C_Controls_.I2C_Disconnect
    
End Sub

Sub CheckTxChannels()
    'Check all TX channels enables across the DSP channel destinations
    DEV = &H64
    ENhi = &H1 '16 bit addressing, EN 0x0100
    ENlo = &H0
    
    I2C_Controls_.I2C_Connect
    Sleep 500
    
    Cells(2, 7).value = "AMP DSP Destination Source 0 Enabled"
    Cells(2, 8).value = "AMP DSP Destination Source 0 Disabled"
    
    'Tx testing (Device is Tx, AP is Rx)
    AP.PSIA.Tx.Data.ChannelB = 1 'Default to channel one, prepare to sweep channel B
    AP.PSIA.Rx.Data.ChannelB = 1 'Default to channel one, prepare to sweep channel B
    For chan = 0 To 15 'Changes the TX channels
        If chan >= 1 Then 'Default is set to send to ch0 and 1
            AP.PSIA.Rx.Data.ChannelB = chan ' Setup AP to send data to correct chan
        End If
        
        'Setup DSP output to channel chan
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1F, chan)
        
        Sleep (1000)
        
        'Value when disabled
        If chan < 1 Then 'Use analyzer A for chn 0, B for everything above 1
            Var = AP.S2CDsp.Analyzer.FuncChARdg("dB")
        ElseIf chan >= 1 Then
            Var = AP.S2CDsp.Analyzer.FuncChBRdg("dB")
        End If
        Cells(chan + 3, 2 * ampChan + 7 + 1).value = Var ' Value when disabled
        
        'Enable specific channel receive - reg 0x18 and 0x19
        If chan < 8 Then
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1A, 2 ^ chan)
        ElseIf chan >= 8 Then
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1B, 2 ^ (chan - 8))
        End If
        
        Sleep (1000) ' Give THDN time to set
        
        Cells(chan + 3, 2).value = chan
        If chan < 1 Then 'Use analyzer A for chn 0, B for everything above 1
            Var = AP.S2CDsp.Analyzer.FuncChARdg("dB")
        ElseIf chan >= 1 Then
            Var = AP.S2CDsp.Analyzer.FuncChBRdg("dB")
        End If
        Cells(chan + 3, 2 * ampChan + 7).value = Var 'offset the data by 3 to make room for info
        
    Next chan
    
    
    I2C_Controls_.I2C_Disconnect
    
End Sub
