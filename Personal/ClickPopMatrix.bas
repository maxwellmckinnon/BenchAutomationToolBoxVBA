Attribute VB_Name = "ClickPopMatrix"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub AX80_A1_ClickDEM()
'
' Measure the SW on and off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

' Check DEM settings too

    TRIALS = 5
    DEV = &H74 ' 0x74 is ax80
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &HFF
    EN = ENhi * 256 + ENlo
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    I2CBridgeDemos.I2CBridgeConnected
    
    'Call I2C.I2CWriteByte16bit(DEVADDRI2C, RevIDReg, &H54)
    
    Sleep (500)
    Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off SW_EN

    'DEM off
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H0)
    Sleep (500)
    
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    
    'DEM clk/8
    ActiveCell.Offset(2, -1 * TRIALS).Select
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H3)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    
    'DEM clk/4
    ActiveCell.Offset(2, -1 * TRIALS).Select
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H7)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i

End Sub

Sub testClickAcrossVOSSettings()
' Checks click and pop over different voltage offset settings
' Measure the SW on and off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    TRIALS = 20
    DEV = &H74 ' 0x74 is ax80
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &HFF
    EN = ENhi * 256 + ENlo
    VOSreghi = &H0
    VOSreglo = &HCA
    VOSreg = VOSreghi * 256 + VOSreglo
    VOSstartlowend = &H10
    VOSfinishhighend = &H20
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    I2CBridgeDemos.I2CBridgeConnected
    
    'Call I2C.I2CWriteByte16bit(DEVADDRI2C, RevIDReg, &H54)
    
    Sleep (500)
    Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off SW_EN
    Call AX80BestBoost.bestBoostWrite
    
    For j = VOSstartlowend To VOSfinishhighend
        Call i2c.I2CWriteByte16bit(DEV, VOSreg, j)
        Sleep (1000)
        ActiveCell.value = "VOS trim 0xc9 = 0x" & Hex(j)
        ActiveCell.Offset(0, 1).Select
        
        For i = 1 To TRIALS
            DoEvents
            AP.BarGraph.Reset (1)
            Sleep (2000)
            Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
            Sleep (1000)
            Result = AP.BarGraph.max(1)
            ActiveCell.value = Result
            AP.BarGraph.Reset (1)
            
            ActiveCell.Offset(1, 0).Select
            Sleep (2000)
            Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
            Sleep (1000)
            Result = AP.BarGraph.max(1)
            ActiveCell.value = Result
            AP.BarGraph.Reset (1)
            
            ActiveCell.Offset(-1, 1).Select
            
        Next i
        ActiveCell.Offset(2, -1 * TRIALS - 1).Select
    Next j
    
End Sub

Sub CLICK_SW_EN_HZ09()
'
' Uses Bill's I2C Bridge HW
' Measure the SW on and off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

'Check DEM settings too

    TRIALS = 10
    DEV = &H64 ' 0x74 is ax80
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &HFF
    EN = ENhi * 256 + ENlo
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    I2CBridgeDemos.I2CBridgeConnected
    
    Call HZ09.bestBoostWrite
    
    Sleep (500)
    Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off SW_EN

    'DEM clk/4
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H6)
    Sleep (2000)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    
End Sub

Sub CLICK_SW_EN_BILL_I2C()
'
' Measure the SW on and off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

'Check DEM settings too

    TRIALS = 200
    DEV = &H74 ' 0x74 is ax80
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &HFF
    EN = ENhi * 256 + ENlo
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    I2CBridgeDemos.I2CBridgeConnected
    
    'Call I2C.I2CWriteByte16bit(DEVADDRI2C, RevIDReg, &H54)
    
    Sleep (500)
    Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off SW_EN

    
    
    'DEM off
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H0)
    Sleep (500)
    
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    
    'DEM clk/8
    ActiveCell.Offset(2, -1 * TRIALS).Select
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H2)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    'DEM clk/4
    ActiveCell.Offset(2, -1 * TRIALS).Select
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &H6)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    
    'DEM clk/2
    ActiveCell.Offset(2, -1 * TRIALS).Select
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &HA)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    
    'DEM clk/1
    ActiveCell.Offset(2, -1 * TRIALS).Select
    Call i2c.I2CWriteByte16bit(DEV, &HC9, &HE)
    For i = 1 To TRIALS
        DoEvents
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 1)  ' turn on SW_EN
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call i2c.I2CWriteByte16bit(DEV, EN, 0)  ' turn off
        Sleep (1000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
        
    Next i
    Sleep (500)
    
    
End Sub

Sub CLICK_SW_EN()
'
' Measure the SW on and off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

    TRIALS = 10
    DEV = &H62
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &H80

    I2C_Controls_.I2C_Connect
    Sleep (500)
    
    For a = 0 To 1 'amp1, 2
        DoEvents
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (500)
        'Check for correct path
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 1)
        AP.DGen.Output = True
        Sleep (500)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var > -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells(1, 1).value = "Fail basic signal path test"
            MsgBox "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 0)
        Sleep (500)
        
        For i = 1 To TRIALS
            Sleep (500)
            AP.BarGraph.Reset (3)
            Sleep (500)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 1)  ' turn on SW_EN
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(1, 0).Select
            Sleep (500)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 0)  ' turn off
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(-1, 1).Select
            
        Next i
    Next a
    Sleep (500)
    
    
    I2C_Controls_.I2C_Disconnect
    
End Sub

Sub WriteDeviceBasicClickPopSetup_AX80()
    'Write the registers to the device and check the path
    DEV = &H74
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5, &HC0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H8, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H9, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HA, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HB, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HC, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HD, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HE, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HF, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H10, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H11, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H12, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H13, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H14, &H78)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H15, &HFF)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H16, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H17, &H55)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H18, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H19, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H20, &HC0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H21, &H1C)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H22, &H44)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H23, &H8)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H24, &H88)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H25, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H26, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H27, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H28, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H30, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H31, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H32, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H33, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H34, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H35, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H36, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H37, &H2)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H38, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H39, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3A, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3C, &H4)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3D, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3F, &H4)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H40, &H1C)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H41, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H42, &H3F)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H43, &H4)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H44, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H45, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H46, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H47, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H48, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H49, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4E, &H29)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H50, &H15)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H51, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H52, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H53, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H54, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H55, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H56, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H57, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H58, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H59, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H60, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H61, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H62, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H63, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H64, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H65, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H66, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H67, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H68, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H69, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H70, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H71, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H72, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H73, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H74, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H75, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H76, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H77, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H78, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H79, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H80, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H81, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H82, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H83, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H84, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H85, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H86, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H87, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H89, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HFF, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &H0, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H40)

    
End Sub

Sub WriteDeviceBasicClickPopSetupMasterMode_AX80()
    'Write the registers to the device and check the path
    DEV = &H74
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5, &HC0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H8, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H9, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HA, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HB, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HC, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HD, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HE, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HF, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H10, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H11, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H12, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H13, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H14, &H78)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H15, &HFF)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H16, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H17, &H55)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H18, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H19, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H1F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H20, &HC0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H21, &H1F)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H22, &H44)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H23, &H8)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H24, &H88)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H25, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H26, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H27, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H28, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H2F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H30, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H31, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H32, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H33, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H34, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H35, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H36, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H37, &H2)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H38, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H39, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3A, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3C, &H4)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3D, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H3F, &H4)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H40, &H1C)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H41, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H42, &H3F)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H43, &H4)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H44, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H45, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H46, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H47, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H48, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H49, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H4E, &H29)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H50, &H15)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H51, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H52, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H53, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H54, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H55, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H56, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H57, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H58, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H59, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H5F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H60, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H61, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H62, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H63, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H64, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H65, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H66, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H67, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H68, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H69, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H6F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H70, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H71, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H72, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H73, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H74, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H75, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H76, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H77, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H78, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H79, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7A, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7B, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7C, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7D, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7E, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H7F, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H80, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H81, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H82, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H83, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H84, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H85, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H86, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H87, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &H89, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HFF, &H1)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &H0, &H0)
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H1, &HFF, &H40)

    
End Sub

Sub WriteDeviceBasicClickPopSetup()
    'Write the registers to the device and check the path
    DEV = &H62
    
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H10, &H20)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H11, &H2)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &H30)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H13, &H5)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H14, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H15, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H16, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H17, &HC0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1A, &H2)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1C, &H4)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H1E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H20, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H21, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H22, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H23, &H40)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H24, &H15)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H25, &H3)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H26, &H8)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H30, &H3)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H31, &HFF)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H32, &HFF)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H33, &HC0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H34, &H15)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H35, &H5)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H36, &H8)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H40, &H1)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H41, &H4)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H42, &H8)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H43, &HC0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H44, &H8)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H45, &H10)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H46, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H47, &H1)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H48, &H4)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H49, &H8)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H4A, &HC0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H4B, &H8)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H4C, &H11)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H4D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H50, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H51, &H4)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H52, &H7)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H53, &H10)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H54, &H11)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H60, &H1)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H61, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H62, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H63, &H2)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H70, &H1)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H81, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H90, &H1)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H91, &H5E)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H92, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H93, &H7)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H94, &H4)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H95, &HF)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H96, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H97, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H98, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H99, &HC)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H9A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H9B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H9C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H9D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H9E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H9F, &H60)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA0, &H54)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA7, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA8, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HA9, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HAA, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HAB, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HAC, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HAD, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HAE, &H80)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HAF, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB0, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB7, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB8, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HB9, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HBA, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HBB, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HBC, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HBD, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HBE, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HBF, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC0, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC7, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC8, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HC9, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HCA, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HCB, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HCC, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HCD, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HCE, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HCF, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD0, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD7, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD8, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HD9, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HDA, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE0, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE7, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE8, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HE9, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HEA, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HEB, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HEC, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HED, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HEE, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HEF, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF0, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &HF6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H0, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H7, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H8, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H9, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &HA, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &HB, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &HC, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &HD, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &HE, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &HF, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H10, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H11, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H12, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H13, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H14, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H15, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H16, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H17, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H18, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H19, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H1F, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H20, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H21, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H22, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H23, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H24, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H25, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H26, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H27, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H28, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H29, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H2F, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H30, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H31, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H32, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H33, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H34, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H35, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H36, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H37, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H38, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H39, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H3F, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H40, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H41, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H42, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H43, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H44, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H45, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H46, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H47, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H48, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H49, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H4F, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H50, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H51, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H52, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H53, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H54, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H55, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H56, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H57, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H58, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H59, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H5F, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H60, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H61, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H62, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H63, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H64, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H65, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H66, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H67, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H68, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H69, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6A, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6B, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6C, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6D, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6E, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H6F, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H70, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H71, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H72, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H73, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H74, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H75, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H76, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H1, &H77, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H7F, &HFE, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H7F, &HFF, &H40)

End Sub

Sub CLICK_HW_EN_AX80()
' Measure the HW off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

    TRIALS = 10
    DEV = &H74
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &H80
    Dim GPIB As String: GPIB = "GPIB::01" '
    Dim GPIBOutputName As String: GPIBOutputName = "P25V" 'Connect HW EN pin to 25V side of PSU
    
    Sleep (500)
    
    For a = 0 To 0 'amp1, 2
        DoEvents
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (500)
        'Check for correct path
        Call WriteDeviceBasicClickPopSetup_AX80
        Call AX80BestBoost.bestBoostWrite
        
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, 1)
        AP.DGen.Output = True
        Sleep (500)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var < -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells(1, 1).value = "Fail basic signal path test"
            MsgBox "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, 0)
        Sleep (500)
        
        For i = 1 To TRIALS
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
            Sleep (500)
            Call WriteDeviceBasicClickPopSetup_AX80
            Call AX80BestBoost.bestBoostWrite
            Sleep (500)
            Sleep (500)
            AP.BarGraph.Reset (1)
            Sleep (500)
            
            'Turn HW off
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
            Sleep (500)
            Result = AP.BarGraph.max(1)
            ActiveCell.value = Result
            AP.BarGraph.Reset (1)
            
            ActiveCell.Offset(0, 1).Select
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
        Next i
    Next a
    Sleep (500)
    
    
End Sub

Sub CLICK_VBATRemoval_AX80()
' Measure the HW off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

    TRIALS = 10
    DEV = &H74
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &H80
    Dim GPIB As String: GPIB = "GPIB::01" '
    Dim GPIBOutputName As String: GPIBOutputName = "P6V" 'Connect HW EN pin to 25V side of PSU
    
    Sleep (500)
    
    For a = 0 To 0 'amp1, 2
        DoEvents
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (500)
        'Check for correct path
        Call WriteDeviceBasicClickPopSetup_AX80
        Call AX80BestBoost.bestBoostWrite
        
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, 1)
        AP.DGen.Output = True
        Sleep (500)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var < -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells(1, 1).value = "Fail basic signal path test"
            MsgBox "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, 0)
        Sleep (500)
        
        For i = 1 To TRIALS
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 3.6)
            Sleep (500)
            Call WriteDeviceBasicClickPopSetup_AX80
            Call AX80BestBoost.bestBoostWrite
            Sleep (500)
            Sleep (500)
            AP.BarGraph.Reset (1)
            Sleep (500)
            
            'Turn HW off
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
            Sleep (500)
            Result = AP.BarGraph.max(1)
            ActiveCell.value = Result
            AP.BarGraph.Reset (1)
            
            ActiveCell.Offset(0, 1).Select
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 3.6)
        Next i
    Next a
    Sleep (500)
    
    
End Sub

Sub CLICK_HW_EN()
' Measure the HW off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

    TRIALS = 10
    DEV = &H62
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &H80
    Dim GPIB As String: GPIB = "GPIB::06" 'Connect HW EN pin to 6V side of PSU
    Dim GPIBOutputName As String: GPIBOutputName = "P6V"
    
    I2C_Controls_.I2C_Connect
    Sleep (500)
    
    For a = 0 To 1 'amp1, 2
        DoEvents
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (500)
        'Check for correct path
        Call WriteDeviceBasicClickPopSetup
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 1)
        AP.DGen.Output = True
        Sleep (500)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var > -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells(1, 1).value = "Fail basic signal path test"
            MsgBox "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 0)
        Sleep (500)
        
        For i = 1 To TRIALS
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
            Sleep (500)
            Call WriteDeviceBasicClickPopSetup
            Sleep (500)
            Sleep (500)
            AP.BarGraph.Reset (3)
            Sleep (500)
            
            'Turn HW off
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(0, 1).Select
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
        Next i
    Next a
    Sleep (500)
    
    
    I2C_Controls_.I2C_Disconnect
    
End Sub

Sub Click_AVDD_Removal()

' Measure the AVDD on and off click/pop - write device to correct path before removing AVDD,
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

    TRIALS = 10
    DEV = &H62
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &H80
    Dim GPIB As String: GPIB = "GPIB::06" 'Connect AVDD pin to 6V side of PSU
    Dim GPIBOutputName As String: GPIBOutputName = "P6V"
    
    I2C_Controls_.I2C_Connect
    Sleep (500)
    
    For a = 0 To 1 'amp1, 2
        DoEvents
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
        Sleep (500)
        'Check for correct path
        Call WriteDeviceBasicClickPopSetup
        Sleep (500)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 0)
        Sleep (500)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 1)
        AP.DGen.Output = True
        Sleep (500)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var > -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells(1, 1).value = "Fail basic signal path test"
            MsgBox "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
        
        Sleep (500)
        
        For i = 1 To TRIALS
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
            Sleep (500)
            AP.BarGraph.Reset (3)
            Sleep (500)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(1, 0).Select
            Call WriteDeviceBasicClickPopSetup
            Sleep (500)
            AP.BarGraph.Reset (3)
            Sleep (500)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)  ' turn off
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(-1, 1).Select
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
'
'            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
'            Sleep (500)
'            Call WriteDeviceBasicClickPopSetup
'            Sleep (500)
'            Sleep (500)
'            AP.BarGraph.Reset (3)
'            Sleep (500)
'
'            'Turn HW off
'            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
'            Sleep (500)
'            Result = AP.BarGraph.Max(3)
'            ActiveCell.value = Result
'            AP.BarGraph.Reset (3)
'
'            ActiveCell.Offset(0, 1).Select
'            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
        Next i
    Next a
    Sleep (500)
    
    
    I2C_Controls_.I2C_Disconnect
    
End Sub

Sub Click_DVDD()
' Measure the DVDD off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max

    TRIALS = 10
    DEV = &H62
    '16 bit addressing, SW_EN 0x0080
    ENhi = &H0
    ENlo = &H80
    Dim GPIB As String: GPIB = "GPIB::03" 'Connect HW EN pin to 25V side of PSU
    Dim GPIBOutputName As String: GPIBOutputName = "P25V"
    
    I2C_Controls_.I2C_Connect
    Sleep (500)
    
    For a = 0 To 1 'amp1, 2
        DoEvents
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (500)
        'Check for correct path
        Call WriteDeviceBasicClickPopSetup
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 1)
        AP.DGen.Output = True
        Sleep (500)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var > -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells(1, 1).value = "Fail basic signal path test"
            MsgBox "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, ENhi, ENlo, 0)
        Sleep (500)
        
        For i = 1 To TRIALS
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
            Sleep (500)
            Call WriteDeviceBasicClickPopSetup
            Sleep (500)
            Sleep (500)
            AP.BarGraph.Reset (3)
            Sleep (500)
            
            'Turn HW off
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(0, 1).Select
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 1.8)
        Next i
    Next a
    Sleep (500)
    
    
    I2C_Controls_.I2C_Disconnect
    
End Sub

Sub CLICK_VBAT()
' Measure the VBAT on and off click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    GPIB = "GPIB::03"
    GPIBOutputName = "P6V"
    TRIALS = 10
    DEV = &H62
   
    'I2C_Controls_.I2C_Connect
    Sleep (200)
    
    For a = 0 To 1 'amp1, 2
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (500)
        'Check for correct path
        Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 3.6)
        Call WriteDeviceBasicClickPopSetup
        AP.DGen.Output = True
        Sleep (1000)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var > -50 Then
            
        Else
            'Did not pass correct signal path test
            Cells.Range(1, 1).value = "Fail basic signal path test"
        End If
            
        AP.DGen.Output = False
        Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)
        Sleep (500)
        
        For i = 1 To TRIALS
            Sleep (500)
            AP.BarGraph.Reset (3)
            Sleep (500)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 3.6) ' turn on VBAT
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(1, 0).Select
            Sleep (500)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB, GPIBOutputName, 0)  ' turn off
            Sleep (500)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(-1, 1).Select
            
        Next i
    Next a
    Sleep (500)
    
End Sub

Sub ClearFlags_AX80()
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HD, &HFF)  ' Clear flags
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HE, &HFF) ' Clear flags
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, &H0, &HF, &HFF) ' Clear flags
End Sub

Sub CLICK_BCLKRemoval_AX80()
' Measure the BCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H74
    CMONhi = &H0
    CMONlo = &H11
    CMONAUTO = &H3
    temp = ActiveCell
    ENhi = &H0
    ENlo = &HFF
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
    
    Call WriteDeviceBasicClickPopSetup_AX80
    Call AX80BestBoost.bestBoostWrite
    
    Sleep (200)
    For c = 2 To 0 Step -1 ' Cmon enabled, disabled
        If c = 2 Then
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, CMONhi, CMONlo, CMONAUTO)
        Else
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        End If
        
        For a = 0 To 0 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (500)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
            Call ClearFlags_AX80
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
            Sleep (1000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -50 Then
                
            Else
                'Did not pass correct signal path test
                Cells(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.Tx.BitClk.Dir = 1
            AP.DGen.Output = False
            Sleep (500)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
                Sleep (300)
                Call ClearFlags_AX80 ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2500)
                AP.BarGraph.Reset (1)
            
                Sleep (500)
                AP.BarGraph.Reset (1)
                Sleep (2000)
                AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
                Sleep (2000)
                Result = AP.BarGraph.max(1)
                ActiveCell.value = Result
                Sleep (300)
                AP.BarGraph.Reset (1)
                
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
                Sleep (300)
                Call ClearFlags_AX80 ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (1)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.Tx.BitClk.Dir = 1  ' turn off BCLK
                Sleep (2000)
                Result = AP.BarGraph.max(1)
                ActiveCell.value = Result
                AP.BarGraph.Reset (1)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = temp
        ActiveCell.Offset(2, 0).Select
    Next c
    Sleep (500)
    AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
    
End Sub

Sub CLICK_LRCLKRemoval_AX80()
' Measure the BCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H74
    CMONhi = &H0
    CMONlo = &H11
    temp = ActiveCell
    ENhi = &H0
    ENlo = &HFF
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
    
    Call WriteDeviceBasicClickPopSetup_AX80
    Call AX80BestBoost.bestBoostWrite
    
    Sleep (200)
    For c = 1 To 0 Step -1 ' Cmon enabled, disabled
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        
        For a = 0 To 0 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (500)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.Tx.FrameClk.Dir = 0 ' turn on LRCLK
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
            Call ClearFlags_AX80
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
            Sleep (1000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -50 Then
                
            Else
                'Did not pass correct signal path test
                Cells(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.Tx.FrameClk.Dir = 1
            AP.DGen.Output = False
            Sleep (500)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
                Sleep (300)
                Call ClearFlags_AX80 ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2500)
                AP.BarGraph.Reset (1)
            
                Sleep (500)
                AP.BarGraph.Reset (1)
                Sleep (2000)
                AP.PSIA.Tx.FrameClk.Dir = 0 ' turn on LRCLK
                Sleep (2000)
                Result = AP.BarGraph.max(1)
                ActiveCell.value = Result
                Sleep (300)
                AP.BarGraph.Reset (1)
                
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
                Sleep (300)
                Call ClearFlags_AX80 ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (1)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.Tx.FrameClk.Dir = 1  ' turn off LRCLK
                Sleep (2000)
                Result = AP.BarGraph.max(1)
                ActiveCell.value = Result
                AP.BarGraph.Reset (1)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = temp
        ActiveCell.Offset(2, 0).Select
    Next c
    Sleep (500)
    AP.PSIA.Tx.FrameClk.Dir = 0 ' turn on LRCLK
    
End Sub

Sub CLICK_MCLKRemoval_AX80()
' Measure the BCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H74
    CMONhi = &H0
    CMONlo = &H11
    CMONAUTO = &H3
    temp = ActiveCell
    ENhi = &H0
    ENlo = &HFF
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
    
    Call WriteDeviceBasicClickPopSetupMasterMode_AX80
    Call AX80BestBoost.bestBoostWrite
    
    Sleep (200)
    For c = 2 To 0 Step -1 ' Cmon enabled, disabled
        If c = 2 Then
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, CMONhi, CMONlo, CMONAUTO)
        Else
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        End If
        
        For a = 0 To 0 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (500)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.MasterClkDir = 1 ' turn on MCLK
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
            Call ClearFlags_AX80
            Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
            Sleep (1000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -50 Then
                
            Else
                'Did not pass correct signal path test
                Cells(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.MasterClkDir = 0
            AP.DGen.Output = False
            Sleep (500)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
                Sleep (300)
                Call ClearFlags_AX80 ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2500)
                AP.BarGraph.Reset (1)
            
                Sleep (500)
                AP.BarGraph.Reset (1)
                Sleep (2000)
                AP.PSIA.MasterClkDir = 1 ' turn on MCLK
                Sleep (2000)
                Result = AP.BarGraph.max(1)
                ActiveCell.value = Result
                Sleep (300)
                AP.BarGraph.Reset (1)
                
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
                Sleep (300)
                Call ClearFlags_AX80 ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (1)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.MasterClkDir = 0 ' turn off LRCLK
                Sleep (2000)
                Result = AP.BarGraph.max(1)
                ActiveCell.value = Result
                AP.BarGraph.Reset (1)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = temp
        ActiveCell.Offset(2, 0).Select
    Next c
    Sleep (500)
    AP.PSIA.MasterClkDir = 1 ' turn on MCLK
    
End Sub

Sub CLICK_BCLKRemoval()
' Measure the BCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    For c = 1 To 0 Step -1 ' Cmon enabled, disabled
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        
        For a = 0 To 1 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (500)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
            Sleep (1000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -50 Then
                
            Else
                'Did not pass correct signal path test
                Cells(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.Tx.BitClk.Dir = 1
            AP.DGen.Output = False
            Sleep (500)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (300)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2500)
                AP.BarGraph.Reset (3)
            
                Sleep (500)
                AP.BarGraph.Reset (3)
                Sleep (2000)
                AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                Sleep (300)
                AP.BarGraph.Reset (3)
                
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (300)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (300)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (300)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.Tx.BitClk.Dir = 1  ' turn off BCLK
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = temp
        ActiveCell.Offset(2, 0).Select
    Next c
    Sleep (500)
    AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
    
End Sub

Sub ClkRemove()
' BCLK, LRCLK, and MCLK on and off

    temp = ActiveCell
    Call CLICK_BCLKRemoval
    

End Sub

Sub CLICK_BCLKRemoval_temp()
' Measure the BCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    c = 0
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        
        For a = 0 To 1 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (2000)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
            Sleep (2000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -50 Then
                
            Else
                'Did not pass correct signal path test
                Cells.Range(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.Tx.BitClk.Dir = 1
            AP.DGen.Output = False
            Sleep (2000)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (2000)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
            
                Sleep (2000)
                AP.BarGraph.Reset (3)
                Sleep (2000)
                AP.PSIA.Tx.BitClk.Dir = 0 ' turn on BCLK
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (2000)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.Tx.BitClk.Dir = 1  ' turn off BCLK
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = temp
        ActiveCell.Offset(2, 0).Select
    
    Sleep (2000)
    
End Sub

Sub CLICK_LRCLKRemoval()
' Measure the LRCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    c = 0 ' C = 0 is CMON off
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        
        For a = 0 To 1 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (2000)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.Tx.FrameClk.Dir = 0 ' turn on LRCLK
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
            Sleep (2000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -60 Then
                
            Else
                'Did not pass correct signal path test
                Cells(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.Tx.FrameClk.Dir = 1
            AP.DGen.Output = False
            Sleep (2000)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (2000)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
            
                Sleep (2000)
                AP.BarGraph.Reset (3)
                Sleep (2000)
                AP.PSIA.Tx.FrameClk.Dir = 0 ' turn on Frame clk
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (2000)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.Tx.FrameClk.Dir = 1  ' turn off Frame clk
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = temp
        ActiveCell.Offset(2, 0).Select
    
    Sleep (2000)
    AP.PSIA.Tx.FrameClk.Dir = 0 ' turn on Frame clk
    
End Sub

Sub CLICK_MCLKRemoval()
' Measure the LRCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
    
'    AP.PSIA.Tx.FrameClk.Dir = 1
'    AP.PSIA.Tx.BitClk.Dir = 1
'    AP.PSIA.MasterClkDir = 0
'    AP.PSIA.MasterClkDir = 1
'    AP.PSIA.Tx.BitClk.Dir = 0
'    AP.PSIA.Tx.FrameClk.Dir = 0
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    'C = 1 ' C = 0 is CMON off
    For c = 1 To 0 Step -1
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, CMONhi, CMONlo, c)
        
        For a = 0 To 1 'amp1, 2
            AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
            Sleep (2000)
            'Check for correct path
            AP.DGen.Output = True
            AP.PSIA.MasterClkDir = 1 ' turn on MCLK
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
            Sleep (2000)
            Var = AP.Anlr.FuncRdg("dBV")
            
            If Var > -60 Then
                
            Else
                'Did not pass correct signal path test
                Cells.Range(1, 1).value = "Fail basic signal path test"
            End If
            
            AP.PSIA.MasterClkDir = 0
            AP.DGen.Output = False
            Sleep (2000)
            
            For i = 1 To TRIALS
            
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (2000)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
            
                Sleep (2000)
                AP.BarGraph.Reset (3)
                Sleep (2000)
                AP.PSIA.MasterClkDir = 1 ' turn on MCLK
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
                Sleep (2000)
                Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
                Sleep (2000)
                AP.DGen.Output = True
                AP.DGen.Output = False
                Sleep (2000)
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(1, 0).Select
                Sleep (2000)
                AP.PSIA.MasterClkDir = 0  ' turn off MCLK
                Sleep (2000)
                Result = AP.BarGraph.max(3)
                ActiveCell.value = Result
                AP.BarGraph.Reset (3)
                
                ActiveCell.Offset(-1, 1).Select
                                
            Next i
        Next a
        ActiveCell = Cells.Range("C36")
        ActiveCell.Select
    Next c
    Sleep (2000)
    AP.PSIA.MasterClkDir = 1 ' turn on MCLK
    
End Sub

Sub CLICK_AMP_EN_AX80()
' Measure the AMP EN click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H74
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
    SPKhi = &H0
    SPKlo = &H3A
    SPKvalOn = &H1
    SPKvalOff = &H0
    ENhi = &H1
    ENlo = &HFF
    Sleep (200)
    
    a = 0
    AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
    Sleep (2000)
    'Check for correct path
    AP.DGen.Output = True
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SPKhi, SPKlo, SPKvalOn) ' Enable amp 1
    Call ClearFlags_AX80
    
    Sleep (2000)
    Var = AP.Anlr.FuncRdg("dBV")
    
    If Var > -60 Then
        
    Else
        'Did not pass correct signal path test
        Cells.Range(1, 1).value = "Fail basic signal path test"
    End If
    
    Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SPKhi, SPKlo, SPKvalOff) ' disable amp 1
    AP.DGen.Output = False
    Sleep (2000)
    
    For i = 1 To TRIALS
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
        Sleep (2000)
        Call ClearFlags_AX80
        Sleep (2000)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H1)
        Sleep (2000)
        AP.DGen.Output = True
        AP.DGen.Output = False
        Sleep (2000)
        AP.BarGraph.Reset (1)
    
        Sleep (2000)
        AP.BarGraph.Reset (1)
        Sleep (2000)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SPKhi, SPKlo, SPKvalOn) ' Enable amp 1
        Sleep (2000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
        Sleep (2000)
        Call ClearFlags_AX80 ' Clear flags
        Sleep (2000)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, ENhi, ENlo, &H0)
        Sleep (2000)
        AP.DGen.Output = True
        AP.DGen.Output = False
        Sleep (2000)
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, SPKhi, SPKlo, SPKvalOff) ' Disable amp 1
        Sleep (2000)
        Result = AP.BarGraph.max(1)
        ActiveCell.value = Result
        AP.BarGraph.Reset (1)
        
        ActiveCell.Offset(-1, 1).Select
    Next i
    'ActiveCell = Cells.Range("C36")
    'ActiveCell.Select

    Sleep (2000)
    
End Sub

Sub CLICK_AMP1EN()
' Measure the LRCLK click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    
    a = 0
    AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
    Sleep (2000)
    'Check for correct path
    AP.DGen.Output = True
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H40, &H1) ' Enable amp 1
    
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
    Sleep (2000)
    Var = AP.Anlr.FuncRdg("dBV")
    
    If Var > -60 Then
        
    Else
        'Did not pass correct signal path test
        Cells.Range(1, 1).value = "Fail basic signal path test"
    End If
    
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H40, &H0) ' disable amp 1
    AP.DGen.Output = False
    Sleep (2000)
    
    For i = 1 To TRIALS
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
        Sleep (2000)
        AP.DGen.Output = True
        AP.DGen.Output = False
        Sleep (2000)
        AP.BarGraph.Reset (3)
    
        Sleep (2000)
        AP.BarGraph.Reset (3)
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H40, &H1) ' Enable amp 1
        Sleep (2000)
        Result = AP.BarGraph.max(3)
        ActiveCell.value = Result
        AP.BarGraph.Reset (3)
        
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
        Sleep (2000)
        AP.DGen.Output = True
        AP.DGen.Output = False
        Sleep (2000)
        AP.BarGraph.Reset (3)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H40, &H0) ' Disable amp 1
        Sleep (2000)
        Result = AP.BarGraph.max(3)
        ActiveCell.value = Result
        AP.BarGraph.Reset (3)
        
        ActiveCell.Offset(-1, 1).Select
    Next i
    ActiveCell = Cells.Range("C36")
    ActiveCell.Select

    Sleep (2000)
    
End Sub
    
Sub CLICK_AMP2EN()
' Measure the AMP2EN click/pop
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    
    a = 1
    AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
    Sleep (2000)
    'Check for correct path
    AP.DGen.Output = True
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H47, &H1) ' Enable amp 1
    
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
    Sleep (2000)
    Var = AP.Anlr.FuncRdg("dBV")
    
    If Var > -60 Then
        
    Else
        'Did not pass correct signal path test
        Cells.Range(1, 1).value = "Fail basic signal path test"
    End If
    
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H47, &H0) ' disable amp 1
    AP.DGen.Output = False
    Sleep (2000)
    
    For i = 1 To TRIALS
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
        Sleep (2000)
        AP.DGen.Output = True
        AP.DGen.Output = False
        Sleep (2000)
        AP.BarGraph.Reset (3)
    
        Sleep (2000)
        AP.BarGraph.Reset (3)
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H47, &H1) ' Enable amp 1
        Sleep (2000)
        Result = AP.BarGraph.max(3)
        ActiveCell.value = Result
        AP.BarGraph.Reset (3)
        
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
        Sleep (2000)
        AP.DGen.Output = True
        AP.DGen.Output = False
        Sleep (2000)
        AP.BarGraph.Reset (3)
        
        ActiveCell.Offset(1, 0).Select
        Sleep (2000)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H47, &H0) ' Disable amp 1
        Sleep (2000)
        Result = AP.BarGraph.max(3)
        ActiveCell.value = Result
        AP.BarGraph.Reset (3)
        
        ActiveCell.Offset(-1, 1).Select
    Next i
    ActiveCell = Cells.Range("C36")
    ActiveCell.Select

    Sleep (2000)
    
End Sub


Sub CLICK_BSTEN()
' Measure the BSTEN click/pop (off only)
' Place cursor in upper left cell of 2xTRIALS data dump
' Not sure if all sleeps are necessary, but some are necessary to prevent false early readbacks ~ -100dB from the AP.BarGraph.Max
    
    PSU_GPIB = "GPIB::5"
    TRIALS = 10
    DEV = &H62
    CMONhi = &H0
    CMONlo = &H1D
    temp = ActiveCell
   
    I2C_Controls_.I2C_Connect
    Sleep (200)
    
    For a = 0 To 1
        AP.Anlr.FuncInput = a ' Select A or B on analyzer bar graph
        Sleep (2000)
        'Check for correct path
        AP.DGen.Output = True
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H60, &H1) ' Enable Boost
        
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
        Sleep (2000)
        Var = AP.Anlr.FuncRdg("dBV")
        
        If Var > -60 Then
            
        Else
            'Did not pass correct signal path test
            Cells.Range(1, 1).value = "Fail basic signal path test"
        End If
        
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H60, &H0) ' Disable Boost
        AP.DGen.Output = False
        Sleep (2000)
        
        For i = 1 To TRIALS
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
            Sleep (2000)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
            Sleep (2000)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
            Sleep (2000)
            AP.DGen.Output = True
            AP.DGen.Output = False
            Sleep (2000)
            AP.BarGraph.Reset (3)
        
            Sleep (2000)
            AP.BarGraph.Reset (3)
            Sleep (2000)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H60, &H1) ' Enable Boost
            Sleep (2000)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H0)
            Sleep (2000)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H18, &HFF)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H19, &HFF) ' Clear flags
            Sleep (2000)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H80, &H1)
            Sleep (2000)
            AP.DGen.Output = True
            AP.DGen.Output = False
            Sleep (2000)
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(1, 0).Select
            Sleep (2000)
            Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H60, &H0) ' Disable Boost
            Sleep (2000)
            Result = AP.BarGraph.max(3)
            ActiveCell.value = Result
            AP.BarGraph.Reset (3)
            
            ActiveCell.Offset(-1, 1).Select
        Next i
    Next a

    Sleep (2000)
    
End Sub

