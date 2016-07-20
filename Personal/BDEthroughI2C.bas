Attribute VB_Name = "BDEthroughI2C"
Sub BDEthroughI2C_HZ03()
    'Check if the BDE level can be read back through I2C correctly
    'Hz03
    
    'Setup requires VBAT on GPIB controlled PSU, I2C controlled through Bill's I2CBridge
    
    
    'Sweep VBAT and measure BDE level and ADC readback value, both over I2C
    
    VBATstart = 15 ' sweep starts at 15V
    VBATend = 6 ' sweet ends at 6V
    VBATstepSize = -0.1
    
    VBATGPIB = "GPIB::01"
    VBATRailName = "P25V"
    
    devaddr = &H62
    bdelev_hiaddr = &H20
    bdelev_loaddr = &HB6
    adcVal_hiaddr = &H20
    adcVal_loaddr = &H54
    Dim readback As Byte
    
    STARTROW = 37
    VBATCol = 1
    BDELevelCol = 2
    ADCReadbackCol = 4
    
    r = STARTROW
    For VBAT = VBATstart To VBATend Step VBATstepSize
        DoEvents
        Call Equipment_GPIB.Power_Supply_E3631A_.Supply_Set_Output(VBATGPIB, VBATRailName, VBAT)
        Sleep (200) ' give voltage and part time to settle
        If (Equipment_I2C.I2C_Controls_.I2C_bridge_16Bit_Read_Control(devaddr, bdelev_hiaddr, bdelev_loaddr, readback)) Then
        Else
            MsgBox ("Error with I2C Read, please fix and restart")
        End If
        
        bdeLev = readback
        Call Equipment_I2C.I2C_Controls_.I2C_bridge_16Bit_Read_Control(devaddr, adcVal_hiaddr, adcVal_loaddr, readback)
        adcVal = readback
        
        Cells(r, VBATCol).value = VBAT
        Cells(r, BDELevelCol).value = bdeLev
        Cells(r, VBATCol + 2).value = VBAT
        Cells(r, ADCReadbackCol).value = adcVal
        
        r = r + 1
    Next VBAT
    
    Call BDEthroughI2C_FormatSpreadsheet
    Call PlotTOC
    
End Sub

Private Sub BDEthroughI2C_FormatSpreadsheet()
    'Format the spreadsheet to Chris's TOC Macros style
    Cells(35, 1).value = "BDE Level"
    Cells(35, 3).value = "ADC Readback Value"
    Cells(33, 2).value = "BDE Level and ADC Readback vs PVDD Voltage"
    Cells(33, 7).value = "PVDD VOLTAGE (V)"
    Cells(33, 11).value = "BDE Level or ADC Readback Value"
End Sub

Private Sub PlotTOC()
    'awesome way to run macro
    CommandBars("TOC Macros").Controls("Plot").Execute

End Sub
