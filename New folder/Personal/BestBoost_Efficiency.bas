Attribute VB_Name = "BestBoost_Efficiency"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If



Sub Efficiency_vs_BoostSettings_AX80()
    'sweep the input digital signal while monitoring the input current and output voltage
    
    CURR_GPIB = "GPIB::11" ' Fluke 8854A
    VOLT_GPIB = "GPIB::12" ' Fluke 8854A
    
    STEPS = CInt(Cells(1, 21).value)
    STARTINPUT = CDec(Cells(2, 21).value)
    STOPINPUT = CDec(Cells(3, 21).value)
    RESISTANCE = CDec(Cells(4, 21).value)
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
    
    For cf = 0 To 16
        X = OUTPUTDATACELL_x
        y = OUTPUTDATACELL_Y + cf * 3
        Cells(35, 1 + cf * 2).value = "cf = " & CStr(cf)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, DT_REG, cf)
        
        For i = 1 To STEPS
            
            DoEvents
            inputLevel = (STOPINPUT - STARTINPUT) / (STEPS - 1) * (i - 1) + STARTINPUT
            AP.DGen.ChAAmpl("dBFS") = inputLevel
            Sleep (1500)
            outputLevel = AP.Anlr.FuncRdg("V")
            inCurrent = Fluke_Meter.ReadCurrent_Fluke(CURR_GPIB) ' Not tested very well
            Call DMM_34401A_.DMM_Get_Reading(VOLT_GPIB, inVoltage)
            
            Cells(X, y).value = inVoltage
            Cells(1, 1).value = "test"
            Cells(X, y + 1).value = inCurrent
            Cells(X, y + 2).value = outputLevel
            X = X + 1
        Next i
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
    
    Dim SheetName As String: SheetName = "SlewRate_" & BOARD
    If Not (Misc.SheetExists(SheetName)) Then
        Sheets.Add.Name = SheetName
    End If
    ActiveWorkbook.Sheets(SheetName).Activate
    
    
    For SR_i = 0 To 8
        X = OUTPUTDATACELL_x
        y = OUTPUTDATACELL_Y + SR_i * 3
        Cells(35, 1 + SR_i * 2).value = "SR_i = 0x" & Hex(SlewRateArray(SR_i + 1))
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H54)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 1, &HFF, &H4D)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, SR_REG, SlewRateArray(SR_i + 1))
        
        For i = 1 To STEPS
            
            DoEvents
            inputLevel = (STOPINPUT - STARTINPUT) / (STEPS - 1) * (i - 1) + STARTINPUT
            AP.DGen.ChAAmpl("dBFS") = inputLevel
            Sleep (1500)
            outputLevel = AP.Anlr.FuncRdg("V")
            Sleep (1500)
            inCurrent = Fluke_Meter.ReadCurrent_Fluke(CURR_GPIB) ' Not tested very well
            'Call DMM_34401A_.DMM_Get_Reading(CURR_GPIB, inCurrent) ' Also sucks
            Call DMM_34401A_.DMM_Get_Reading(VOLT_GPIB, inVoltage)
            
            Cells(X, y).value = inVoltage
            Cells(1, 1).value = "test"
            Cells(X, y + 1).value = inCurrent
            Cells(X, y + 2).value = outputLevel
            X = X + 1
        Next i
    Next SR_i
    

End Sub

Sub CalculateRunTime_Seconds()
'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

'*****************************
'Insert Your Code Here...
'*****************************
Call Efficiency_AX80

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub

Sub QFix()
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
    For SR_i = 0 To 8
        
        Cells(35, 1 + SR_i * 2).value = "SR_i = 0x" & Hex(SlewRateArray(SR_i + 1))
    Next SR_i
End Sub

