Attribute VB_Name = "AX80_VBATUVLO"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub VBATUVLO__ENLO_VBATOK_AX80()
    'Part is in shutdown and VBATOKAY signal is sent to the ATB1 pin for monitoring

    VBAT_MAX = 2.35
    VBAT_LowerUVLOMAX = 2.2 ' Value above what is UVLO level. Speeds up search by skipping to this point after crossing hysteresis point by using VBAT_MAX
    VBAT_MIN = 2
    VBAT_STEP = 0.001
    Threshold = 0.5 ' if VBAT is below 0.5, this counts as a low signal (UVLO), if above, counts as high (not yet UVLO)
    ITERATIONS = 10
    
    GPIB_PSU = "GPIB::01" 'VBAT
    GPIB_VOLT = "GPIB::12" 'Monitoring ATB1
    
    Dim inVoltage As Double
    Dim outVoltage As Double: outVoltage = VBAT_MAX
    STEPS = (VBAT_MAX - VBAT_MIN) / VBAT_STEP + 1
    
    Cells(i + 1, 1).value = "Trial #"
    Cells(i + 1, 2).value = "VBATOK Signal Collapse Point"
    Cells(i + 1, 3).value = "VBAT Signal Collapse Point"
    Cells(i + 1, 4).value = "VBATOK Recovery Point"
    Cells(i + 1, 5).value = "VBAT Recovery Point"
    
    Cells(1, 7).value = "VBATOK Voltage"
    Cells(1, 8).value = "VBAT Voltage"
    
    For i = 1 To ITERATIONS
    
        Call Power_Supply_E3631A_.Supply_Set_Output(GPIB_PSU, "P6V", VBAT_MAX)
        Sleep (1000)
        outVoltage = VBAT_LowerUVLOMAX
        Call Power_Supply_E3631A_.Supply_Set_Output(GPIB_PSU, "P6V", VBAT_LowerUVLOMAX)
        For j = 1 To STEPS ' Decreasing VBAT test
            DoEvents
            Sleep (100)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB_PSU, "P6V", outVoltage)
            'Reset device to correct ATB1 <-> VBATOKAY mode
            Call AX80General.ATB1_VBATOK
            Sleep (500)
            Call DMM_34401A_.DMM_Get_Reading(GPIB_VOLT, inVoltage)
            Sleep (100)
            Cells(j + 1, 5 + i * 2).value = inVoltage
            Cells(j + 1, 6 + i * 2).value = outVoltage
            
            If (inVoltage > Threshold) Then
                'no UVLO yet
            Else
                'output has collapsed, record shutoff point
                Cells(i + 1, 1).value = i
                Cells(i + 1, 2).value = inVoltage
                Cells(i + 1, 3).value = outVoltage
                Exit For
            End If
            
            outVoltage = VBAT_LowerUVLOMAX - j * VBAT_STEP
        Next j
        
        Sleep (1000)
        For j = 1 To STEPS
            DoEvents
            Sleep (100)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB_PSU, "P6V", outVoltage)
            Sleep (500)
            Call AX80General.ATB1_VBATOK
            Sleep (500)
            Call DMM_34401A_.DMM_Get_Reading(GPIB_VOLT, inVoltage)
            If (inVoltage > Threshold) Then
                'Signal has returned, record recovery point
                
                Cells(i + 1, 4).value = inVoltage
                Cells(i + 1, 5).value = outVoltage
                Exit For
            Else
                
            End If
            outVoltage = outVoltage + j * VBAT_STEP
            
            If outVoltage > VBAT_MAX Then
                Cells(i + 1, 3).value = "Hasn't Recovered yet"
                Exit For
            End If
            
        Next j
    Next i
    
End Sub

Sub VBATUVLO_AX80()
    'Setup the path with a very small signal and THDN monitoring on AP (~-57dB THDN should be good)

    VBAT_MAX = 2.25
    VBAT_MIN = 2
    VBAT_STEP = 0.01
    THDN_THRESHOLD = -50 ' dB
    THDN_THRESHOLD_TOOSMALL = -100
    SMALL_SIGNAL_LEVEL = -50 ' dB
    
    GPIB_PSU = "GPIB::01"
    GPIB_VOLT = "GPIB::12"
    
    ITERATIONS = 100 ' Times to run test
    
    Dim inVoltage As Double
    Dim outVoltage As Double: outVoltage = VBAT_MAX
    STEPS = (VBAT_MAX - VBAT_MIN) / VBAT_STEP + 1
    
    Cells(i + 1, 1).value = "Trial #"
    Cells(i + 1, 2).value = "Signal Collapse Point"
    Cells(i + 1, 3).value = "Recovery Point"
    For i = 1 To ITERATIONS
        For j = 1 To STEPS ' Decreasing VBAT test
            DoEvents
            Sleep (100)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB_PSU, "P6V", outVoltage)
            Sleep (500)
            Call DMM_34401A_.DMM_Get_Reading(GPIB_VOLT, inVoltage)
            Sleep (100)
            Var = AP.Anlr.FuncRdg("dB")
            If (Var < THDN_THRESHOLD And Var > THDN_THRESHOLD_TOOSMALL) Then
                'Signal is good
            Else
                'output has collapsed, record shutoff point
                Cells(i + 1, 1).value = i
                Cells(i + 1, 2).value = inVoltage
                Exit For
            End If
            
            outVoltage = VBAT_MAX - j * VBAT_STEP
        Next j
        
        Sleep (1000)
        For j = 1 To STEPS
            DoEvents
            Sleep (100)
            Call Power_Supply_E3631A_.Supply_Set_Output(GPIB_PSU, "P6V", outVoltage)
            Sleep (100)
            Call DMM_34401A_.DMM_Get_Reading(GPIB_VOLT, inVoltage)
            Var = AP.Anlr.FuncRdg("dB")
            If (Var < THDN_THRESHOLD And Var > THDN_THRESHOLD_TOOSMALL) Then
                'Signal has returned, record recovery point
                
                Cells(i + 1, 3).value = inVoltage
                Exit For
            Else
                
            End If
            outVoltage = outVoltage + j * VBAT_STEP
            
            If outVoltage > VBAT_MAX Then
                Cells(i + 1, 3).value = "Hasn't Recovered yet"
                Exit For
            End If
            
        Next j
    Next i
    
End Sub
