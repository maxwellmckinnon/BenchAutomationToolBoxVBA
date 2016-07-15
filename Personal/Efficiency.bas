Attribute VB_Name = "Efficiency"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub EfficiencyOverPVDD()
    VBAT_PVDD_GPIB = "GPIB::01" 'PSU Agilent 3631A on VBAT and PVDD PSU
    PVDD_NAME = "P25V"
    BOARD_NAME = "465A"
    
    PVDDs = Array(6.5, 8, 8.5, 10)
    For Each PVDD In PVDDs
        Worksheets.Add
        ActiveSheet.Name = BOARD_NAME & " PVDD = " & CStr(PVDD)
        Call GPIB.Power_Supply_E3631A_.Supply_Set_Output(VBAT_PVDD_GPIB, PVDD_NAME, PVDD)
        Sleep (500)
        
        Call ClassDEfficiency_AX80
        
    Next PVDD
    
End Sub

Sub ClassDEfficiency_AX80()
    'sweep the input digital signal while monitoring the input current and output voltage
    'just the efficiency of VBAT - ignores DVDD
    
    '########################
    'Automation Constants
    '########################
    VBATCURR_GPIB = "GPIB::12" ' Fluke 8845A on VBAT current
    PVDDVOLT_GPIB = "GPIB::11" ' Fluke 8845A on PVDD voltage
    PVDDCURR_GPIB = "GPIB::10" 'DMM Agilent 33401 on PVDD current
    VBAT_PVDD_GPIB = "GPIB::01" 'PSU Agilent 3631A on VBAT and PVDD PSU
    
    STEPS = 100
    STARTINPUT = -60 ' dBFS
    STOPINPUT = 0 ' dBFS
    RESISTANCE = 8.17 ' ohms
    BOARD = "412A"
    
    DEV = &H74
    
    OUTPUTDATACELL_x = 2
    OUTPUTDATACELL_Y = 1
    
    '########################
    'Automation
    '########################

    Dim PVDD_I As Double
    Dim VBAT_V As Double
    X = OUTPUTDATACELL_x
    ep = 0.001 ' Prevent AP error from setting output to +0.0000
    For i = 1 To STEPS
        DoEvents
        inputLevel = (STOPINPUT - STARTINPUT) / (STEPS - 1) * (i - 1) + STARTINPUT
        If Abs(inputLevel) < ep Then inputLevel = 0
        AP.DGen.ChAAmpl("dBFS") = inputLevel
        Sleep (1500)
        outputLevel = AP.Anlr.FuncRdg("V")
        Call DMM_34401A_.DMM_Get_Reading(PVDDCURR_GPIB, PVDD_I)
        PVDD_V = GPIB.Fluke_Meter.ReadVoltage_Fluke(PVDDVOLT_GPIB)
        Call GPIB.Power_Supply_E3631A_.Supply_Measure_Voltage(VBAT_PVDD_GPIB, "P6V", VBAT_Vset, VBAT_V)
        VBAT_I = GPIB.Fluke_Meter.ReadCurrent_Fluke(VBATCURR_GPIB)
        
        Cells(X, 1).value = inputLevel
        Cells(X, 2).value = outputLevel
        Cells(X, 5).value = PVDD_V
        Cells(X, 6).value = PVDD_I
        Cells(X, 8).value = VBAT_V
        Cells(X, 9).value = VBAT_I
        
        X = X + 1
    Next i
End Sub

Sub Efficiency_AX80()
    'sweep the input digital signal while monitoring the input current and output voltage
    'just the efficiency of VBAT - ignores DVDD
    
    CURR_GPIB = "GPIB::13" ' Fluke 8854A
    VOLT_GPIB = "GPIB::12" ' Fluke 8854A
    
    STEPS = 60
    STARTINPUT = -40 ' dBFS
    STOPINPUT = 0 ' dBFS
    RESISTANCE = 8.17 ' ohms
    BOARD = "412A"
    
    DEV = &H74
    
    OUTPUTDATACELL_x = 2
    OUTPUTDATACELL_Y = 2
    
    Dim inVoltage As Double
    Dim inCurrent As Double
    
    X = OUTPUTDATACELL_x
    y = OUTPUTDATACELL_Y + cf * 3

    For i = 1 To STEPS
        DoEvents
        inputLevel = (STOPINPUT - STARTINPUT) / (STEPS - 1) * (i - 1) + STARTINPUT
        If inputLevel > -0.1 Then inputLevel = 0
        AP.DGen.ChAAmpl("dBFS") = inputLevel
        Sleep (1500)
        outputLevel = AP.Anlr.FuncRdg("V")
        'inCurrent = Fluke_Meter.ReadCurrent_Fluke(CURR_GPIB) ' Not tested very well
        Misc.Sleep (1000)
        Call DMM_Get_Reading(CURR_GPIB, inCurrent, 5)
        Misc.Sleep (1000)
        Call DMM_Get_Reading(VOLT_GPIB, inVoltage)
        
        Cells(X, y).value = inputLevel
        Cells(X, y + 1).value = inVoltage
        'Cells(1, 1).value = "test"
        Cells(X, y + 2).value = inCurrent
        Cells(X, y + 3).value = outputLevel
        
        X = X + 1
    Next i
End Sub

Sub eff_over_freq()
    
    Dim freqs(2) As Double: freqs(0) = 300: freqs(1) = 500: freqs(2) = 2000
    
    For Each Freq In freqs
        Sheets("403A Efficiency vs fq 1k").Copy Before:=Sheets(1)
        Sheets(1).Select
        ActiveSheet.Name = "403A Efficiency vs fq " & Str(Freq)
        
        
        AP.DGen.Freq("Hz") = Freq
        Call Efficiency_AX80
    Next

End Sub

